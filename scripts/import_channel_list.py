"""
import_channel_list.py
导入「移动短信通道清单」Excel → channel_list 表
全量覆盖写入（不带有效日期，真真手工维护最新版）

用法:
    python3 scripts/import_channel_list.py <xlsx 路径>
"""
import sys
import zipfile
import re
import sqlite3
import os

def parse_xlsx(xlsx_path):
    """用纯 Python zipfile 解析 xlsx（绕开 openpyxl 的 fill 类型 bug）"""
    z = zipfile.ZipFile(xlsx_path)
    # 1. sharedStrings
    ss_xml = z.read('xl/sharedStrings.xml').decode('utf-8')
    ss = re.findall(r'<si[^>]*>(.*?)</si>', ss_xml, re.DOTALL)
    def parse_si(si_inner):
        parts = re.findall(r'<t[^>]*>([^<]*)</t>', si_inner)
        return ''.join(parts)
    shared_strings = [parse_si(s) for s in ss]
    ss_map = {i: s for i, s in enumerate(shared_strings)}

    # 2. 找 sheet1
    rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid_map = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml))
    # 假设目标 sheet 在 rId1（第一个 sheet）
    # 实际：先看 workbook.xml 里第一个 sheet 的 r:id
    wb_xml = z.read('xl/workbook.xml').decode('utf-8')
    first_sheet = re.search(r'<sheet[^/]*r:id="(rId\d+)"', wb_xml)
    rid = first_sheet.group(1) if first_sheet else 'rId1'
    target_path = f"xl/{rid_map[rid]}"

    xml = z.read(target_path).decode('utf-8')
    rows_data = re.findall(r'<row[^>]*r="(\d+)"[^>]*>(.*?)</row>', xml, re.DOTALL)

    def parse_row(row_inner, ss_map):
        cells = re.findall(r'<c\s+([^/>]*?)(?:/>|>(.*?)</c>)', row_inner, re.DOTALL)
        result = {}
        for attrs, inner in cells:
            ref_match = re.search(r'r="([A-Z]+)\d+"', attrs)
            if not ref_match: continue
            col = ref_match.group(1)
            t_match = re.search(r't="([^"]+)"', attrs)
            ctype = t_match.group(1) if t_match else 'n'
            if not inner or inner.strip() == '':
                result[col] = None; continue
            if ctype == 's':
                v = re.search(r'<v>([^<]*)</v>', inner)
                if v:
                    try:
                        result[col] = ss_map.get(int(v.group(1)), v.group(1))
                    except ValueError:
                        result[col] = v.group(1)
                else:
                    result[col] = ''
            else:
                v = re.search(r'<v>([^<]*)</v>', inner)
                result[col] = v.group(1) if v else ''
        return result

    # 3. R1 = 列头；R2+ = 数据
    out = []
    for rn, inner in rows_data[1:]:
        rdata = parse_row(inner, ss_map)
        out.append(rdata)
    return out


def import_to_db(rows, db_path='data/baobei.db'):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # 1. 清空旧数据
    cur.execute("DELETE FROM channel_list")
    print(f"已清空 channel_list")

    # 2. 插入新数据
    # 通道清单列头对应：
    # A 通道id (忽略)
    # B 来源 (忽略)
    # C 通道号码 (8位) → port_code
    # D 通道名称 → channel_name
    # E 通道分类 → channel_category
    # F 供应商简称 → region
    # G 运营商 → carrier
    # H 信息类型 → info_type
    # I 是否金融行业 → is_finance
    # J 报备人 → report_handler

    inserted = 0
    skipped = 0
    for r in rows:
        port_code = r.get('C', '').strip() if r.get('C') else ''
        if not port_code or not port_code.isdigit() or len(port_code) != 8:
            skipped += 1
            continue
        port_id = port_code[-4:]  # 后 4 位
        channel_name = r.get('D', '').strip() if r.get('D') else ''
        channel_category = r.get('E', '').strip() if r.get('E') else ''
        region = r.get('F', '').strip() if r.get('F') else ''
        carrier = r.get('G', '').strip() if r.get('G') else ''
        info_type = r.get('H', '').strip() if r.get('H') else None
        is_finance = 1 if r.get('I', '').strip() == '是' else 0 if r.get('I', '').strip() == '否' else None
        report_handler = r.get('J', '').strip() if r.get('J') else ''

        if not all([channel_name, channel_category, region, carrier, report_handler]):
            skipped += 1
            continue

        cur.execute("""
            INSERT INTO channel_list
            (port_code, port_id, channel_name, channel_category, region, carrier, info_type, is_finance, report_handler)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (port_code, port_id, channel_name, channel_category, region, carrier, info_type, is_finance, report_handler))
        inserted += 1

    conn.commit()
    conn.close()
    return inserted, skipped


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python3 scripts/import_channel_list.py <xlsx 路径>")
        sys.exit(1)
    xlsx = sys.argv[1]
    if not os.path.exists(xlsx):
        print(f"❌ 文件不存在: {xlsx}")
        sys.exit(1)

    print(f"读取 {xlsx} ...")
    rows = parse_xlsx(xlsx)
    print(f"解析到 {len(rows)} 行")

    print("导入 channel_list ...")
    inserted, skipped = import_to_db(rows)
    print(f"✅ 完成: 插入 {inserted} 条，跳过 {skipped} 条")
