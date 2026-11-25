"""
import_rejection_log.py
导入「高风险&是否可访问引流&签名.xlsx → 签名不合规」sheet → rejection_log 表
全量覆盖写入

清洗规则:
  1. 跳过空行（无签名）
  2. 签名去 \t 制表符 + strip
  3. 企业名是 #N/A 的设为 NULL
  4. 日期是 Excel 序列号（如 46139）

用法:
    python3 scripts/import_rejection_log.py <xlsx 路径>
"""
import sys
import os
import zipfile
import re
import sqlite3


def parse_xlsx(xlsx_path, sheet_rid='rId3'):
    """复用 import_channel_list 的解析逻辑"""
    z = zipfile.ZipFile(xlsx_path)
    # 1. sharedStrings
    ss_xml = z.read('xl/sharedStrings.xml').decode('utf-8')
    ss = re.findall(r'<si[^>]*>(.*?)</si>', ss_xml, re.DOTALL)
    def parse_si(si_inner):
        parts = re.findall(r'<t[^>]*>([^<]*)</t>', si_inner)
        return ''.join(parts)
    shared_strings = [parse_si(s) for s in ss]
    ss_map = {i: s for i, s in enumerate(shared_strings)}

    # 2. 找目标 sheet
    rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid_map = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml))
    target_path = f"xl/{rid_map[sheet_rid]}"

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

    return [parse_row(inner, ss_map) for _, inner in rows_data[1:]]


def clean_str(v):
    """清洗字符串：去 \t、strip、#N/A → None"""
    if v is None: return None
    s = str(v).replace('\t', '').strip()
    if s == '' or s == '#N/A': return None
    return s


def import_to_db(rows, db_path):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("DELETE FROM rejection_log")

    inserted = 0
    skipped_empty = 0
    skipped_no_signature = 0

    for r in rows:
        sig = clean_str(r.get('B'))
        if not sig:
            skipped_empty += 1
            continue
        company = clean_str(r.get('A'))
        reason = clean_str(r.get('C'))
        region = clean_str(r.get('D'))
        carrier = clean_str(r.get('E'))
        date_raw = r.get('F')
        remark = clean_str(r.get('G'))

        # 必填字段检查
        if not (reason and carrier):
            skipped_no_signature += 1
            continue

        # 日期：Excel 序列号（整数）
        date_val = None
        if date_raw and str(date_raw).replace('.', '').isdigit():
            try:
                date_val = int(float(date_raw))
            except (ValueError, TypeError):
                pass

        cur.execute("""
            INSERT INTO rejection_log
            (company_name, signature, rejection_reason, region, carrier, rejection_date, remark, source)
            VALUES (?, ?, ?, ?, ?, ?, ?, 'auto_from_excel')
        """, (company, sig, reason, region, carrier, date_val, remark))
        inserted += 1

    conn.commit()
    conn.close()
    return inserted, skipped_empty, skipped_no_signature


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python3 scripts/import_rejection_log.py <xlsx 路径>")
        sys.exit(1)
    xlsx = sys.argv[1]
    db = sys.argv[2] if len(sys.argv) > 2 else 'data/baobei.db'
    if not os.path.exists(xlsx):
        print(f"❌ 文件不存在: {xlsx}")
        sys.exit(1)

    print(f"读取 {xlsx} (签名不合规 sheet) ...")
    rows = parse_xlsx(xlsx)
    print(f"解析到 {len(rows)} 行")

    print(f"导入 rejection_log → {db} ...")
    inserted, skipped_e, skipped_ns = import_to_db(rows, db)
    print(f"✅ 完成: 插入 {inserted} 条，跳过 {skipped_e} 空行 / {skipped_ns} 无效行")
