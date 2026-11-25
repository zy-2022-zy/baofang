#!/usr/bin/env python3
"""
v0.7 导入三方/生态合作通道资源表 (channel_eco)
输入: xlsx 文件，Sheet1 列: 来源/资源类型/通道状态/通道ID/通道名/网关账号/运营商/供应商
"""
import os
import sys
import openpyxl
import sqlite3

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'baobei.db')


def main(xlsx_path):
    if not os.path.exists(xlsx_path):
        print(f'❌ 文件不存在: {xlsx_path}')
        sys.exit(1)

    print(f'导入 channel_eco ...')
    print(f'  xlsx: {xlsx_path}')

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[wb.sheetnames[0]]  # Sheet1

    rows = []
    for r in range(2, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, 9)]
        # 来源/资源类型/通道状态/通道ID/通道名/网关账号/运营商/供应商
        if not row_vals[0] or not row_vals[3]:
            continue
        rows.append({
            'source': str(row_vals[0]).strip(),
            'resource_type': str(row_vals[1]).strip() if row_vals[1] else '',
            'channel_status': str(row_vals[2]).strip() if row_vals[2] else '',
            'channel_id': str(row_vals[3]).strip(),
            'channel_name': str(row_vals[4]).strip() if row_vals[4] else '',
            'gateway_account': str(row_vals[5]).strip() if row_vals[5] else '',
            'carrier': str(row_vals[6]).strip() if row_vals[6] else '',
            'supplier': str(row_vals[7]).strip() if row_vals[7] else ''
        })

    print(f'  解析到 {len(rows)} 行')

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    inserted = 0
    skipped = 0
    try:
        cur.execute('DELETE FROM channel_eco')
        for r in rows:
            if not r['channel_id']:
                skipped += 1
                continue
            try:
                cur.execute("""
                    INSERT INTO channel_eco (source, resource_type, channel_status, channel_id, channel_name, gateway_account, carrier, supplier)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (r['source'], r['resource_type'], r['channel_status'], r['channel_id'],
                      r['channel_name'], r['gateway_account'], r['carrier'], r['supplier']))
                inserted += 1
            except sqlite3.IntegrityError:
                skipped += 1
        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f'❌ 错误: {e}')
        sys.exit(1)
    finally:
        cur.close()
        conn.close()

    print(f'✅ 完成: 插入 {inserted} 条，跳过 {skipped} 条')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python3 import_channel_eco.py <xlsx_path>')
        sys.exit(1)
    main(sys.argv[1])
