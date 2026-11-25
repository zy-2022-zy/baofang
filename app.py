"""
需求管理系统 · Flask 后端 (v0.3)
范围: 仅报备实名（第 1 块）

启动: python3 app/app.py

v0.3 变更：
- 13 列新表单字段
- 报备主端口支持多值（顿号/全角逗号/英文逗号/空格分隔）
- 完整子端口 3 种合法值
- 自动从 channel_list 带 region/report_handler/is_finance
- 业务大类/细类动态配置 (business_config)
- admin 鉴权（role=E_admin）
- 批次列表 JOIN user → submitter_name
- 申请时间/申请人显示
"""
import os
import sys
import json
import sqlite3
import re
import shutil
import subprocess
import tempfile
from datetime import datetime, date, timedelta
from functools import wraps
from flask import Flask, request, jsonify, g, send_from_directory, send_file
from flask_cors import CORS
try:
    from weasyprint import HTML as _WeasyHTML
    _WEASYPRINT_OK = True
except Exception:
    _WEASYPRINT_OK = False
    _WeasyHTML = None

def excel_date_to_iso(v):
    """Excel 日期转 ISO 格式 (YYYY-MM-DD)。
    支持：Excel 序列号 / 'YYYY-MM-DD' / 'YYYY-MM-DD HH:MM:SS' / 各种日期字符串 / 空值
    """
    if v is None or v == '': return None
    if isinstance(v, (int, float)):
        try:
            epoch = datetime(1899, 12, 30)
            return (epoch + timedelta(days=int(v))).strftime('%Y-%m-%d')
        except:
            return None
    s = str(v).strip()
    if not s: return None
    if re.match(r'^\d{4}-\d{1,2}-\d{1,2}', s):
        return s[:10]
    for fmt in ('%Y/%m/%d', '%Y.%m.%d', '%Y\u5e74%m\u6708%d\u65e5', '%m/%d/%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except:
            continue
    return s

# ========== 路径配置 ==========
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, 'data', 'baobei.db')
SCHEMA_PATH = os.path.join(BASE_DIR, 'schema.sql')
FRONTEND_DIR = os.path.join(os.path.dirname(BASE_DIR), 'frontend')

# ========== Flask 初始化 ==========
app = Flask(__name__, static_folder=None)
CORS(app)

# ========== 前端静态托管（公网部署用） ==========
@app.route('/')
def serve_index():
    resp = send_from_directory(FRONTEND_DIR, 'index.html')
    resp.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp

@app.route('/<path:filename>')
def serve_static(filename):
    return send_from_directory(FRONTEND_DIR, filename)

# ========== DB 工具 ==========
def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db

@app.teardown_appcontext
def close_db(error):
    db = g.pop('db', None)
    if db is not None:
        db.close()

def query(sql, args=(), one=False):
    cur = get_db().execute(sql, args)
    rows = cur.fetchall()
    cur.close()
    return (rows[0] if rows else None) if one else rows

def execute(sql, args=()):
    db = get_db()
    cur = db.execute(sql, args)
    db.commit()
    last_id = cur.lastrowid
    cur.close()
    return last_id

# ========== 业务工具 ==========
def generate_batch_no():
    """v0.21.1: 用 MAX(batch_no) 解析数字部分 + 1,避免 UNIQUE 冲突
    场景: 第一次 POST 有 warnings 返回 need_confirm (未插入 batch),
    第二次强制提交重新生成时, 旧实现用 COUNT+1 拿同一序号 → UNIQUE 冲突
    """
    import re as _re
    today = datetime.now().strftime('%Y%m%d')
    prefix = f'B{today}-'
    row = query("SELECT MAX(batch_no) AS m FROM demand_batch WHERE batch_no LIKE ?", (prefix + '%',), one=True)
    max_no = (row['m'] if row and row['m'] else None)
    if max_no:
        m = _re.match(rf'^{_re.escape(prefix)}(\d+)$', max_no)
        seq = int(m.group(1)) + 1 if m else 1
    else:
        seq = 1
    return f'{prefix}{seq:06d}' 

def generate_batch_name(submitter_dept, submitter_name, signatures):
    """
    v0.4 自动批次名：{部门}-{提交人}-{签名摘要}-{时间HHMM}
    例：战略客户部-战略客户部代表-信用飞/信用宝-1142
    签名多于 2 个：信用飞/信用宝等3个
    """
    sigs = [s for s in (signatures or []) if s]
    if not sigs:
        sig_part = '匿名'
    elif len(sigs) <= 2:
        sig_part = '/'.join(sigs)
    else:
        sig_part = f'{sigs[0]}/{sigs[1]}等{len(sigs)}个'
    time_part = datetime.now().strftime('%H%M')
    return f'{submitter_dept}-{submitter_name}-{sig_part}-{time_part}'

def parse_port_field(value):
    """
    v0.3 解析端口字段（去括号）：
    - "3228" → [{'port_id': '3228'}]
    - "3228、1802" → 2 条
    - "3228，1802，4336" → 3 条
    支持分隔符：顿号、全角逗号、英文逗号、空格、换行
    """
    value = str(value).strip()
    if not value:
        return None
    normalized = re.sub(r'[，,\s\n\r]+', '、', value)
    parts = [p.strip() for p in normalized.split('、') if p.strip()]
    if not parts:
        return None
    for p in parts:
        if not re.match(r'^\d{4}$', p):
            return None
    return [{'port_id': p} for p in parts]

def is_valid_complete_sub_port(value):
    """v0.5 校验完整子端口合法值: 纯数字 / 无尾号要求(申请新通道) / 有尾号要求:数字
    v0.21.1: 兼容全角/半角括号 + 冒号前后空格
    """
    if not value:
        return False
    v = str(value).strip()
    if re.match(r'^\d+$', v):
        return True
    # 兼容全角/半角括号
    if v in ('无尾号要求(申请新通道)', '无尾号要求（申请新通道）'):
        return True
    if v in ('有尾号要求(手填任意数字字符串)', '有尾号要求（手填任意数字字符串）'):
        return True
    if re.match(r'^有尾号要求\s*[:：]\s*\d+$', v):
        return True
    return False

# v0.6 业务类型映射：channel_list.channel_category → biz_type
def map_ch_category_to_biz_type(ch_category):
    if not ch_category:
        return 'sms'
    s = str(ch_category)
    if '5G' in s or '5g' in s.lower() or '阅信' in s:
        return '5g'
    if '视频' in s or '视信' in s:
        return 'video'
    return 'sms'

BIZ_TYPE_REVERSE = {'sms': '短信', '5g': '5G', 'video': '视信'}

# ========== admin 鉴权装饰器 ==========
def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user_id = request.args.get('user_id', 0, type=int)
        if not user_id:
            body = request.get_json(silent=True) or {}
            user_id = body.get('user_id', 0)
        if not user_id:
            return jsonify({'error': '需要 user_id'}), 400
        u = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True)
        if not u or u['role'] != 'E_admin':
            return jsonify({'error': '需要 E_admin 权限'}), 403
        return fn(*args, **kwargs)
    return wrapper

# ========== API ==========

# 1. GET /api/users
@app.route('/api/users', methods=['GET'])
def list_users():
    rows = query("SELECT id, name, dept, role FROM user ORDER BY id")
    return jsonify([dict(r) for r in rows])

# 2. GET /api/users/me
@app.route('/api/users/me', methods=['GET'])
def get_current_user():
    user_id = request.args.get('user_id', 1, type=int)
    row = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True)
    if not row:
        return jsonify({'error': 'user not found'}), 404
    return jsonify(dict(row))

# 3. GET /api/channel/search
@app.route('/api/channel/search', methods=['GET'])
def channel_search():
    port_id = request.args.get('port_id', '').strip()
    carrier = request.args.get('carrier', '').strip()
    if not port_id or not carrier:
        return jsonify({'error': 'port_id and carrier required'}), 400
    rows = query(
        "SELECT port_code, port_id, channel_name, channel_category, region, carrier, info_type, is_finance, report_handler "
        "FROM channel_list WHERE port_id = ? AND carrier = ?",
        (port_id, carrier)
    )
    if not rows:
        return jsonify({'found': False, 'port_id': port_id, 'carrier': carrier, 'candidates': []}), 200
    return jsonify({
        'found': True,
        'port_id': port_id,
        'carrier': carrier,
        'candidates': [dict(r) for r in rows]
    })

# 4. GET /api/business-config
@app.route('/api/business-config', methods=['GET'])
def get_business_config():
    rows = query(
        "SELECT id, config_type, value, sort_order, enabled "
        "FROM business_config WHERE enabled = 1 ORDER BY config_type, sort_order, id"
    )
    return jsonify([dict(r) for r in rows])

# 5. POST /api/business-config
@app.route('/api/business-config', methods=['POST'])
@admin_required
def add_business_config():
    data = request.get_json() or {}
    config_type = data.get('config_type', '').strip()
    value = data.get('value', '').strip()
    if config_type not in ('category', 'sub_category'):
        return jsonify({'error': 'config_type 必须是 category 或 sub_category'}), 400
    if not value:
        return jsonify({'error': 'value 不能为空'}), 400
    try:
        new_id = execute(
            "INSERT INTO business_config (config_type, value, sort_order, enabled) VALUES (?, ?, ?, 1)",
            (config_type, value, data.get('sort_order', 99))
        )
        return jsonify({'id': new_id, 'config_type': config_type, 'value': value}), 201
    except sqlite3.IntegrityError:
        return jsonify({'error': f'已存在：{config_type} = {value}'}), 409

# 6. PATCH /api/business-config/<id>
@app.route('/api/business-config/<int:cid>', methods=['PATCH'])
@admin_required
def toggle_business_config(cid):
    data = request.get_json() or {}
    if 'enabled' not in data:
        return jsonify({'error': '需要 enabled 字段'}), 400
    enabled = 1 if data['enabled'] else 0
    execute("UPDATE business_config SET enabled = ? WHERE id = ?", (enabled, cid))
    return jsonify({'id': cid, 'enabled': bool(enabled)})

# 7. DELETE /api/business-config/<id>
@app.route('/api/business-config/<int:cid>', methods=['DELETE'])
@admin_required
def del_business_config(cid):
    execute("DELETE FROM business_config WHERE id = ?", (cid,))
    return jsonify({'id': cid, 'deleted': True})

# v0.4 通道基础信息 CRUD（admin）
@app.route('/api/channels', methods=['GET'])
def list_channels():
    keyword = request.args.get('keyword', '').strip()
    port_id = request.args.get('port_id', '').strip()
    carrier = request.args.get('carrier', '').strip()
    # v0.18: 支持 port_id + carrier 精确查（供 5G/视信查报备人）
    if port_id and carrier:
        rows = query(
            "SELECT * FROM channel_list WHERE port_id = ? AND carrier = ? LIMIT 1",
            (port_id, carrier)
        )
        return jsonify([dict(r) for r in rows])
    if keyword:
        rows = query(
            "SELECT * FROM channel_list WHERE port_code LIKE ? OR port_id LIKE ? OR channel_name LIKE ? OR report_handler LIKE ? ORDER BY port_code LIMIT 500",
            (f'%{keyword}%', f'%{keyword}%', f'%{keyword}%', f'%{keyword}%')
        )
    else:
        rows = query("SELECT * FROM channel_list ORDER BY port_code LIMIT 500")
    return jsonify([dict(r) for r in rows])

@app.route('/api/channels', methods=['POST'])
@admin_required
def add_channel():
    d = request.get_json() or {}
    required = ['port_code', 'port_id', 'channel_name', 'channel_category', 'region', 'carrier', 'report_handler']
    for k in required:
        if not d.get(k):
            return jsonify({'error': f'{k} 不能为空'}), 400
    try:
        new_id = execute("""
            INSERT INTO channel_list (port_code, port_id, channel_name, channel_category, region, carrier, info_type, is_finance, report_handler)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (d['port_code'], d['port_id'], d['channel_name'], d['channel_category'], d['region'], d['carrier'],
              d.get('info_type'), 1 if d.get('is_finance') else 0, d['report_handler']))
        return jsonify({'id': new_id}), 201
    except sqlite3.IntegrityError as e:
        return jsonify({'error': f'主键冲突：{d["port_code"]}+{d["channel_category"]}+{d.get("info_type", "")}'}), 409

@app.route('/api/channels/<int:cid>', methods=['PATCH'])
@admin_required
def update_channel(cid):
    d = request.get_json() or {}
    fields = []
    values = []
    for k in ['port_code', 'port_id', 'channel_name', 'channel_category', 'region', 'carrier', 'info_type', 'is_finance', 'report_handler']:
        if k in d:
            fields.append(f'{k} = ?')
            values.append(d[k] if k != 'is_finance' else (1 if d[k] else 0))
    if not fields:
        return jsonify({'error': '无更新字段'}), 400
    values.append(cid)
    execute(f'UPDATE channel_list SET {", ".join(fields)} WHERE id = ?', values)
    return jsonify({'id': cid, 'updated': True})

@app.route('/api/channels/<int:cid>', methods=['DELETE'])
@admin_required
def del_channel(cid):
    execute("DELETE FROM channel_list WHERE id = ?", (cid,))
    return jsonify({'id': cid, 'deleted': True})

# v0.9: 主端口 8 位补全用 — 返回 port_code_map
@app.route('/api/port-code-map', methods=['GET'])
def get_port_code_map():
    """返回 {port_id, carrier, port_code} 列表，用于前端“端口分配”工具主端口 8 位补全。"""
    rows = query("SELECT DISTINCT port_id, carrier, port_code FROM channel_list WHERE port_code IS NOT NULL AND port_code != ''")
    return jsonify([dict(r) for r in rows])

# v0.7 三方/生态通道资源 (公开读，admin 才能改)
@app.route('/api/channels-eco', methods=['GET'])
def list_channels_eco():
    keyword = request.args.get('keyword', '').strip()
    channel_id = request.args.get('channel_id', '').strip()
    if channel_id:
        rows = query("SELECT * FROM channel_eco WHERE channel_id = ?", (channel_id,))
    elif keyword:
        kw = f'%{keyword}%'
        rows = query("""SELECT * FROM channel_eco
                       WHERE channel_id LIKE ? OR channel_name LIKE ? OR supplier LIKE ? OR gateway_account LIKE ?
                       ORDER BY supplier, channel_id LIMIT 500""",
                       (kw, kw, kw, kw))
    else:
        rows = query("SELECT * FROM channel_eco ORDER BY supplier, channel_id LIMIT 500")
    return jsonify([dict(r) for r in rows])

@app.route('/api/channels-eco/<int:cid>', methods=['GET'])
@admin_required
def get_channel_eco(cid):
    row = query("SELECT * FROM channel_eco WHERE id = ?", (cid,), one=True)
    if not row:
        return jsonify({'error': '通道不存在'}), 404
    return jsonify(dict(row))

# v0.11 三方通道资源管理：增/改/删 (admin only)
@app.route('/api/channels-eco', methods=['POST'])
@admin_required
def add_channel_eco():
    d = request.get_json() or {}
    required = ['source', 'resource_type', 'channel_status', 'channel_id', 'channel_name', 'gateway_account', 'carrier', 'supplier']
    for k in required:
        if not d.get(k):
            return jsonify({'error': f'{k} 不能为空'}), 400
    try:
        new_id = execute("""
            INSERT INTO channel_eco (source, resource_type, channel_status, channel_id, channel_name, gateway_account, carrier, supplier)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (d['source'], d['resource_type'], d['channel_status'], d['channel_id'], d['channel_name'],
              d['gateway_account'], d['carrier'], d['supplier']))
        return jsonify({'id': new_id}), 201
    except sqlite3.IntegrityError as e:
        return jsonify({'error': f'通道ID「{d["channel_id"]}」已存在'}), 409

@app.route('/api/channels-eco/<int:cid>', methods=['PATCH'])
@admin_required
def update_channel_eco(cid):
    d = request.get_json() or {}
    fields = []
    values = []
    for k in ['source', 'resource_type', 'channel_status', 'channel_id', 'channel_name', 'gateway_account', 'carrier', 'supplier']:
        if k in d:
            fields.append(f'{k} = ?')
            values.append(d[k])
    if not fields:
        return jsonify({'error': '无更新字段'}), 400
    values.append(cid)
    execute(f'UPDATE channel_eco SET {", ".join(fields)} WHERE id = ?', values)
    return jsonify({'id': cid, 'updated': True})

@app.route('/api/channels-eco/<int:cid>', methods=['DELETE'])
@admin_required
def del_channel_eco(cid):
    execute("DELETE FROM channel_eco WHERE id = ?", (cid,))
    return jsonify({'id': cid, 'deleted': True})

# v0.4 驳回日志（admin 只读 + 导入）
@app.route('/api/rejections', methods=['GET'])
@admin_required
def list_rejections():
    keyword = request.args.get('keyword', '').strip()
    if keyword:
        rows = query(
            "SELECT * FROM rejection_log WHERE signature LIKE ? OR rejection_reason LIKE ? OR company_name LIKE ? OR region LIKE ? ORDER BY id DESC LIMIT 500",
            (f'%{keyword}%', f'%{keyword}%', f'%{keyword}%', f'%{keyword}%')
        )
    else:
        rows = query("SELECT * FROM rejection_log ORDER BY id DESC LIMIT 500")
    return jsonify([dict(r) for r in rows])

@app.route('/api/rejections/import', methods=['POST'])
@admin_required
def import_rejections():
    """v0.4 驳回日志全量覆盖导入。客户端解析 Excel 后传 JSON 数组过来。"""
    data = request.get_json() or {}
    rows = data.get('rows', [])
    if not rows:
        return jsonify({'error': 'rows 不能为空'}), 400
    inserted = 0
    skipped = 0
    errors = []
    db = get_db()
    cur = db.cursor()
    try:
        cur.execute("DELETE FROM rejection_log")
        for idx, r in enumerate(rows):
            sig = (r.get('signature') or '').strip()
            carrier = (r.get('carrier') or '').strip()
            reason = (r.get('rejection_reason') or '').strip()
            row_has_error = False
            if not sig:
                errors.append({'row': idx + 1, 'field': '签名', 'msg': '签名为空'})
                row_has_error = True
            if not carrier:
                errors.append({'row': idx + 1, 'field': '运营商', 'msg': '运营商为空'})
                row_has_error = True
            if not reason:
                errors.append({'row': idx + 1, 'field': '驳回原因', 'msg': '驳回原因为空'})
                row_has_error = True
            if row_has_error:
                skipped += 1
                continue
            try:
                cur.execute("""
                    INSERT INTO rejection_log (company_name, signature, rejection_reason, region, carrier, rejection_date, remark, source, reporter_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, 'admin_import', ?)
                """, (
                    (r.get('company_name') or '').strip() or None,
                    sig, reason,
                    (r.get('region') or '').strip() or None,
                    carrier,
                    excel_date_to_iso(r.get('rejection_date')),
                    (r.get('remark') or '').strip() or None,
                    request.args.get('user_id', 0, type=int) or 0
                ))
                inserted += 1
            except Exception as e:
                errors.append({'row': idx + 1, 'field': '-', 'msg': f'插入失败：{e}'})
                skipped += 1
        db.commit()
    except Exception as e:
        db.rollback()
        return jsonify({'error': str(e)}), 500
    finally:
        cur.close()
    return jsonify({'inserted': inserted, 'skipped': skipped, 'total': len(rows), 'errors': errors}), 201

# 8. POST /api/demands
@app.route('/api/demands', methods=['POST'])
def create_demands():
    payload = request.get_json() or {}
    batch_name = payload.get('batch_name')
    biz_line = payload.get('biz_line', 'realname')
    submitter_id = payload.get('submitter_id', 1)
    rows_input = payload.get('rows', [])
    force_submitted = payload.get('force_submitted', False)

    if not rows_input:
        return jsonify({'error': 'rows 不能为空'}), 400

    submitter = query("SELECT * FROM user WHERE id = ?", (submitter_id,), one=True)
    if not submitter:
        return jsonify({'error': 'submitter not found'}), 404
    submitter_dept = submitter['dept']

    batch_no = generate_batch_no()

    # v0.6 业务类型（短信/5G/视信）
    biz_type = (payload.get('biz_type') or 'sms').strip()
    if biz_type not in ('sms', '5g', 'video'):
        return jsonify({'error': f'biz_type 必须是 sms/5g/video，当前: {biz_type}'}), 400

    # v0.7 通道来源（自营/三方）
    biz_source = (payload.get('biz_source') or 'self').strip()
    if biz_source not in ('self', 'eco'):
        return jsonify({'error': f'biz_source 必须是 self/eco，当前: {biz_source}'}), 400

    cat_set = {r['value'] for r in query("SELECT value FROM business_config WHERE config_type='category' AND enabled=1")}
    sub_set = {r['value'] for r in query("SELECT value FROM business_config WHERE config_type='sub_category' AND enabled=1")}

    sub_demands = []
    warnings = []
    # v0.20 错误收集：每行错误都进 errors，最后一次性返回
    errors = []

    def add_error(row, field, msg):
        errors.append({'row': row + 1, 'field': field, 'msg': msg})

    for idx, row in enumerate(rows_input):
        # v0.7 三方/生态报备分支
        row_biz_source = (row.get('biz_source') or biz_source).strip()
        if row_biz_source == 'eco':
            # 三方报备必填: 通道ID + 签名 + 尾号要求
            eco_chid_input = (str(row.get('channel_id') or row.get('channel_eco_id') or '')).strip()
            signature = (row.get('signature') or '').strip()
            complete_sub_port = (row.get('complete_sub_port') or '').strip()
            company_account = (row.get('company_account') or '').strip() or None
            company_name = (row.get('company_name') or '').strip() or None
            sales_owner = (row.get('sales_owner') or '').strip() or None
            _uf = row.get('is_finance')
            if _uf is True or _uf == 1 or _uf == '1' or _uf == 'true':
                is_finance_user = 1
            elif _uf is False or _uf == 0 or _uf == '0' or _uf == 'false':
                is_finance_user = 0
            else:
                is_finance_user = None
            remark = (row.get('remark') or '').strip() or None
            row_has_error = False
            if not eco_chid_input:
                add_error(idx, '通道ID', '三方报备必填通道ID'); row_has_error = True
            else:
                eco_row = query("SELECT * FROM channel_eco WHERE channel_id = ?", (eco_chid_input,), one=True)
                if not eco_row:
                    add_error(idx, '通道ID', f'通道ID "{eco_chid_input}" 在三方通道资源表里找不到'); row_has_error = True
            if not signature:
                add_error(idx, '签名', '签名不能为空'); row_has_error = True
            if not complete_sub_port:
                add_error(idx, '完整子端口', '尾号要求不能为空'); row_has_error = True
            elif not is_valid_complete_sub_port(complete_sub_port):
                add_error(idx, '完整子端口', f'尾号要求格式错误（{complete_sub_port}）'); row_has_error = True
            if row_has_error:
                continue
            sd = {
                'row_index': idx,
                'signature': signature,
                'complete_sub_port': complete_sub_port,
                'company_account': company_account,
                'company_name': company_name,
                'sales_owner': sales_owner,
                'carrier': eco_row['carrier'],
                'port_id': eco_row['gateway_account'],
                'channel_id': eco_row['channel_id'],
                'sub_port_count': 1,
                'extension_count': 0,
                'business_category': '',
                'business_sub_category': '',
                'is_finance': bool(is_finance_user) if is_finance_user is not None else False,
                'remark': remark,
                'region': eco_row['supplier'],
                'report_handler': eco_row['supplier'],
                'initial_check_warned': False,
                'initial_check_reason': None,
                'channel_eco_id': eco_row['id'],
                'channel_supplier': eco_row['supplier'],
            }
            sub_demands.append(sd)
            continue

        port_field = (row.get('port_id') or '').strip()
        carrier = (row.get('carrier') or '').strip()
        signature = (row.get('signature') or '').strip()
        complete_sub_port = (row.get('complete_sub_port') or '').strip()
        company_account = (row.get('company_account') or '').strip() or None
        company_name = (row.get('company_name') or '').strip() or None
        sales_owner = (row.get('sales_owner') or '').strip() or None
        sub_port_count = int(row.get('sub_port_count') or 1)
        extension_count = int(row.get('extension_count') or 0)
        biz_cat = (row.get('business_category') or '').strip()
        biz_sub = (row.get('business_sub_category') or '').strip()
        # 三态：True/False=用户选过了，None=用户未选
        _uf = row.get('is_finance')
        if _uf is True or _uf == 1 or _uf == '1' or _uf == 'true':
            is_finance_user = 1
        elif _uf is False or _uf == 0 or _uf == '0' or _uf == 'false':
            is_finance_user = 0
        else:
            is_finance_user = None
        remark = (row.get('remark') or '').strip() or None

        row_has_error = False
        if not port_field:
            add_error(idx, '报备主端口', '报备主端口不能为空'); row_has_error = True
        if not carrier:
            add_error(idx, '所属运营商', '所属运营商不能为空'); row_has_error = True
        if not signature:
            add_error(idx, '签名', '签名不能为空'); row_has_error = True
        if not complete_sub_port:
            add_error(idx, '完整子端口', '完整子端口不能为空'); row_has_error = True
        # v0.18: 5G/视信 必须纯数字（主端口从 csp 5-8 位自动取）
        if complete_sub_port and biz_type in ('5g', 'video') and not re.match(r'^\d+$', complete_sub_port):
            add_error(idx, '完整子端口', f'5G阅信/视信 完整子端口必须为纯数字（{complete_sub_port}）'); row_has_error = True
        if complete_sub_port and not is_valid_complete_sub_port(complete_sub_port):
            add_error(idx, '完整子端口', f'完整子端口格式错误（{complete_sub_port}），必须是：纯数字 / "无尾号要求（申请新通道）" / "有尾号要求:具体数字"'); row_has_error = True
        # v0.20 规则 1：有尾号要求 → 拓位位置必填（不再限 extCount>1）
        tail_extend_position = (row.get('tail_extend_position') or '').strip()
        if complete_sub_port.startswith('有尾号要求:') and not tail_extend_position:
            add_error(idx, '拓位要求', '完整子端口为「有尾号要求」时必须选择「尾号前」或「尾号后」'); row_has_error = True
        if biz_cat not in cat_set:
            add_error(idx, '业务大类', f'业务大类 "{biz_cat}" 不在配置中'); row_has_error = True
        if biz_sub not in sub_set:
            add_error(idx, '业务细类', f'业务细类 "{biz_sub}" 不在配置中'); row_has_error = True

        parsed = parse_port_field(port_field) if port_field else []
        if port_field and not parsed:
            add_error(idx, '报备主端口', f'报备主端口格式错误 ({port_field})，必须是 4 位数字，支持顿号/全角逗号/英文逗号/空格分隔多值'); row_has_error = True

        if row_has_error:
            continue

        # v0.10 原始多值记录：优先用前端传值（Excel/UI/3 条路径一致），否则按是否多值推断
        client_port_id_raw = (row.get('port_id_raw') or '').strip()
        if client_port_id_raw:
            port_id_raw = client_port_id_raw
        else:
            port_id_raw = port_field if len(parsed) > 1 else None

        row_sub_demands_start = len(sub_demands)
        for p in parsed:
            port_id = p['port_id']
            ch = query(
                "SELECT region, report_handler, is_finance, channel_category FROM channel_list WHERE port_id = ? AND carrier = ? LIMIT 1",
                (port_id, carrier), one=True
            )
            if not ch:
                add_error(idx, '报备主端口', f'通道清单找不到 {port_id}+{carrier}'); continue
            region = ch['region']
            report_handler = ch['report_handler']
            final_is_finance = is_finance_user if is_finance_user is not None else (1 if ch['is_finance'] else 0)
            # v0.6 校验业务类型：取该主端口+运营商下所有 channel_category，看有没有 match 当前 tab 的
            ch_categories = query(
                "SELECT DISTINCT channel_category FROM channel_list WHERE port_id = ? AND carrier = ?",
                (port_id, carrier)
            )
            ch_biz_types = {map_ch_category_to_biz_type(c['channel_category']) for c in ch_categories}
            if biz_type not in ch_biz_types:
                types_str = '、'.join(BIZ_TYPE_REVERSE[t] for t in ch_biz_types)
                cats_str = '、'.join(c['channel_category'] for c in ch_categories)
                add_error(idx, '报备主端口', f'通道 {port_id}+{carrier} 属于 {types_str}（{cats_str}），与当前业务类型 {BIZ_TYPE_REVERSE[biz_type]} 不匹配'); continue

            warn = None
            rl = query(
                "SELECT rejection_reason, region FROM rejection_log WHERE signature = ? AND carrier = ?",
                (signature, carrier)
            )
            if rl:
                reasons = [r['rejection_reason'] for r in rl if r['rejection_reason']]
                regions_str = ','.join(set(r['region'] for r in rl if r['region'])) or '任意'
                warn = f'该签名 "{signature}" 曾在 {carrier}（{regions_str}）被驳回：{reasons[0]}'

            sd = {
                'row_index': idx,
                'signature': signature,
                'complete_sub_port': complete_sub_port,
                'company_account': company_account,
                'company_name': company_name,
                'sales_owner': sales_owner,
                'carrier': carrier,
                'port_id': port_id,
                'port_id_raw': port_id_raw,
                'sub_port_count': sub_port_count,
                'extension_count': extension_count,
                'tail_extend_position': tail_extend_position or None,
                'business_category': biz_cat,
                'business_sub_category': biz_sub,
                'is_finance': bool(final_is_finance),
                'remark': remark,
                'region': region,
                'report_handler': report_handler,
                'initial_check_warned': bool(warn),
                'initial_check_reason': warn,
            }
            sub_demands.append(sd)
            if warn:
                warnings.append({'sub_index': len(sub_demands) - 1, 'reason': warn})
        # 如果这一行所有子端口都失败
        if len(sub_demands) == row_sub_demands_start and parsed:
            # 已经记录了 continue
            pass

    # v0.20：先检查 errors（多行错误统一返回）
    if errors and not sub_demands:
        return jsonify({
            'error': f'共 {len(errors)} 个字段错误',
            'errors': errors
        }), 400
    # 如果有部分子需求成功 + 有错误，仍然报错
    if errors:
        return jsonify({
            'error': f'共 {len(errors)} 个字段错误',
            'errors': errors,
            'partial_success': len(sub_demands)
        }), 400

    # v0.4 自动批次名：当 batch_name 为空时按规则生成
    if not batch_name:
        sigs = sorted({sd['signature'] for sd in sub_demands if sd.get('signature')})
        batch_name = generate_batch_name(submitter_dept, submitter['name'], sigs)

    if warnings and not force_submitted:
        return jsonify({
            'need_confirm': True,
            'batch_no_preview': batch_no,
            'sub_demands_preview': sub_demands,
            'warnings': warnings
        }), 200

    has_warning = bool(warnings)
    db = get_db()
    cur = db.cursor()
    try:
        cur.execute("""
            INSERT INTO demand_batch
            (batch_no, batch_name, biz_line, submitter_id, submitter_dept, sub_demand_count, has_warning, force_submitted, status, biz_type, biz_source, raw_payload)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'submitted', ?, ?, ?)
        """, (batch_no, batch_name, biz_line, submitter_id, submitter_dept, len(sub_demands),
              1 if has_warning else 0, 1 if force_submitted else 0, biz_type, biz_source, json.dumps(payload, ensure_ascii=False)))
        for sd in sub_demands:
            cur.execute("""
                INSERT INTO demand
                (batch_no, row_index, biz_type, biz_source, channel_eco_id, channel_supplier,
                 signature, complete_sub_port, company_account, company_name, sales_owner,
                 carrier, port_id, port_id_raw, sub_port_count, extension_count,
                 tail_extend_position, port_pre_occupy,
                 business_category, business_sub_category, is_finance, remark,
                 region, report_handler,
                 initial_check_warned, initial_check_reason, force_submitted, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (batch_no, sd['row_index'], biz_type, biz_source, sd.get('channel_eco_id'), sd.get('channel_supplier'),
                  sd['signature'], sd['complete_sub_port'], sd['company_account'], sd['company_name'], sd['sales_owner'],
                  sd['carrier'], sd['port_id'], sd.get('port_id_raw'),
                  sd['sub_port_count'], sd['extension_count'],
                  sd.get('tail_extend_position'), sd.get('port_pre_occupy'),
                  sd['business_category'], sd['business_sub_category'], 1 if sd['is_finance'] else 0, sd['remark'],
                  sd['region'], sd['report_handler'],
                  1 if sd['initial_check_warned'] else 0, sd['initial_check_reason'],
                  1 if force_submitted else 0, 'initial_checked'))

        db.commit()
    except Exception as e:
        db.rollback()
        return jsonify({'error': str(e)}), 500
    finally:
        cur.close()

    return jsonify({
        'need_confirm': False,
        'batch_no': batch_no,
        'batch_name': batch_name,
        'submitter_id': submitter_id,
        'submitter_name': submitter['name'],
        'submitter_dept': submitter_dept,
        'submit_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'sub_demand_count': len(sub_demands),
        'has_warning': has_warning,
        'sub_demands': sub_demands
    }), 201

# 8.5 v0.8 作废批次
VOIDABLE_STATUS = ('submitted', 'pending', 'initial_checked')
# v0.22: 终态 demand 行状态，这些状态不可写报备结果、不可作废
TERMINAL_STATUSES = ('cancelled', 'voided', 'placeholder')
# v0.22: 可写报备结果的状态（pending/in_progress 可以被改写成 success/failed/in_progress）
REPORT_WRITABLE_STATUSES = ('pending', 'in_progress')

@app.route('/api/batches/<batch_no>/void', methods=['POST'])
def void_batch(batch_no):
    """作废整批：状态机要求 status in (submitted, pending, initial_checked) 才可作废。
    权限：需求方本人 OR E_admin OR baobei 部门。"""
    body = request.get_json() or {}
    user_id = body.get('user_id', 0) or request.args.get('user_id', 0, type=int)
    if not user_id:
        return jsonify({'error': '需要 user_id'}), 400
    batch = query("SELECT * FROM demand_batch WHERE batch_no = ?", (batch_no,), one=True)
    if not batch:
        return jsonify({'error': '批次不存在'}), 404
    if batch['status'] not in VOIDABLE_STATUS:
        return jsonify({'error': f'当前状态「{batch["status"]}」不允许作废（只能在 pending/initial_checked 作废）'}), 400
    user = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True)
    if not user:
        return jsonify({'error': '用户不存在'}), 404
    is_owner = batch['submitter_id'] == user_id
    is_admin = user['role'] == 'E_admin'
    is_baobei = user['dept'] == 'baobei'
    if not (is_owner or is_admin or is_baobei):
        return jsonify({'error': '无权作废：仅需求方本人/管理员/实名部可作废'}), 403
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    execute("UPDATE demand_batch SET status='cancelled', cancelled_at=?, cancelled_by=? WHERE batch_no=?",
            (now, user_id, batch_no))
    execute("UPDATE demand SET status='cancelled' WHERE batch_no=?", (batch_no,))
    return jsonify({
        'ok': True,
        'batch_no': batch_no,
        'status': 'cancelled',
        'cancelled_at': now,
        'cancelled_by': user_id,
        'cancelled_by_name': user['name']
    }), 200

# v0.8.2: 行级作废 — 单条 demand 单独作废
@app.route('/api/demands/<int:demand_id>/void', methods=['POST'])
def void_demand(demand_id):
    """作废单行 demand：状态机要求 demand.status in (pending/initial_checked/submitted) 才可作废。
    权限：所属批次的提交人本人 OR E_admin OR baobei 部门。"""
    body = request.get_json() or {}
    user_id = body.get('user_id', 0) or request.args.get('user_id', 0, type=int)
    if not user_id:
        return jsonify({'error': '需要 user_id'}), 400
    d = query("SELECT * FROM demand WHERE id = ?", (demand_id,), one=True)
    if not d:
        return jsonify({'error': 'demand 不存在'}), 404
    if d['status'] not in VOIDABLE_STATUS:
        return jsonify({'error': f'当前行状态「{d["status"]}」不允许作废（只能在 pending/initial_checked/submitted 作废）'}), 400
    if d['cancelled_at']:
        return jsonify({'error': '该行已作废'}), 400
    batch = query("SELECT * FROM demand_batch WHERE batch_no = ?", (d['batch_no'],), one=True)
    if not batch:
        return jsonify({'error': '所属批次不存在'}), 404
    user = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True)
    if not user:
        return jsonify({'error': '用户不存在'}), 404
    is_owner = batch['submitter_id'] == user_id
    is_admin = user['role'] == 'E_admin'
    is_baobei = user['dept'] == 'baobei'
    if not (is_owner or is_admin or is_baobei):
        return jsonify({'error': '无权作废：仅需求方本人/管理员/实名部可作废'}), 403
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    execute("UPDATE demand SET status='cancelled', cancelled_at=?, cancelled_by=? WHERE id=?",
            (now, user_id, demand_id))
    # 检查批次内是否全部已作废 — 若是则同步批次状态
    bno = d['batch_no']
    remaining = query("SELECT COUNT(*) AS c FROM demand WHERE batch_no=? AND status NOT IN ('cancelled', 'placeholder', 'voided')", (bno,), one=True)
    if remaining['c'] == 0:
        execute("UPDATE demand_batch SET status='cancelled' WHERE batch_no=? AND status NOT IN ('cancelled', 'voided')", (bno,))
    return jsonify({
        'ok': True,
        'demand_id': demand_id,
        'batch_no': bno,
        'status': 'cancelled',
        'cancelled_at': now,
        'cancelled_by': user_id,
        'cancelled_by_name': user['name'],
        'batch_auto_cancelled': remaining['c'] == 0
    }), 200

# v0.20: 批量作废选中行
@app.route('/api/demands/batch-void', methods=['POST'])
def batch_void_demands():
    body = request.get_json() or {}
    user_id = body.get('user_id', 0)
    ids = body.get('ids', [])
    if not user_id:
        return jsonify({'error': '需要 user_id'}), 400
    if not ids or not isinstance(ids, list):
        return jsonify({'error': 'ids 不能为空'}), 400
    user = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True)
    if not user:
        return jsonify({'error': '用户不存在'}), 404
    is_admin = user['role'] == 'E_admin'
    is_baobei = user['dept'] == 'baobei'
    if not (is_admin or is_baobei):
        return jsonify({'error': '权限不足：仅实名部 / 管理员可批量作废'}), 403
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    success_ids = []
    skip_ids = []
    fail_ids = []
    affected_batches = set()
    for did in ids:
        try:
            did = int(did)
        except:
            continue
        d = query("SELECT id, batch_no, status, cancelled_at FROM demand WHERE id = ?", (did,), one=True)
        if not d:
            skip_ids.append(did)
            continue
        if d['cancelled_at'] or d['status'] == 'cancelled':
            skip_ids.append(did)
            continue
        if d['status'] not in VOIDABLE_STATUS:
            fail_ids.append(did)
            continue
        try:
            execute("UPDATE demand SET status='cancelled', cancelled_at=?, cancelled_by=? WHERE id=?",
                    (now, user_id, did))
            success_ids.append(did)
            affected_batches.add(d['batch_no'])
        except Exception as e:
            fail_ids.append(did)
    # 同步批状态：全作废的批次变 cancelled
    for bno in affected_batches:
        remaining = query("SELECT COUNT(*) AS c FROM demand WHERE batch_no=? AND status NOT IN ('cancelled', 'placeholder', 'voided')", (bno,), one=True)
        if remaining['c'] == 0:
            execute("UPDATE demand_batch SET status='cancelled' WHERE batch_no=? AND status NOT IN ('cancelled', 'voided')", (bno,))
    return jsonify({
        'ok': True,
        'success_count': len(success_ids),
        'skip_count': len(skip_ids),
        'fail_count': len(fail_ids),
        'success_ids': success_ids,
        'skip_ids': skip_ids,
        'fail_ids': fail_ids
    }), 200

# v0.11 报备结果回填：受理方（实名部 / 管理员）可“处理”或“驳回”
@app.route('/api/demands/<int:demand_id>/report', methods=['POST'])
def report_demand(demand_id):
    """更新报备结果：action=start → in_progress, action=reject → failed（需填失败原因）, action=success → success
    权限：仅 baobei 部门 OR E_admin。已作废行拒绝。
    v0.8 升级：写 report_log + 触发 batch.status 联动。
    可选 body.business_time（业务时间）和 import_source（'manual' / 'tool:xxx'）。"""
    body = request.get_json() or {}
    user_id = body.get('user_id', 0) or request.args.get('user_id', 0, type=int)
    if not user_id:
        return jsonify({'error': '需要 user_id'}), 400
    action = (body.get('action') or '').strip()
    if action not in ('start', 'reject', 'success'):
        return jsonify({'error': 'action 必须是 start / reject / success'}), 400
    d = query("SELECT * FROM demand WHERE id = ?", (demand_id,), one=True)
    if not d:
        return jsonify({'error': 'demand 不存在'}), 404
    # v0.22: demand.status 为终态（cancelled/voided/placeholder）不可写
    if d['status'] in TERMINAL_STATUSES or d['cancelled_at']:
        return jsonify({'error': f'该行状态为「{d["status"]}」，不可回填报备结果'}), 400
    # v0.22: report_status 已是终态（success/failed）时，不允许再次操作（防止重复回传覆盖）
    cur_rs = d['report_status'] or 'pending'
    if cur_rs in ('success', 'failed'):
        return jsonify({'error': f'当前报备状态已是「{cur_rs}」，不可重复操作'}), 400
    user = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True)
    if not user:
        return jsonify({'error': '用户不存在'}), 404
    is_admin = user['role'] == 'E_admin'
    is_baobei = user['dept'] == 'baobei'
    if not (is_admin or is_baobei):
        return jsonify({'error': '权限不足：仅实名部 / 管理员可操作'}), 403
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    old_status = d['report_status'] or 'pending'
    business_time = (body.get('business_time') or '').strip() or now  # 业务时间，未传则用 now
    import_source = (body.get('import_source') or 'manual').strip()
    new_status = ''
    failure_reason = None
    if action == 'start':
        new_status = 'in_progress'
        execute("UPDATE demand SET report_status='in_progress', report_completed_by=?, report_completed_at=NULL WHERE id=?",
                (user['name'], demand_id))
        execute("INSERT INTO report_log (demand_id, batch_no, old_status, new_status, failure_reason, business_time, changed_by, changed_at, operator_user_id, import_source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (demand_id, d['batch_no'], old_status, new_status, None, business_time, user['name'], now, user_id, import_source))
        # 联动 batch.status
        recompute_batch_status(d['batch_no'])
        return jsonify({'ok': True, 'demand_id': demand_id, 'report_status': 'in_progress', 'report_completed_by': user['name']})
    elif action == 'reject':
        failure_reason = (body.get('failure_reason') or '').strip()
        if not failure_reason:
            return jsonify({'error': '驳回必填失败原因'}), 400
        new_status = 'failed'
        execute("UPDATE demand SET report_status='failed', report_failure_reason=?, report_completed_by=?, report_completed_at=? WHERE id=?",
                (failure_reason, user['name'], business_time, demand_id))
        execute("INSERT INTO report_log (demand_id, batch_no, old_status, new_status, failure_reason, business_time, changed_by, changed_at, operator_user_id, import_source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (demand_id, d['batch_no'], old_status, new_status, failure_reason, business_time, user['name'], now, user_id, import_source))
        recompute_batch_status(d['batch_no'])
        return jsonify({'ok': True, 'demand_id': demand_id, 'report_status': 'failed', 'report_failure_reason': failure_reason, 'report_completed_by': user['name'], 'report_completed_at': business_time})
    else:  # success
        new_status = 'success'
        execute("UPDATE demand SET report_status='success', report_failure_reason=NULL, report_completed_by=?, report_completed_at=? WHERE id=?",
                (user['name'], business_time, demand_id))
        execute("INSERT INTO report_log (demand_id, batch_no, old_status, new_status, failure_reason, business_time, changed_by, changed_at, operator_user_id, import_source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (demand_id, d['batch_no'], old_status, new_status, None, business_time, user['name'], now, user_id, import_source))
        recompute_batch_status(d['batch_no'])
        return jsonify({'ok': True, 'demand_id': demand_id, 'report_status': 'success', 'report_completed_by': user['name'], 'report_completed_at': business_time})

# v0.12 批量报备回填
@app.route('/api/demands/batch-report', methods=['POST'])
def batch_report_demands():
    """批量报备回填：body: {ids: [1,2,3], action: 'start'|'reject'|'success', failure_reason?: ''}
    权限：仅 baobei 部门 OR E_admin。
    v0.8 升级：每行写 report_log + 触发 batch.status 联动。"""
    body = request.get_json() or {}
    user_id = body.get('user_id', 0)
    ids = body.get('ids', [])
    action = (body.get('action') or '').strip()
    if not user_id:
        return jsonify({'error': '需要 user_id'}), 400
    if not ids or not isinstance(ids, list):
        return jsonify({'error': 'ids 不能为空'}), 400
    if action not in ('start', 'reject', 'success'):
        return jsonify({'error': 'action 必须是 start / reject / success'}), 400
    if action == 'reject':
        failure_reason = (body.get('failure_reason') or '').strip()
        if not failure_reason:
            return jsonify({'error': '批量驳回必填失败原因'}), 400
    user = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True)
    if not user:
        return jsonify({'error': '用户不存在'}), 404
    is_admin = user['role'] == 'E_admin'
    is_baobei = user['dept'] == 'baobei'
    if not (is_admin or is_baobei):
        return jsonify({'error': '权限不足：仅实名部 / 管理员可操作'}), 403
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    business_time = (body.get('business_time') or '').strip() or now
    import_source = (body.get('import_source') or 'manual').strip()
    success_ids = []
    skip_ids = []
    fail_ids = []
    affected_batches = set()
    for did in ids:
        try:
            did = int(did)
        except:
            continue
        d = query("SELECT id, batch_no, status, cancelled_at, report_status FROM demand WHERE id = ?", (did,), one=True)
        if not d:
            skip_ids.append(did)
            continue
        # v0.22: 终态 demand 行不可写（cancelled/voided/placeholder）
        if d['status'] in TERMINAL_STATUSES or d['cancelled_at']:
            skip_ids.append(did)
            continue
        cur_rs = d.get('report_status') or 'pending'
        # v0.22: 已终态（success/failed）不可重复操作
        if cur_rs in ('success', 'failed'):
            skip_ids.append(did)
            continue
        old_status = cur_rs
        try:
            if action == 'start':
                new_status = 'in_progress'
                execute("UPDATE demand SET report_status='in_progress', report_completed_by=?, report_completed_at=NULL WHERE id=?",
                        (user['name'], did))
            elif action == 'reject':
                new_status = 'failed'
                execute("UPDATE demand SET report_status='failed', report_failure_reason=?, report_completed_by=?, report_completed_at=? WHERE id=?",
                        (body['failure_reason'], user['name'], business_time, did))
            else:  # success
                new_status = 'success'
                execute("UPDATE demand SET report_status='success', report_failure_reason=NULL, report_completed_by=?, report_completed_at=? WHERE id=?",
                        (user['name'], business_time, did))
            execute("INSERT INTO report_log (demand_id, batch_no, old_status, new_status, failure_reason, business_time, changed_by, changed_at, operator_user_id, import_source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (did, d['batch_no'], old_status, new_status, (body['failure_reason'] if action == 'reject' else None), business_time, user['name'], now, user_id, import_source))
            success_ids.append(did)
            affected_batches.add(d['batch_no'])
        except Exception as e:
            fail_ids.append({'id': did, 'err': str(e)})
    # 联动 batch.status（去重批）
    for bno in affected_batches:
        recompute_batch_status(bno)
    return jsonify({
        'ok': True,
        'action': action,
        'success_count': len(success_ids),
        'skip_count': len(skip_ids),
        'fail_count': len(fail_ids),
        'success_ids': success_ids,
        'skip_ids': skip_ids,
        'fail_ids': fail_ids,
        'completed_by': user['name']
    })

# ========== v0.8 报备结果回传增强 ==========

# 联动状态机：重算 demand_batch.status（按子需求行的 report_status 分布）
def recompute_batch_status(batch_no):
    """根据所有子需求行的 report_status，重算整批 status。
    规则（不计入 cancelled 和 placeholder 行）：
    - 全部 cancelled（全无有效行）→ batch.status='cancelled'
    - 全部 success → reported
    - 全是 failed（无 success/in_progress/pending）→ partial_failed
    - 全部 in_progress → processing
    - 混合（含 in_progress 或 含 failed+success/pending）→ partial_processing
    - 全是 pending → 保持原状，不动
    """
    rows = query(
        "SELECT report_status, status FROM demand WHERE batch_no = ? AND (cancelled_at IS NULL AND status != 'cancelled') AND (status != 'placeholder')",
        (batch_no,)
    )
    # v0.22: 全是 cancelled（无任何有效行）→ 批次直接变 cancelled
    if not rows:
        execute("UPDATE demand_batch SET status='cancelled' WHERE batch_no=? AND status NOT IN ('cancelled', 'voided')", (batch_no,))
        return
    statuses = [r['report_status'] or 'pending' for r in rows]
    has_success = 'success' in statuses
    has_failed = 'failed' in statuses
    has_in_progress = 'in_progress' in statuses
    has_pending = 'pending' in statuses
    new_status = None
    # 全部 success
    if has_success and not has_failed and not has_in_progress and not has_pending:
        new_status = 'reported'
    # 全是 failed（无 success/in_progress/pending）→ partial_failed（不能有 pending/in_progress 否则是混合）
    elif has_failed and not has_success and not has_in_progress and not has_pending:
        new_status = 'partial_failed'
    # 全部 in_progress
    elif has_in_progress and not has_success and not has_failed and not has_pending:
        new_status = 'processing'
    # 混合（有任何 in_progress，或 failed 与其他状态共存）
    elif has_in_progress or has_failed:
        new_status = 'partial_processing'
    else:
        # 全是 pending → 不动
        return
    execute("UPDATE demand_batch SET status = ? WHERE batch_no = ?", (new_status, batch_no))


# GET /api/demands/<id>/report-log  查询某行的回传历史
@app.route('/api/demands/<int:demand_id>/report-log', methods=['GET'])
def get_demand_report_log(demand_id):
    """查询某子需求行的回传历史（按时间倒序）。权限：仅 baobei 部门 OR E_admin。"""
    user_id = request.args.get('user_id', 0, type=int)
    if not user_id:
        return jsonify({'error': '需要 user_id'}), 400
    u = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True)
    if not u:
        return jsonify({'error': '用户不存在'}), 404
    if u['dept'] != 'baobei' and u['role'] != 'E_admin':
        return jsonify({'error': '权限不足：仅实名部 / 管理员可查'}), 403
    rows = query(
        "SELECT * FROM report_log WHERE demand_id = ? ORDER BY changed_at DESC, id DESC",
        (demand_id,)
    )
    return jsonify({'ok': True, 'demand_id': demand_id, 'logs': [dict(r) for r in rows]})


# GET /api/batches/<batch_no>/report-summary  批次内状态汇总
@app.route('/api/batches/<batch_no>/report-summary', methods=['GET'])
def get_batch_report_summary(batch_no):
    """批次内报备状态汇总：pending / in_progress / success / failed 四个数字。权限：仅 baobei 部门 OR E_admin。"""
    user_id = request.args.get('user_id', 0, type=int)
    if not user_id:
        return jsonify({'error': '需要 user_id'}), 400
    u = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True)
    if not u:
        return jsonify({'error': '用户不存在'}), 404
    if u['dept'] != 'baobei' and u['role'] != 'E_admin':
        return jsonify({'error': '权限不足：仅实名部 / 管理员可查'}), 403
    # 有效行：非 cancelled、非 placeholder
    rows = query(
        "SELECT report_status, status FROM demand WHERE batch_no = ? AND (cancelled_at IS NULL AND status != 'cancelled') AND (status != 'placeholder')",
        (batch_no,)
    )
    summary = {'pending': 0, 'in_progress': 0, 'success': 0, 'failed': 0, 'total': 0}
    for r in rows:
        st = r['report_status'] or 'pending'
        if st not in summary:
            st = 'pending'
        summary[st] += 1
        summary['total'] += 1
    # 查出当前批 batch.status
    b = query("SELECT status FROM demand_batch WHERE batch_no = ?", (batch_no,), one=True)
    summary['batch_status'] = b['status'] if b else None
    return jsonify({'ok': True, 'batch_no': batch_no, 'summary': summary})


# ========== v0.9 报备结果·工具导入 ==========

# 公共函数：根据运营商子端口 + 签名 + 运营商 在 demand 表中匹配
def match_demand_for_subport(operator_subport, signature, carrier):
    """根据运营商下发的子端口号(9-20位) + 签名 + 运营商，匹配一条 demand 行。
    返回: {id, batch_no, match_type} | None
    匹配策略（三层）:
      ① 精确：demand.port_id = operator_subport（v0.14 拆出新行）
              或 demand.complete_sub_port = operator_subport（v0.6 旧行）
      ② 末位：demand.complete_sub_port 以 '有尾号要求:数字' 开头，且
              operator_subport 以 '数字' 结尾
      ③ 兜底：取 (port_id+carrier+signature) 圈出的候选中，
              port_alloc_inserted=1 的最早行
    """
    if not operator_subport or not str(operator_subport).isdigit():
        return None
    op = str(operator_subport).strip()
    if len(op) < 9 or len(op) > 20:
        return None
    # Step 1: 运营商子端口前 8 位 → 主端口 8 位
    main_port_8 = op[:8]
    # Step 2: JOIN channel_list 拿 port_id（4 位）
    ch = query(
        "SELECT port_id FROM channel_list WHERE port_code = ? AND carrier = ?",
        (main_port_8, carrier), one=True
    )
    # v0.23: 如果 channel_list 里没找到(可能是 expandOneRow fallback 生成的假 prefix, 如 10694336)
    # 用 operator_subport 第 5-8 位 作为 port_id 候选重试
    # 例: 10694336712340 → [4:8]='3367' 不是 port_id, [5:8]='367' 也不对
    # 其实最准确是 fallback 用 operator_subport[4:8] 作为 port_id(取第 5-8 位)
    # 例: 10694336712340 → '3367' 不是。试试不同位置:
    # 真真需求里 port_id 永远在子端口的 [4:8] 位(例: 10684336 = 1068 + port_id=4336)
    # 但 fallback 的话 10694336 = 1069 + 4336,  port_id 位置依然是 [4:8]='3367' 错了
    # 正确推断: fallback '1069'+port, 所以 port_id 在末尾 4 位: '3367' → 不是 4336
    # 需要: 从后 8 位中提取 port_id(假设 pos [-4:])
    # 10694336712340[12-4:12] = '3367' 还是不对
    # 因为 fallback 是 '1069'+port = '1069'+'4336' = '10694336'
    # 那 operator_subport = '10694336' + '7712340' = '106943367712340'
    # port_id 是 '4336', 位置在 [4:8] = '3367' 是错的
    # 正确位置: op[4:8] 是不对的, op[3:7]='9433', op[4:8]='4336' 如果是 [4:8] 看错了
    # 重新数: 10694336712340
    #         0123456789012
    #         1069 4336 712340
    #         pos 0-3: '1069'
    #         pos 4-7: '4336' = port_id!
    # 哦 100 % 对. op[4:8] 就是 port_id=4336.  我之前数错了
    # 检查: '10694336'[4:8] = '4336' ✓
    # 同样 '10684336'[4:8] = '4336' ✓
    # 所以 port_id 总是 op[4:8], 不管 prefix 怎么 fallback
    if not ch:
        # 尝试用 op[4:8] 作为 port_id 直接圈候选 (v0.23)
        port_id_4 = op[4:8]
    else:
        port_id_4 = ch['port_id']
    # Step 3: 圈候选
    # v0.9.4: 同一 (port_id+carrier+signature) 跨批次存在时，优先匹配最新 batch (真真 6/15 决策 3)
    # 二次排序：port_alloc_inserted=1 优先；最后 fallback 取 demand.created_at ASC（最久未处理优先）
    candidates = query(
        """SELECT d.id, d.batch_no, d.port_id, d.complete_sub_port, d.concrete_sub_port, d.port_alloc_inserted, d.created_at, d.report_status,
                  b.created_at as batch_created_at, b.status as batch_status
           FROM demand d
           LEFT JOIN demand_batch b ON d.batch_no = b.batch_no
           WHERE d.port_id = ? AND d.carrier = ? AND d.signature = ?
             AND (d.cancelled_at IS NULL AND d.status != 'cancelled')
             AND d.status != 'placeholder'
             AND (b.status IS NULL OR b.status != 'cancelled')
           ORDER BY b.created_at DESC, (d.port_alloc_inserted = 1) DESC, d.created_at ASC""",
        (port_id_4, carrier, signature)
    )
    if not candidates:
        return None
    # Step 4 ① 精确匹配
    # v0.10: 增加 concrete_sub_port 字段匹配（拓位工具同步后的具体子端口）
    for c in candidates:
        if c['port_id'] == op:
            return {'id': c['id'], 'batch_no': c['batch_no'], 'match_type': 'exact_port_id', 'candidates_count': len(candidates)}
        if (c['complete_sub_port'] or '').strip() == op:
            return {'id': c['id'], 'batch_no': c['batch_no'], 'match_type': 'exact_complete_sub_port', 'candidates_count': len(candidates)}
        if (c['concrete_sub_port'] or '').strip() == op:
            return {'id': c['id'], 'batch_no': c['batch_no'], 'match_type': 'exact_concrete_sub_port', 'candidates_count': len(candidates)}
    # Step 4 ② 末位匹配（有尾号要求:数字）
    for c in candidates:
        cs = (c['complete_sub_port'] or '').strip()
        if (cs.startswith('有尾号要求:') or cs.startswith('有尾号要求：')) and len(cs) > 5:
            tail = cs.split(':', 1)[-1].strip().rstrip('：')
            if tail and tail.isdigit() and op.endswith(tail):
                return {'id': c['id'], 'batch_no': c['batch_no'], 'match_type': 'tail_match', 'candidates_count': len(candidates)}
    # Step 4 ②b 后段匹配 (v0.23: 兼容 expandOneRow fallback 生成的假 prefix)
    # 背景: expandOneRow 查不到 lookupPort8 时曾 fallback '1069'+port (错前缀)
    # 修复后可能还会出现历史 Excel,按「后段 (8 位之后) 在 concrete_sub_port 里出现」匹配
    if len(op) > 8:
        suffix = op[8:]
        for c in candidates:
            concrete = (c['concrete_sub_port'] or '').strip()
            if concrete and len(concrete) > 8 and concrete[8:] == suffix:
                return {'id': c['id'], 'batch_no': c['batch_no'], 'match_type': 'suffix_match', 'candidates_count': len(candidates)}
    # Step 4 ③ 兜底：取 port_alloc_inserted=1 的最早行（已 ORDER BY 优先）
    return {'id': candidates[0]['id'], 'batch_no': candidates[0]['batch_no'], 'match_type': 'fallback_first_candidate', 'candidates_count': len(candidates)}


# POST /api/demands/import-report  工具批量导入（admin-only）
@app.route('/api/demands/import-report', methods=['POST'])
def import_report_demands():
    """v0.9 工具批量导入。
    body: {
        user_id: 6 (admin),
        source: 'import:xxx',  # 导入源标记
        column_mapping: {sub_port: 5, carrier: 2, signature: 6, result: 11, failure_reason: 13, business_time: 14},
        business_time_overwrite: false,  # True=用「最后更新时间」覆盖
        failure_reason_from_remark: true,  # True=失败时从「备注」列取
        rows: [
            {sub_port, carrier, signature, result, failure_reason, business_time}
        ]
    }
    返回: {ok, total, success, unmatched, skipped, affected_batches, details: [...]}
    """
    body = request.get_json() or {}
    user_id = body.get('user_id', 0)
    if not user_id:
        return jsonify({'error': '需要 user_id'}), 400
    u = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True)
    if not u or (u['dept'] != 'baobei' and u['role'] != 'E_admin'):
        return jsonify({'error': '权限不足：仅实名部 / 管理员可操作'}), 403
    rows = body.get('rows', [])
    if not rows or not isinstance(rows, list):
        return jsonify({'error': 'rows 不能为空'}), 400
    source = (body.get('source') or 'tool:import').strip()
    failure_from_remark = bool(body.get('failure_reason_from_remark', True))
    mapping = body.get('column_mapping') or {}
    # 转换结果值为枚举
    def normalize_result(v):
        s = (str(v) if v is not None else '').strip()
        if not s: return None
        if '成功' in s: return 'success'
        if '失败' in s or '驳回' in s: return 'failed'
        if '中' in s or '进行' in s: return 'in_progress'
        return None
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    success_count = 0
    unmatched_count = 0
    skipped_count = 0
    affected_batches = set()
    details = []  # [{row_no, sub_port, status: 'success'/'unmatched'/'skipped', demand_id, batch_no, match_type, error}]
    # 用于「结果冲突」去重: (sub_port, carrier, signature) → 最后一条覆盖
    # 但只要按顺序处理即可，不需 dedup
    for idx, r in enumerate(rows):
        row_no = idx + 2  # Excel 源行号（第 1 行是表头）
        try:
            sub_port = str(r.get('sub_port') or '').strip()
            carrier = str(r.get('carrier') or '').strip()
            signature = str(r.get('signature') or '').strip()
            result_raw = r.get('result')
            failure_reason = (str(r.get('failure_reason') or '').strip()) if r.get('failure_reason') is not None else ''
            business_time = (str(r.get('business_time') or '').strip()) or now
        except Exception as e:
            details.append({'row_no': row_no, 'status': 'skipped', 'error': '行数据解析失败: ' + str(e)})
            skipped_count += 1
            continue
        # 必要的字段检查
        if not sub_port or not carrier or not signature:
            details.append({'row_no': row_no, 'sub_port': sub_port, 'status': 'skipped', 'error': '子端口/运营商/签名不能为空'})
            skipped_count += 1
            continue
        result = normalize_result(result_raw)
        if not result:
            details.append({'row_no': row_no, 'sub_port': sub_port, 'status': 'skipped', 'error': '报备结果无法识别: ' + str(result_raw)})
            skipped_count += 1
            continue
        # 匹配
        m = match_demand_for_subport(sub_port, signature, carrier)
        if not m:
            details.append({'row_no': row_no, 'sub_port': sub_port, 'carrier': carrier, 'signature': signature, 'status': 'unmatched', 'reason': '主端口反推失败或无候选行'})
            unmatched_count += 1
            continue
        demand_id = m['id']
        batch_no = m['batch_no']
        d = query("SELECT id, batch_no, status, cancelled_at, report_status FROM demand WHERE id = ?", (demand_id,), one=True)
        if not d:
            details.append({'row_no': row_no, 'sub_port': sub_port, 'status': 'skipped', 'error': '匹配行已不存在'})
            skipped_count += 1
            continue
        if d['cancelled_at'] or d['status'] == 'cancelled':
            details.append({'row_no': row_no, 'sub_port': sub_port, 'demand_id': demand_id, 'status': 'skipped', 'error': '已作废行不导入'})
            skipped_count += 1
            continue
        if d['status'] == 'placeholder':
            details.append({'row_no': row_no, 'sub_port': sub_port, 'demand_id': demand_id, 'status': 'skipped', 'error': '占位行不导入'})
            skipped_count += 1
            continue
        # 前置条件：默认覆盖所有状态（运营商下发是全量快照）。勾选 protect_manual_results=1 才跳过终态。
        cur_rs = d['report_status'] or 'pending'
        if cur_rs in ('success', 'failed') and body.get('protect_manual_results', False):
            details.append({'row_no': row_no, 'sub_port': sub_port, 'demand_id': demand_id, 'status': 'skipped', 'error': f'当前 report_status={cur_rs} 已是终态，受保护跳过（取消「保护人工结果」可覆盖）'})
            skipped_count += 1
            continue
        old_status = cur_rs
        # 失败原因：默认从「备注」列取（如果勾选了 failure_reason_from_remark 且 result=failed 且 failure_reason 为空）
        if result == 'failed' and failure_from_remark and not failure_reason:
            # 失败原因应该已经来自 Excel「备注」列（由前端列映射控制）；这里是兑底
            failure_reason = (str(r.get('remark') or '').strip()) or '(未填原因)'
        # 执行写
        try:
            if result == 'success':
                execute("UPDATE demand SET report_status='success', report_failure_reason=NULL, report_completed_by=?, report_completed_at=? WHERE id=?",
                        (u['name'], business_time, demand_id))
            elif result == 'failed':
                execute("UPDATE demand SET report_status='failed', report_failure_reason=?, report_completed_by=?, report_completed_at=? WHERE id=?",
                        (failure_reason, u['name'], business_time, demand_id))
            else:  # in_progress
                execute("UPDATE demand SET report_status='in_progress', report_completed_by=?, report_completed_at=NULL WHERE id=?",
                        (u['name'], demand_id))
            # 写 log
            execute("INSERT INTO report_log (demand_id, batch_no, old_status, new_status, failure_reason, business_time, changed_by, changed_at, operator_user_id, import_source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (demand_id, batch_no, old_status, result, (failure_reason if result == 'failed' else None), business_time, u['name'], now, user_id, source))
            details.append({'row_no': row_no, 'sub_port': sub_port, 'demand_id': demand_id, 'batch_no': batch_no, 'status': 'success', 'match_type': m['match_type'], 'result': result})
            success_count += 1
            affected_batches.add(batch_no)
        except Exception as e:
            details.append({'row_no': row_no, 'sub_port': sub_port, 'demand_id': demand_id, 'status': 'skipped', 'error': '写入失败: ' + str(e)})
            skipped_count += 1
    # 联动 status
    for bno in affected_batches:
        recompute_batch_status(bno)
    return jsonify({
        'ok': True,
        'total': len(rows),
        'success': success_count,
        'unmatched': unmatched_count,
        'skipped': skipped_count,
        'affected_batches': sorted(list(affected_batches)),
        'details': details
    })


# 9. GET /api/demands
@app.route('/api/demands', methods=['GET'])
def list_demands():
    batch_no = request.args.get('batch_no')
    # v0.10: concrete_only=1 只返回 is_concrete_sub_port=1 的行（默认 0 返回全部）
    concrete_only = request.args.get('concrete_only', '0') in ('1', 'true', 'yes')
    # v0.20: alloc_only=1 只返回 port_alloc_inserted=1 的行（工具生成的新行）
    alloc_only = request.args.get('alloc_only', '0') in ('1', 'true', 'yes')
    # v0.23: full=1 不带 batch_no 时返回全量(去掉默认 LIMIT 200),供前端「全部明细」用
    full = request.args.get('full', '0') in ('1', 'true', 'yes')
    where_parts = []
    params = []
    if concrete_only:
        where_parts.append('d.is_concrete_sub_port = 1')
    if alloc_only:
        where_parts.append('d.port_alloc_inserted = 1')
    where_extra = (' AND ' + ' AND '.join(where_parts)) if where_parts else ''
    if batch_no:
        where_parts2 = ['d.batch_no = ?'] + where_parts
        where_sql = ' AND '.join(where_parts2)
        rows = query(f"""
            SELECT d.*, c.port_code AS port_code_full, c.channel_category AS ch_category
            FROM demand d
            LEFT JOIN (SELECT port_id, carrier, MIN(port_code) AS port_code, MIN(channel_category) AS channel_category FROM channel_list GROUP BY port_id, carrier) c
              ON d.port_id = c.port_id AND d.carrier = c.carrier
            WHERE {where_sql}
            ORDER BY d.row_index, d.id
        """, (batch_no,))
    else:
        where_parts2 = ['1=1'] + where_parts
        where_sql = ' AND '.join(where_parts2)
        rows = query(f"""
            SELECT d.*, c.port_code AS port_code_full, c.channel_category AS ch_category
            FROM demand d
            LEFT JOIN (SELECT port_id, carrier, MIN(port_code) AS port_code, MIN(channel_category) AS channel_category FROM channel_list GROUP BY port_id, carrier) c
              ON d.port_id = c.port_id AND d.carrier = c.carrier
            WHERE {where_sql}
            ORDER BY d.id DESC {'' if full else 'LIMIT 200'}
        """)
    # 补齐展示用 port_code（优先 8 位完整 port_code，退回 4 位 port_id）
    # v0.23: 多端口时聚合所有区域（port_id_raw 含多值时查 channel_list 补 all_regions）
    def get_all_regions(port_id_raw, carrier):
        if not port_id_raw:
            return ''
        parts = re.split(r'[、,，\s/]+', str(port_id_raw))
        parts = [p.strip() for p in parts if p.strip()]
        if not parts:
            return ''
        regions = set()
        for pid in parts:
            row = query(
                "SELECT region FROM channel_list WHERE port_id=? AND carrier=? LIMIT 1",
                (pid, carrier), one=True
            )
            if row and row['region']:
                regions.add(row['region'])
        return '、'.join(sorted(regions)) if regions else ''

    out = []
    for r in rows:
        d = dict(r)
        d['port_code_display'] = d.get('port_code_full') or d['port_id']
        # v0.23: 多端口区域聚合
        port_raw = d.get('port_id_raw') or ''
        carrier = d.get('carrier') or ''
        if port_raw and ('、' in str(port_raw) or '/' in str(port_raw) or ',' in str(port_raw)):
            d['all_regions'] = get_all_regions(port_raw, carrier)
        else:
            d['all_regions'] = d.get('region') or ''
        out.append(d)
    return jsonify(out)

# v0.10: 批次明细汇总（按 signature 聚合，不展开 port_id）
# 返回 Excel 「需求明细_汇总」格式：同签名多行合一，报备主端口用「、」拼接
# 新加 concrete_sub_ports（具体子端口列表，换行拼接）
@app.route('/api/demands/summary', methods=['GET'])
def list_demands_summary():
    batch_no = request.args.get('batch_no')
    if not batch_no:
        return jsonify({'error': '需要 batch_no'}), 400
    rows = query("""
        SELECT d.*, c.port_code AS port_code_full, c.channel_category AS ch_category
        FROM demand d
        LEFT JOIN (SELECT port_id, carrier, MIN(port_code) AS port_code, MIN(channel_category) AS channel_category FROM channel_list GROUP BY port_id, carrier) c
          ON d.port_id = c.port_id AND d.carrier = c.carrier
        WHERE d.batch_no = ?
        ORDER BY d.row_index, d.id
    """, (batch_no,))
    if not rows:
        return jsonify([])

    batch = query("SELECT * FROM demand_batch WHERE batch_no=?", (batch_no,), one=True)
    submitter_id = batch['submitter_id'] if batch else 0
    submitter = query("SELECT name, dept FROM user WHERE id=?", (submitter_id,), one=True) if submitter_id else None

    from collections import OrderedDict
    groups = OrderedDict()
    for r in rows:
        d = dict(r)
        sig = d.get('signature') or '_unknown_'
        if sig not in groups:
            groups[sig] = {
                'batch_no': batch_no,
                'batch_name': (batch['batch_name'] if batch else '') or '',
                'submitter': (submitter['name'] if submitter else '') or '',
                'dept': (submitter['dept'] if submitter else '') or '',
                'applied_at': (batch['created_at'] if batch else '') or '',
                'signature': sig,
                'complete_sub_port': '',
                'company_account': d.get('company_account') or '',
                'company_name': d.get('company_name') or '',
                'sales_owner': d.get('sales_owner') or '',
                'carrier': d.get('carrier') or '',
                'port_ids': [],
                'port_id_raw_display': '',
                'sub_port_count_total': 0,
                'extension_count': d.get('extension_count') or 0,
                'business_category': d.get('business_category') or '',
                'business_sub_category': d.get('business_sub_category') or '',
                'is_finance': d.get('is_finance') or 0,
                'tail_extend_position': d.get('tail_extend_position') or '',
                'region': d.get('region') or '',
                'report_handler': d.get('report_handler') or '',
                'initial_check_warned': d.get('initial_check_warned') or 0,
                'initial_check_reason': d.get('initial_check_reason') or '',
                'port_pre_occupy': d.get('port_pre_occupy') or '',
                'remark': d.get('remark') or '',
                'concrete_sub_ports': [],
                'report_status': 'pending',
                'report_failure_reason': '',
                'report_completed_at': '',
                'report_completed_by': '',
                'expansion_row_count': 0,
            }
        g = groups[sig]
        # csp：取首个非占位
        csp = d.get('complete_sub_port') or ''
        if not g['complete_sub_port'] and csp:
            g['complete_sub_port'] = csp
        # port_id 4 位
        pid = d.get('port_id') or ''
        if pid and pid not in g['port_ids']:
            g['port_ids'].append(pid)
        # 原多值 port_id_raw
        raw = d.get('port_id_raw') or ''
        if raw and not g['port_id_raw_display']:
            g['port_id_raw_display'] = raw
        # 子端口数 sum
        g['sub_port_count_total'] += d.get('sub_port_count') or 0
        # concrete_sub_port
        csp2 = d.get('concrete_sub_port') or ''
        if csp2 and csp2 not in g['concrete_sub_ports']:
            g['concrete_sub_ports'].append(csp2)
        # 报备状态 priority: success > failed > in_progress > pending
        rs = d.get('report_status') or 'pending'
        if d.get('port_alloc_inserted') == 1:
            g['expansion_row_count'] += 1
        else:
            g['expansion_row_count'] += 1
        if rs == 'success':
            g['report_status'] = 'success'
        elif rs == 'failed' and g['report_status'] not in ('success',):
            g['report_status'] = 'failed'
        elif rs == 'in_progress' and g['report_status'] not in ('success', 'failed'):
            g['report_status'] = 'in_progress'
        # 报备完成时间
        if d.get('report_completed_at') and not g['report_completed_at']:
            g['report_completed_at'] = d['report_completed_at']
        if d.get('report_completed_by') and not g['report_completed_by']:
            g['report_completed_by'] = d['report_completed_by']
        if d.get('report_failure_reason') and not g['report_failure_reason']:
            g['report_failure_reason'] = d['report_failure_reason']

    out = []
    for sig, g in groups.items():
        g['port_id_display'] = '、'.join(g['port_ids'])
        if not g['port_id_raw_display']:
            g['port_id_raw_display'] = g['port_id_display']
        if not g['complete_sub_port']:
            for d in rows:
                if d['signature'] == sig and d.get('complete_sub_port'):
                    g['complete_sub_port'] = d['complete_sub_port']
                    break
        # concrete_sub_ports 拼接（换行分隔，UI 可多行展示）
        g['concrete_sub_ports_str'] = '\n'.join(g['concrete_sub_ports'])
        g['concrete_count'] = len(g['concrete_sub_ports'])
        g['sub_port_count'] = g['sub_port_count_total']
        # 拆出子行数：取 is_concrete=1 的需求行（具体子端口行）
        g['expansion_sub_rows'] = sum(1 for d in rows if d['signature'] == sig and d['is_concrete_sub_port'] == 1)
        del g['port_ids']
        out.append(g)
    return jsonify(out)


# 10. GET /api/batches - JOIN user
@app.route('/api/batches', methods=['GET'])
def list_batches():
    user_id = request.args.get('user_id', 0, type=int)
    user_row = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True) if user_id else None
    if not user_row:
        return jsonify({'error': '需要 user_id'}), 400
    role = user_row['role']
    user_dept = user_row['dept']
    can_see_all = user_dept in ('customer_service', 'operations', 'baobei') or role == 'E_admin'

    if can_see_all:
        rows = query("""
            SELECT b.*, u.name AS submitter_name
            FROM demand_batch b
            JOIN user u ON b.submitter_id = u.id
            ORDER BY b.created_at DESC LIMIT 200
        """)
    else:
        rows = query("""
            SELECT b.*, u.name AS submitter_name
            FROM demand_batch b
            JOIN user u ON b.submitter_id = u.id
            WHERE b.submitter_dept = ?
            ORDER BY b.created_at DESC LIMIT 200
        """, (user_dept,))
    # v0.22: 一次查询给每批加 report_summary(已报备/报备中/待报备/失败)
    # 仅算有效行:非 cancelled、非 placeholder
    if rows:
        batch_nos = [b['batch_no'] for b in rows]
        placeholders = ','.join('?' * len(batch_nos))
        # 一条 SQL 按 batch_no + report_status 聚合
        summary_rows = query(f"""
            SELECT batch_no,
                   report_status,
                   COUNT(*) AS cnt
            FROM demand
            WHERE batch_no IN ({placeholders})
              AND cancelled_at IS NULL
              AND status != 'cancelled'
              AND status != 'placeholder'
            GROUP BY batch_no, report_status
        """, tuple(batch_nos))
        # 转 map
        summary_map = {}
        for r in summary_rows:
            bn = r['batch_no']
            st = r['report_status'] or 'pending'
            if st not in ('pending', 'in_progress', 'success', 'failed'):
                st = 'pending'
            summary_map.setdefault(bn, {'pending': 0, 'in_progress': 0, 'success': 0, 'failed': 0, 'total': 0})
            summary_map[bn][st] = r['cnt']
            summary_map[bn]['total'] += r['cnt']
    # 转 dict 并注入 report_summary(sqlite3.Row 不支持赋值)
    out = []
    for b in rows:
        d = dict(b)
        d['report_summary'] = summary_map.get(d['batch_no'], {'pending': 0, 'in_progress': 0, 'success': 0, 'failed': 0, 'total': 0})
        out.append(d)
    return jsonify(out)

# v0.9: 端口分配工具 — 回传完整子端口
@app.route('/api/port-alloc/apply', methods=['POST'])
def port_alloc_apply():
    """v0.14 报备人员回传"报备完整子端口"到原批次（真真设计）。
    body: {rows: [{完整子端口, 报备主端口, ...}], user_id}
    逻辑:
      1. 按批次号分组
      2. 原行 status='placeholder', orig_status=<原 status>
      3. 插入 N 个新行（每行 1 个报备完整子端口）：status='pending', report_status='pending', port_alloc_inserted=1
      4. 整批 demand_batch.status = 'reporting'
    """
    body = request.get_json() or {}
    user_id = body.get('user_id', 0)
    rows = body.get('rows', [])
    if not user_id:
        return jsonify({'error': '需要 user_id'}), 400
    if not rows:
        return jsonify({'error': 'rows 不能为空'}), 400
    user = query("SELECT * FROM user WHERE id=?", (user_id,), one=True)
    if not user:
        return jsonify({'error': '用户不存在'}), 404
    if user['dept'] != 'baobei' and user['role'] != 'E_admin':
        return jsonify({'error': '权限不足：仅实名报备部门和管理员可回传端口'}), 403
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    created = 0
    marked = 0
    skipped_signed_done = 0  # v0.23: 上传里该签名的原行已处理过，跳过
    by_batch = {}
    for r in rows:
        bn = r.get('批次号') or r.get('batch_no')
        if not bn: continue
        if bn not in by_batch: by_batch[bn] = []
        by_batch[bn].append(r)
    for batch_no, batch_rows in by_batch.items():
        # v0.23: 拿该批次全部行（alloc=0 原行 + alloc=1 新行）
        all_rows = query(
            "SELECT * FROM demand WHERE batch_no=? ORDER BY row_index, id",
            (batch_no,)
        )
        # 按签名分组：原行（alloc=0）和新行（alloc=1）
        by_sig_orig = {}   # sig -> [orig row] (alloc=0)
        by_sig_alloc = {} # sig -> [alloc=1 rows] (工具已生成行)
        for ar in all_rows:
            by_sig_orig.setdefault(ar['signature'], []).append(ar)
        # 按签名分组 upload 行
        upload_by_sig = {}
        for r in batch_rows:
            sig = (r.get('签名') or r.get('signature') or '').strip()
            if not sig: continue
            upload_by_sig.setdefault(sig, []).append(r)
        # v0.23: 处理每个有 upload 行的签名
        # "已处理过" = 该签名在 demand 表里已有 alloc=1 行（有工具生成的新行）
        for sig, sig_uploads in upload_by_sig.items():
            sig_alloc_rows = by_sig_alloc.get(sig, [])
            if sig_alloc_rows:
                # v0.23: 该签名已有 alloc=1 新行，说明已处理过，跳过
                skipped_signed_done += len(sig_uploads)
                continue
            # 未处理过（alloc=1 行不存在），按 v0.20 逻辑处理
            sig_origs = by_sig_orig.get(sig, [])
            original_rows = sig_origs
            if not original_rows:
                # 在所有 alloc=0 行里找一个作为模板（跨签名借用，同批次字段通用）
                any_orig = [r for r in all_rows if r['port_alloc_inserted'] == 0]
                if any_orig:
                    original_rows = any_orig
                else:
                    continue  # 整批没有 alloc=0 原行（遗留）
            for orow in original_rows:
                if orow['status'] != 'placeholder':
                    execute("UPDATE demand SET status='placeholder', orig_status=? WHERE id=?",
                            (orow['status'], orow['id']))
            marked += sum(1 for o in original_rows if o['status'] != 'placeholder')
            max_row_idx = max((r['row_index'] for r in all_rows), default=0)
            by_orig_row = {or_['row_index']: or_ for or_ in original_rows}
            for r in sig_uploads:
                # v0.16: 优先使用"待报备完整子端口"（工具生成的真号），退回"完整子端口"
                concrete = r.get('待报备完整子端口') or r.get('完整子端口') or r.get('complete_sub_port') or ''
                if not concrete: continue
                # v0.17: 兼容"报备主端口"和"报备主端口（汇总）"两个列名
                port4 = (r.get('报备主端口') or r.get('报备主端口（汇总）') or '').strip()
                if not port4: continue
                # is_concrete_sub_port: 判断 concrete 是否是占位字符串
                def _is_placeholder_pa(v):
                    if not v: return True
                    return v.startswith('无尾号要求') or v.startswith('有尾号要求')
                is_concrete_pa = 0 if _is_placeholder_pa(concrete) else 1
                orig_row_idx = parseInt_safe(r.get('行号'), None)
                sig_from_row = sig
                port4_for_match = port4
                orig = None
                if sig_from_row:
                    for or_ in original_rows:
                        if or_['signature'] == sig_from_row and or_['port_id'] == port4_for_match:
                            orig = or_
                            break
                    if not orig:
                        for or_ in original_rows:
                            if or_['signature'] == sig_from_row:
                                orig = or_
                                break
                    if orig and sig_from_row:
                        same_sig_rows = [or_ for or_ in original_rows if or_['signature'] == sig_from_row and or_['port_id_raw']]
                        if same_sig_rows:
                            all_ports = []
                            for sr in same_sig_rows:
                                raw = sr['port_id_raw']
                                if not raw: continue
                                if re.search(r'[、，,\s]', raw):
                                    parts = re.split(r'[、，,\s]+', raw)
                                    for p in parts:
                                        if p and p not in all_ports: all_ports.append(p)
                                else:
                                    if raw not in all_ports: all_ports.append(raw)
                            if all_ports:
                                orig_port_id_raw = '、'.join(all_ports)
                            else:
                                orig_port_id_raw = orig['port_id_raw']
                        else:
                            orig_port_id_raw = orig['port_id_raw']
                    else:
                        orig_port_id_raw = orig['port_id_raw'] if orig else None
                if not orig and orig_row_idx is not None and orig_row_idx in by_orig_row:
                    orig = by_orig_row[orig_row_idx]
                if not orig:
                    orig = original_rows[0]
                orig_csp = orig['complete_sub_port']
                new_row_idx = max_row_idx + 1
                max_row_idx += 1
                try:
                    execute("""
                        INSERT INTO demand
                        (batch_no, row_index, biz_type, biz_source, signature, complete_sub_port,
                         company_account, company_name, sales_owner, carrier, port_id, port_id_raw,
                         sub_port_count, extension_count, tail_extend_position, port_pre_occupy,
                         business_category, business_sub_category, is_finance, remark,
                         region, report_handler,
                         initial_check_warned, initial_check_reason, force_submitted, status,
                         report_status, port_alloc_inserted, created_at, updated_at,
                         concrete_sub_port, is_concrete_sub_port)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', 'in_progress', 1, ?, ?,
                                ?, ?)
                    """, (batch_no, new_row_idx, orig['biz_type'], 'self',
                          orig['signature'], orig_csp,
                          orig['company_account'], orig['company_name'], orig['sales_owner'],
                          orig['carrier'], port4, orig_port_id_raw,
                          orig['sub_port_count'], orig['extension_count'],
                          orig['tail_extend_position'], r.get('端口预占位') or orig['port_pre_occupy'],
                          orig['business_category'], orig['business_sub_category'],
                          orig['is_finance'], orig['remark'],
                          orig['region'], orig['report_handler'],
                          orig['initial_check_warned'], orig['initial_check_reason'],
                          orig['force_submitted'],
                          now, now,
                          concrete, is_concrete_pa))
                    created += 1
                except Exception as e:
                    return jsonify({'error': f'插入失败：{e}'}), 500
        execute("UPDATE demand_batch SET status='reporting' WHERE batch_no=?", (batch_no,))
    return jsonify({
        'ok': True,
        'created': created,
        'marked_placeholder': marked,
        'skipped_signed_done': skipped_signed_done,  # v0.23: 签名已处理跳过数
        'batches': list(by_batch.keys()),
        'msg': f'原 {marked} 行变占位行，新增 {created} 个报备子端口行，跳过 {skipped_signed_done} 个已处理签名，整批进 reporting'
    }), 200

def parseInt_safe(v, default=None):
    try:
        return int(str(v).strip()) if v not in (None, '') else default
    except:
        return default

# v0.10: 端口拓位工具（sheet3-expand / port-extend）拆完后调此接口，
#        把"拆分后"表的具体子端口同步到需求管理系统的 demand 表。
#        原占位行写 concrete_sub_port 字段，复制 N 行新建 demand 行。
@app.route('/api/demands/sync-from-expansion', methods=['POST'])
def sync_from_expansion():
    body = request.get_json() or {}
    batch_no = body.get('batch_no')
    source_tool = body.get('source_tool', 'sheet3-expand')
    expansion_rows = body.get('expansion_rows', [])
    # v0.21: 权限校验(参考 port_alloc_apply)
    user_id = body.get('user_id', 0)
    if not user_id:
        return jsonify({'error': '需要 user_id'}), 400
    user = query("SELECT * FROM user WHERE id=?", (user_id,), one=True)
    if not user:
        return jsonify({'error': '用户不存在'}), 404
    if user['dept'] != 'baobei' and user['role'] != 'E_admin':
        return jsonify({'error': '权限不足:仅实名报备部/管理员可同步工具'}), 403
    # v0.20: 可选 demand_ids — 只同步选中的子需求。未提供=整批同步
    demand_ids_filter = body.get('demand_ids', None)
    if isinstance(demand_ids_filter, list):
        # 转 int,去 None
        demand_ids_filter = [int(x) for x in demand_ids_filter if isinstance(x, (int, float)) or (isinstance(x, str) and x.isdigit())]
    elif demand_ids_filter is not None:
        demand_ids_filter = None  # 传错则降级为整批

    if not batch_no:
        return jsonify({'error': '需要 batch_no'}), 400
    if not expansion_rows:
        return jsonify({'error': 'expansion_rows 不能为空'}), 400

    batch = query("SELECT * FROM demand_batch WHERE batch_no=?", (batch_no,), one=True)
    if not batch:
        return jsonify({'error': f'批次不存在: {batch_no}'}), 404

    # v0.20: 获取选中子需求的 (signature, carrier, port_id) 范围
    selected_keys = set()  # (signature, carrier, port_id)
    selected_orig_ids = set()  # 选中的原行 id（用于护原占位行不被删）
    if demand_ids_filter:
        for r in query("""
            SELECT id, signature, carrier, port_id FROM demand
            WHERE batch_no=? AND id IN ({})
        """.format(','.join('?' * len(demand_ids_filter))), (batch_no, *demand_ids_filter)):
            selected_keys.add((r['signature'], r['carrier'], r['port_id']))
            selected_orig_ids.add(r['id'])
    else:
        # 整批模式：拿全批所有 (sig+carrier+port) 作 为 selected_keys
        for r in query("""
            SELECT DISTINCT signature, carrier, port_id FROM demand
            WHERE batch_no=? AND port_alloc_inserted=0
        """, (batch_no,)):
            selected_keys.add((r['signature'], r['carrier'], r['port_id']))

    created = 0
    updated = 0
    skipped = 0
    details = []
    for r in expansion_rows:
        seq = (r.get('expansion_seq') or '原行').strip()
        signature = r.get('signature', '') or ''
        csp = r.get('complete_sub_port', '') or ''         # 原完整子端口（占位或具体）
        concrete = r.get('concrete_sub_port', '') or ''   # 待报备完整子端口（具体子端口）
        port4 = (r.get('port_id') or '').strip()
        carrier = r.get('carrier', '') or ''

        # v0.20: 按 demand_ids 过滤 — 未选中的 (sig+carrier+port) 跳过同步
        if demand_ids_filter is not None:
            key = (signature, carrier, port4)
            if key not in selected_keys:
                skipped += 1
                details.append({'action': 'skipped', 'reason': '未在 demand_ids 范围内', 'signature': signature, 'port_id': port4, 'concrete': concrete})
                continue
        # 兜底：concrete 缺失时退回 csp
        if not concrete:
            concrete = csp
        # 是否具体子端口：concrete 有值且不是占位字符串（无尾号要求/有尾号要求:xxx）
        def is_placeholder(v):
            if not v: return True
            return v.startswith('无尾号要求') or v.startswith('有尾号要求')
        is_concrete = 0 if is_placeholder(concrete) else 1

        # 阶段 1：仅限「原行」匹配现有 demand 行（同 csp 的原占位行）
        if seq == '原行':
            existing = query("""
                SELECT id FROM demand
                WHERE batch_no=? AND signature=? AND port_id=? AND carrier=? AND complete_sub_port=?
                LIMIT 1
            """, (batch_no, signature, port4, carrier, csp), one=True)
            if existing:
                execute("""
                    UPDATE demand
                    SET concrete_sub_port=?, is_concrete_sub_port=?, report_status=\'in_progress\', updated_at=CURRENT_TIMESTAMP
                    WHERE id=?
                """, (concrete, is_concrete, existing['id']))
                updated += 1
                details.append({'action': 'updated', 'demand_id': existing['id'], 'concrete': concrete, 'expansion_seq': seq})
                continue

        # 阶段 2：按 concrete_sub_port 幂等键（拓位后具体子端口不与原 csp 相同）
        existing = query("""
            SELECT id FROM demand
            WHERE batch_no=?
              AND COALESCE(concrete_sub_port, '') = ?
              AND signature=? AND port_id=?
            LIMIT 1
        """, (batch_no, concrete, signature, port4), one=True)
        if existing:
            execute("""
                UPDATE demand
                SET concrete_sub_port=?, is_concrete_sub_port=?, report_status=\'in_progress\', updated_at=CURRENT_TIMESTAMP
                WHERE id=?
            """, (concrete, is_concrete, existing['id']))
            updated += 1
            details.append({'action': 'updated', 'demand_id': existing['id'], 'concrete': concrete, 'expansion_seq': seq, 'note': '按 concrete 幂等'})
            continue

        # 阶段 3：找原行（同 sig+port+carrier）作为字段模板
        orig = query("""
            SELECT * FROM demand
            WHERE batch_no=? AND signature=? AND port_id=? AND carrier=?
            ORDER BY id ASC LIMIT 1
        """, (batch_no, signature, port4, carrier), one=True)
        if not orig:
            skipped += 1
            details.append({'action': 'skipped', 'reason': '原行不存在', 'signature': signature, 'port_id': port4, 'concrete': concrete})
            continue

        # 阶段 4：新建 demand 行
        max_idx = query("SELECT MAX(row_index) AS m FROM demand WHERE batch_no=?", (batch_no,), one=True)
        new_idx = (max_idx['m'] or 0) + 1
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        new_id = execute("""
            INSERT INTO demand
            (batch_no, row_index, biz_type, biz_source, signature, complete_sub_port,
             company_account, company_name, sales_owner, carrier, port_id, port_id_raw,
             sub_port_count, extension_count, tail_extend_position, port_pre_occupy,
             business_category, business_sub_category, is_finance, remark,
             region, report_handler,
             initial_check_warned, initial_check_reason, force_submitted, status,
             report_status, port_alloc_inserted, created_at, updated_at,
             concrete_sub_port, is_concrete_sub_port)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', 'in_progress', 1, ?, ?,
             ?, ?)
        """, (batch_no, new_idx, orig['biz_type'], 'self',
              orig['signature'], csp,
              orig['company_account'], orig['company_name'], orig['sales_owner'],
              orig['carrier'], port4, orig['port_id_raw'],
              orig['sub_port_count'], orig['extension_count'],
              orig['tail_extend_position'], orig['port_pre_occupy'],
              orig['business_category'], orig['business_sub_category'],
              orig['is_finance'], orig['remark'],
              orig['region'], orig['report_handler'],
              orig['initial_check_warned'], orig['initial_check_reason'],
              orig['force_submitted'],
              now, now,
              concrete, is_concrete))
        created += 1
        details.append({'action': 'created', 'demand_id': new_id, 'concrete': concrete, 'expansion_seq': seq})

    # 重新统计 concrete_count
    cnt = query("""
        SELECT COUNT(*) AS c FROM demand
        WHERE batch_no=? AND is_concrete_sub_port=1
    """, (batch_no,), one=True)
    concrete_count = cnt['c']
    execute("UPDATE demand_batch SET concrete_count=? WHERE batch_no=?", (concrete_count, batch_no))

    # v0.21.2: 同步完成后触发 recompute_batch_status, 整批 status 变 'partial_processing'/'processing'
    # (新插入 N 行 report_status='in_progress')
    recompute_batch_status(batch_no)

    # v0.20: 仅删选中的原占位行 / 选中的孤儿行。未选中的保持原状 (pending)
    # 1. 选中的原占位行 (sig+carrier+port 在 selected_keys 中 且 is_concrete_sub_port=0)
    # 2. 选中的孤儿行 (port_alloc_inserted=1 且 未在本次同步的 concrete_sub_port 中)
    # 未选中的原行（demand_ids 为空或者不在选中范围）保持 pending
    deleted_demands = 0
    if demand_ids_filter is not None:
        # 子集模式: 只删 selected_keys 内的原占位行 和选中的孤儿
        # v0.21.2: 用 selected_orig_ids (精确 demand_id 列表) 替代 (sig+car+port) 三元组
        # 避免误删: 同 (sig+car+port) 但不同报备人的需求
        if selected_orig_ids:
            id_placeholders = ','.join('?' * len(selected_orig_ids))
            # 1. 选中的原占位行 (port_alloc_inserted=0 且 未 sync)
            sql1 = f"""
                DELETE FROM demand
                WHERE batch_no=? AND is_concrete_sub_port=0 AND id IN ({id_placeholders})
            """
            cur1 = get_db().execute(sql1, (batch_no, *selected_orig_ids))
            deleted_demands += cur1.rowcount
            cur1.close()
            get_db().commit()
            # 2. 选中的孤儿行 (port_alloc_inserted=1 且 未在本次 concrete)
            sql2 = f"""
                DELETE FROM demand
                WHERE batch_no=? AND port_alloc_inserted=1
                  AND id IN ({id_placeholders})
                  AND concrete_sub_port NOT IN ({','.join('?' * len(expansion_rows))})
            """
            params2 = [batch_no] + list(selected_orig_ids) + [er.get('concrete_sub_port', '') for er in expansion_rows]
            cur2 = get_db().execute(sql2, tuple(params2))
            deleted_demands += cur2.rowcount
            cur2.close()
            get_db().commit()
    else:
        # 整批模式: 保留原物理删逻辑（v0.10）
        cur3 = get_db().execute("""
            DELETE FROM demand
            WHERE batch_no=? AND (
                is_concrete_sub_port=0
                OR (port_alloc_inserted=1 AND id NOT IN (
                    SELECT id FROM demand WHERE batch_no=? AND concrete_sub_port IS NOT NULL
                ))
            )
        """, (batch_no, batch_no))
        deleted_demands = cur3.rowcount
        cur3.close()
        get_db().commit()

    return jsonify({
        'ok': True,
        'batch_no': batch_no,
        'source_tool': source_tool,
        'created': created,
        'updated': updated,
        'skipped': skipped,
        'concrete_count': concrete_count,
        'deleted_placeholder': deleted_demands,
        'details_count': len(details)
    }), 200

@app.route('/api/batches/<batch_no>/unreport', methods=['POST'])
def unreport_batch(batch_no):
    """v0.14 撤回整批（baobei/admin 可见，整批从 reporting 退回 submitted）。"""
    body = request.get_json() or {}
    user_id = body.get('user_id', 0) or request.args.get('user_id', 0, type=int)
    if not user_id:
        return jsonify({'error': '需要 user_id'}), 400
    user = query("SELECT * FROM user WHERE id=?", (user_id,), one=True)
    if not user:
        return jsonify({'error': '用户不存在'}), 404
    if user['dept'] != 'baobei' and user['role'] != 'E_admin':
        return jsonify({'error': '权限不足：仅实名部/管理员可撤回'}), 403
    batch = query("SELECT * FROM demand_batch WHERE batch_no=?", (batch_no,), one=True)
    if not batch:
        return jsonify({'error': '批次不存在'}), 404
    if batch['status'] != 'reporting':
        return jsonify({'error': f'当前批次状态「{batch["status"]}」不允许撤回（仅 reporting 可撤回）'}), 400
    taken_over = query("SELECT COUNT(*) AS c FROM demand WHERE batch_no=? AND port_alloc_inserted=1 AND report_status != 'pending' AND status != 'cancelled'", (batch_no,), one=True)
    if taken_over['c'] > 0:
        return jsonify({'error': f'有 {taken_over["c"]} 个新行已被受理（处理/成功/失败），不能整批撤回'}), 400
    del_count = execute("DELETE FROM demand WHERE batch_no=? AND port_alloc_inserted=1", (batch_no,))
    rec_count = execute("UPDATE demand SET status=orig_status, orig_status=NULL WHERE batch_no=? AND status='placeholder' AND orig_status IS NOT NULL", (batch_no,))
    execute("UPDATE demand_batch SET status='submitted' WHERE batch_no=?", (batch_no,))
    return jsonify({
        'ok': True,
        'batch_no': batch_no,
        'deleted_new_rows': del_count,
        'recovered_orig_rows': rec_count,
        'new_batch_status': 'submitted',
        'msg': f'撤回成功：删 {del_count} 个新行，恢复 {rec_count} 个原行，整批退到 submitted'
    })

"""
需求管理系统 · Flask 后端 (v0.3)
范围: 仅报备实名（第 1 块）

启动: python3 app/app.py

v0.3 变更：
- 13 列新表单字段
- 报备主端口支持多值（顿号/全角逗号/英文逗号/空格分隔）
- 完整子端口 3 种合法值
- 自动从 channel_list 带 region/report_handler/is_finance
- 业务大类/细类动态配置 (business_config)
- admin 鉴权（role=E_admin）
- 批次列表 JOIN user → submitter_name
- 申请时间/申请人显示
"""
import os
import sys
import json
import sqlite3
import re
from datetime import datetime, date, timedelta
from functools import wraps
from flask import Flask, request, jsonify, g, send_from_directory
from flask_cors import CORS

# ========== 路径配置 ==========
# ========== 业务工具 ==========
def generate_batch_no():
    """v0.21.1: 用 MAX(batch_no) 解析数字部分 + 1,避免 UNIQUE 冲突"""
    import re as _re
    today = datetime.now().strftime('%Y%m%d')
    prefix = f'B{today}-'
    row = query("SELECT MAX(batch_no) AS m FROM demand_batch WHERE batch_no LIKE ?", (prefix + '%',), one=True)
    max_no = (row['m'] if row and row['m'] else None)
    if max_no:
        m = _re.match(rf'^{_re.escape(prefix)}(\d+)$', max_no)
        seq = int(m.group(1)) + 1 if m else 1
    else:
        seq = 1
    return f'{prefix}{seq:06d}' 

def generate_batch_name(submitter_dept, submitter_name, signatures):
    """
    v0.4 自动批次名：{部门}-{提交人}-{签名摘要}-{时间HHMM}
    例：战略客户部-战略客户部代表-信用飞/信用宝-1142
    签名多于 2 个：信用飞/信用宝等3个
    """
    sigs = [s for s in (signatures or []) if s]
    if not sigs:
        sig_part = '匿名'
    elif len(sigs) <= 2:
        sig_part = '/'.join(sigs)
    else:
        sig_part = f'{sigs[0]}/{sigs[1]}等{len(sigs)}个'
    time_part = datetime.now().strftime('%H%M')
    return f'{submitter_dept}-{submitter_name}-{sig_part}-{time_part}'

def parse_port_field(value):
    """
    v0.3 解析端口字段（去括号）：
    - "3228" → [{'port_id': '3228'}]
    - "3228、1802" → 2 条
    - "3228，1802，4336" → 3 条
    支持分隔符：顿号、全角逗号、英文逗号、空格、换行
    """
    value = str(value).strip()
    if not value:
        return None
    normalized = re.sub(r'[，,\s\n\r]+', '、', value)
    parts = [p.strip() for p in normalized.split('、') if p.strip()]
    if not parts:
        return None
    for p in parts:
        if not re.match(r'^\d{4}$', p):
            return None
    return [{'port_id': p} for p in parts]

def is_valid_complete_sub_port(value):
    """v0.5 校验完整子端口合法值: 纯数字 / 无尾号要求(申请新通道) / 有尾号要求:数字
    v0.21.1: 兼容全角/半角括号 + 冒号前后空格
    """
    if not value:
        return False
    v = str(value).strip()
    if re.match(r'^\d+$', v):
        return True
    # 兼容全角/半角括号
    if v in ('无尾号要求(申请新通道)', '无尾号要求（申请新通道）'):
        return True
    if v in ('有尾号要求(手填任意数字字符串)', '有尾号要求（手填任意数字字符串）'):
        return True
    if re.match(r'^有尾号要求\s*[:：]\s*\d+$', v):
        return True
    return False

# v0.6 业务类型映射：channel_list.channel_category → biz_type
def map_ch_category_to_biz_type(ch_category):
    if not ch_category:
        return 'sms'
    s = str(ch_category)
    if '5G' in s or '5g' in s.lower() or '阅信' in s:
        return '5g'
    if '视频' in s or '视信' in s:
        return 'video'
    return 'sms'

BIZ_TYPE_REVERSE = {'sms': '短信', '5g': '5G', 'video': '视信'}

# ========== admin 鉴权装饰器 ==========
def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user_id = request.args.get('user_id', 0, type=int)
        if not user_id:
            body = request.get_json(silent=True) or {}
            user_id = body.get('user_id', 0)
        if not user_id:
            return jsonify({'error': '需要 user_id'}), 400
        u = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True)
        if not u or u['role'] != 'E_admin':
            return jsonify({'error': '需要 E_admin 权限'}), 403
        return fn(*args, **kwargs)
    return wrapper


# GET /api/admin/check
@app.route('/api/admin/check', methods=['GET'])
def admin_check():
    user_id = request.args.get('user_id', 0, type=int)
    u = query("SELECT * FROM user WHERE id = ?", (user_id,), one=True) if user_id else None
    if not u:
        return jsonify({'is_admin': False}), 200
    return jsonify({'is_admin': u['role'] == 'E_admin', 'user': dict(u)})

# ========== 文档转换（v0.22 承诺函 PDF 转换） ==========
def _find_libreoffice():
    """跨平台找 LibreOffice 可执行文件"""
    lo = shutil.which('libreoffice')
    if lo:
        return lo
    lo = shutil.which('soffice')
    if lo:
        return lo
    # Windows 常见路径
    for p in [
        os.path.expandvars(r'%PROGRAMFILES%\LibreOffice\program\soffice.exe'),
        os.path.expandvars(r'%PROGRAMFILES(X86)%\LibreOffice\program\soffice.exe'),
        os.path.expandvars(r'%LOCALAPPDATA%\LibreOffice\program\soffice.exe'),
    ]:
        if os.path.isfile(p):
            return p
    return None

@app.route('/api/convert/docx-to-pdf', methods=['POST'])
def convert_docx_to_pdf():
    """
    POST 接收 .docx 文件 (multipart/form-data, field=file)，转 PDF 返回。
    用途：昆明联通承诺书 HTML 工具一键出 PDF。
    依赖：服务器安装 LibreOffice（apt-get install -y libreoffice）
    """
    if 'file' not in request.files:
        return jsonify({'error': '没有上传文件（field 名应为 file）'}), 400
    f = request.files['file']
    if not f.filename or not f.filename.lower().endswith('.docx'):
        return jsonify({'error': '只支持 .docx 文件'}), 400

    lo_bin = _find_libreoffice()

    def _do_weasyprint_convert(docx_path, tmp):
        """纯 Python DOCX→PDF（用 python-docx + WeasyPrint）"""
        from docx import Document as _Doc
        from io import BytesIO as _BytesIO
        doc = _Doc(docx_path)
        # 生成 HTML（中文合同风格，A4 竖版）
        css = """
        @page { size: A4; margin: 2cm 2.5cm; }
        body { font-family: 'SimSun', '宋体', 'Noto Serif CJK SC', serif; font-size: 14px; line-height: 1.8; color: #000; }
        h1 { text-align: center; font-size: 22px; font-weight: bold; margin: 0 0 8px 0; }
        .subtitle { text-align: center; font-size: 12px; color: #555; margin-bottom: 20px; }
        table { border-collapse: collapse; width: 100%; margin: 12px 0; }
        td { border: 1px solid #000; padding: 6px 10px; font-size: 13px; vertical-align: middle; }
        .label { font-weight: bold; width: 90px; text-align: center; background: #f5f5f5; }
        .content { font-weight: bold; }
        p { margin: 6px 0; text-indent: 2em; }
        .section-title { font-weight: bold; margin: 14px 0 6px 0; text-indent: 0; }
        .footer { margin-top: 30px; text-align: right; }
        """
        html_parts = ['<!DOCTYPE html>', '<html><head>',
            '<meta charset="UTF-8">',
            '<style>' + css + '</style>',
            '</head><body>']
        # 标题
        html_parts.append('<h1>商业性通信短信息端口合规承诺书</h1>')
        html_parts.append('<p class="subtitle">（本承诺书由企业自愿签署，具有法律效力）</p>')
        # 提取正文所有段落文本
        paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        # 第一个表（端口/签名/日期）
        tables = doc.tables
        if tables:
            html_parts.append('<table>')
            for row in tables[0].rows:
                cells = [c.text.strip() for c in row.cells]
                html_parts.append('<tr>')
                for ci, cell in enumerate(cells):
                    cls = 'label' if ci == 0 else 'content'
                    html_parts.append(f'<td class="{cls}">{cell}</td>')
                html_parts.append('</tr>')
            html_parts.append('</table>')
        # 后续段落（去掉标题行）
        body_paras = paras[1:] if paras else []
        for txt in body_paras:
            # 小标题/条款用粗体
            if txt.startswith('一') or txt.startswith('二') or txt.startswith('三') or txt.startswith('（'):
                html_parts.append(f'<p class="section-title">{txt}</p>')
            else:
                html_parts.append(f'<p>{txt}</p>')
        html_parts.append('</body></html>')
        html_str = '\n'.join(html_parts)
        pdf_buf = _WeasyHTML(string=html_str).write_pdf()
        return pdf_buf

    with tempfile.TemporaryDirectory() as tmp:
        safe_name = re.sub(r'[^\w\.\-]', '_', f.filename) or 'input.docx'
        docx_path = os.path.join(tmp, safe_name)
        f.save(docx_path)
        from io import BytesIO
        try:
            if lo_bin:
                # LibreOffice 路径（优先，质量最高）
                result = subprocess.run(
                    [lo_bin, '--headless', '--convert-to', 'pdf', '--outdir', tmp, docx_path],
                    capture_output=True, text=True, timeout=60
                )
                if result.returncode != 0:
                    raise RuntimeError(result.stderr or result.stdout)
                pdf_path = os.path.join(tmp, safe_name.rsplit('.', 1)[0] + '.pdf')
                if not os.path.isfile(pdf_path):
                    raise FileNotFoundError('PDF 未生成')
                with open(pdf_path, 'rb') as pf:
                    pdf_bytes = pf.read()
            elif _WEASYPRINT_OK:
                # WeasyPrint 兜底（纯 Python，不需要系统级依赖）
                pdf_bytes = _do_weasyprint_convert(docx_path, tmp)
            else:
                return jsonify({
                    'error': '服务器未安装 LibreOffice，且 WeasyPrint 也不可用',
                    'hint': '请用 Word/WPS 打开 docx 后另存为 PDF'
                }), 503
            pdf_name = (f.filename.rsplit('.', 1)[0] or 'document') + '.pdf'
            return send_file(
                BytesIO(pdf_bytes),
                as_attachment=True,
                download_name=pdf_name,
                mimetype='application/pdf'
            )
        except subprocess.TimeoutExpired:
            return jsonify({'error': 'LibreOffice 转换超时（>60s）'}), 504
        except Exception as e:
            return jsonify({'error': 'PDF 转换失败', 'detail': str(e)}), 500

# ========== 反代：承诺函生成平台 (5001) ==========
import requests as _requests
COMMITMENT_BACKEND = os.environ.get('COMMITMENT_BACKEND', 'http://127.0.0.1:5001')

@app.route('/commitment/', defaults={'path': ''}, methods=['GET', 'POST', 'PUT', 'DELETE', 'PATCH', 'OPTIONS'])
@app.route('/commitment/<path:path>', methods=['GET', 'POST', 'PUT', 'DELETE', 'PATCH', 'OPTIONS'])
def commitment_proxy(path):
    """反代 /commitment/* 到 commitment-letter-platform (5001)
    例：/commitment/api/health → http://127.0.0.1:5001/api/health
    """
    target = f'{COMMITMENT_BACKEND}/{path}'
    try:
        if request.method in ('POST', 'PUT', 'PATCH'):
            resp = _requests.request(
                request.method, target,
                data=request.get_data(),
                headers={k: v for k, v in request.headers.items() if k.lower() not in ('host', 'content-length', 'connection')},
                params=request.args,
                allow_redirects=False,
                timeout=120
            )
        else:
            resp = _requests.request(
                request.method, target,
                headers={k: v for k, v in request.headers.items() if k.lower() not in ('host', 'content-length', 'connection')},
                params=request.args,
                allow_redirects=False,
                timeout=120
            )
        # 透传响应（去掉 hop-by-hop headers）
        excluded = ('connection', 'keep-alive', 'proxy-authenticate', 'proxy-authorization',
                    'te', 'trailers', 'transfer-encoding', 'upgrade')
        headers = [(k, v) for k, v in resp.headers.items() if k.lower() not in excluded]
        return (resp.content, resp.status_code, headers)
    except _requests.exceptions.ConnectionError:
        return jsonify({'error': '承诺函平台后端未启动（5001）'}), 503
    except Exception as e:
        return jsonify({'error': f'反代错误: {str(e)}'}), 500

# ========== 启动 ==========
if __name__ == '__main__':
    if not os.path.exists(DB_PATH):
        print(f"⚠️ 数据库不存在: {DB_PATH}")
        sys.exit(1)
    print(f"✅ 数据库: {DB_PATH}")
    with app.app_context():
        stats = {}
        for t in ['channel_list', 'rejection_log', 'demand', 'demand_batch', 'user', 'business_config']:
            r = query(f"SELECT COUNT(*) AS cnt FROM {t}", one=True)
            stats[t] = r['cnt']
    print(f"📊 数据统计: {stats}")
    print(f"🚀 启动 Flask @ http://0.0.0.0:5000")
    app.run(host='0.0.0.0', port=5000, debug=False)
