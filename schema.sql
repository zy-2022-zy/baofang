-- 需求管理系统 数据库 Schema (v0.3)
-- 数据库: SQLite
-- 范围: 仅报备实名（第 1 块）
-- 变更: 13 列表单字段 + business_config 表 + 申请时间/申请人显示

-- ===============================================
-- 1. channel_list 全盘通道清单（真真导入）
-- ===============================================
DROP TABLE IF EXISTS channel_list;
CREATE TABLE channel_list (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    port_code       TEXT NOT NULL,                -- 8 位完整主端口（如 10693228）
    port_id         TEXT NOT NULL,                -- port_code 后 4 位（如 3228）
    channel_name    TEXT NOT NULL,                -- 通道名称
    channel_category TEXT NOT NULL,               -- 短信 / 视信 / 5G
    region          TEXT NOT NULL,                -- 区域（佛山/东莞/揭阳...）
    carrier         TEXT NOT NULL,                -- 移动 / 联通 / 电信
    info_type       TEXT,                         -- 短信 / 视信
    is_finance      BOOLEAN,                      -- 通道是否金融行业
    report_handler  TEXT NOT NULL,                -- 报备人
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE UNIQUE INDEX idx_channel_port_code ON channel_list(port_code, channel_category, info_type);
CREATE INDEX idx_channel_port_carrier ON channel_list(port_id, carrier);

-- ===============================================
-- 2. demand_batch 需求批次表
-- ===============================================
DROP TABLE IF EXISTS demand_batch;
CREATE TABLE demand_batch (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_no        TEXT NOT NULL UNIQUE,         -- B{YYYYMMDD}-{6位}
    batch_name      TEXT,                         -- 可空：单条=NULL，批量=必填
    biz_line        TEXT NOT NULL,                -- realname
    submitter_id    INTEGER NOT NULL,             -- 关联 user.id
    submitter_dept  TEXT NOT NULL,                -- 5 部门之一
    sub_demand_count INTEGER NOT NULL DEFAULT 0,  -- 拆出多少条
    has_warning     BOOLEAN DEFAULT 0,            -- 是否带警告
    force_submitted BOOLEAN DEFAULT 0,            -- 是否强制提交
    status          TEXT NOT NULL DEFAULT 'submitted',
    biz_type        TEXT NOT NULL DEFAULT 'sms',  -- v0.6: sms/5g/video
    biz_source      TEXT NOT NULL DEFAULT 'self', -- v0.7: self/eco (自营/三方)
    raw_payload     TEXT,                         -- JSON: 原始表单数据
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    cancelled_at    DATETIME,                     -- v0.8 作废时间
    cancelled_by    INTEGER,                      -- v0.8 作废人 user.id
    concrete_count  INTEGER DEFAULT 0             -- v0.10 拓位后具体子端口计数（sync-from-expansion 同步后统计）
);

CREATE INDEX idx_batch_dept ON demand_batch(submitter_dept, created_at);
CREATE INDEX idx_batch_user ON demand_batch(submitter_id, created_at);

-- ===============================================
-- 3. demand 需求子需求表（13 字段版 v0.3）
-- ===============================================
DROP TABLE IF EXISTS demand;
CREATE TABLE demand (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_no                TEXT NOT NULL,        -- 关联 demand_batch
    row_index               INTEGER NOT NULL,      -- 在原表单里的行号（用于排序）
    biz_type                TEXT NOT NULL DEFAULT 'sms',  -- v0.6: sms/5g/video
    biz_source              TEXT NOT NULL DEFAULT 'self', -- v0.7: self/eco
    channel_eco_id          INTEGER,              -- v0.7: 三方通道ID（FK channel_eco.id）
    channel_supplier        TEXT,                  -- v0.7: 供应商（冗余存）

    -- 13 个核心字段（按真真给的顺序）
    signature               TEXT NOT NULL,         -- 1. 签名
    complete_sub_port       TEXT NOT NULL,         -- 2. 完整子端口（手填/特殊值）
    company_account         TEXT,                  -- 3. 企业账号
    company_name            TEXT,                  -- 4. 企业名称
    sales_owner             TEXT,                  -- 5. 所属销售
    carrier                 TEXT NOT NULL,         -- 6. 所属运营商
    port_id                 TEXT NOT NULL,         -- 7. 报备主端口（4 位，拆分后单值）
    port_id_raw             TEXT,                  -- v0.10: 原始多值（如"4336、0287、3356"），baobei/admin 导出时用此字段保持原样
    sub_port_count          INTEGER NOT NULL DEFAULT 1,  -- 8. 子端口数
    extension_count         INTEGER NOT NULL DEFAULT 0,  -- 9. 拓位数
    tail_extend_position TEXT,                   -- v0.8: 拓位要求 (before/after) 有尾号要求+拓位>1时必填
    port_pre_occupy         TEXT,                   -- v0.8: 端口预占位（需求方填，认证部下载后填实际）
    port_alloc_inserted     INTEGER DEFAULT 0,      -- v0.14: 标记“工具拆行”插入的新行（1=插入的；0=原始）
    orig_status             TEXT,                  -- v0.14: 占位行原 status（撤回整批时恢复用）
    business_category       TEXT NOT NULL,         -- 10. 业务大类（55/77...）
    business_sub_category   TEXT NOT NULL,         -- 11. 业务细类（1/2/4...）
    is_finance              BOOLEAN DEFAULT 0,     -- 12. 是否金融类客户（从 channel_list 自动带）
    remark                  TEXT,                  -- 13. 备注

    -- 自动带出（从 channel_list 查）
    region                  TEXT,                  -- 区域
    report_handler          TEXT,                  -- 报备人（≠所属销售）

    -- 初判
    initial_check_warned    BOOLEAN DEFAULT 0,
    initial_check_reason    TEXT,
    force_submitted         BOOLEAN DEFAULT 0,
    status                  TEXT NOT NULL DEFAULT 'pending',

    -- v0.5 报备结果反馈（真真部门处理后回填）
    report_status           TEXT DEFAULT 'pending',     -- pending/in_progress/success/failed
    report_failure_reason   TEXT,                       -- 失败原因
    report_completed_at     DATETIME,                   -- 报备完成时间
    report_completed_by     TEXT,                       -- 报备完成人

    created_at              DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at              DATETIME DEFAULT CURRENT_TIMESTAMP,
    cancelled_at            DATETIME,            -- v0.8 作废时间
    cancelled_by            INTEGER,             -- v0.8 作废人 user.id
    concrete_sub_port       TEXT,                -- v0.10 拓位后待报备的完整子端口（sheet3-expand/port-extend 同步后写入）
    is_concrete_sub_port    INTEGER DEFAULT 0    -- v0.10 是否为具体子端口（0=占位/原行，1=拓位后具体子端口）
);

CREATE INDEX idx_demand_batch ON demand(batch_no);
CREATE INDEX idx_demand_port_carrier ON demand(port_id, carrier);
CREATE INDEX idx_demand_status ON demand(status);
CREATE INDEX idx_demand_signature ON demand(signature, carrier);

-- ===============================================
-- 4. rejection_log 驳回日志（真真导入）
-- ===============================================
DROP TABLE IF EXISTS rejection_log;
CREATE TABLE rejection_log (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    company_name    TEXT,                         -- 企业名称（A 列，可能 #N/A）
    signature       TEXT NOT NULL,                -- 签名（B 列，匹配键）
    rejection_reason TEXT NOT NULL,               -- 驳回原因（C 列）
    region          TEXT,                         -- 区域（D 列，混填）
    carrier         TEXT NOT NULL,                -- 运营商（E 列）
    rejection_date  INTEGER,                      -- 日期（F 列，Excel 序列号）
    remark          TEXT,                         -- 备注（G 列）
    source          TEXT DEFAULT 'auto_from_excel',
    reporter_id     INTEGER,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_rejection_sig_carrier ON rejection_log(signature, carrier);

-- ===============================================
-- 5. user 用户表（mock 6 人）
-- ===============================================
DROP TABLE IF EXISTS user;
CREATE TABLE user (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    name            TEXT NOT NULL,
    dept            TEXT NOT NULL,
    role            TEXT NOT NULL,
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP
);

INSERT INTO user (id, name, dept, role) VALUES
    (1, '战略客户部代表', 'strategic',        'A_requester'),
    (2, '客服部代表',     'customer_service', 'A_requester'),
    (3, '资源采购部代表', 'resource',         'A_requester'),
    (4, '运营部代表',     'operations',       'A_requester'),
    (5, '实名与报备支持部代表', 'baobei',     'D_handler'),
    (6, '系统管理员',     'baobei',           'E_admin');

-- ===============================================
-- 6. business_config 业务大类/细类配置（v0.3 新增）
-- ===============================================
DROP TABLE IF EXISTS business_config;
CREATE TABLE business_config (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    config_type TEXT NOT NULL,                    -- 'category' 或 'sub_category'
    value       TEXT NOT NULL,                    -- '55' / '77' / '1' / '2' / '4'
    sort_order  INTEGER DEFAULT 0,
    enabled     BOOLEAN DEFAULT 1,
    created_at  DATETIME DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(config_type, value)
);

-- 种子数据
INSERT INTO business_config (config_type, value, sort_order) VALUES
    ('category',     '55', 1),
    ('category',     '77', 2),
    ('sub_category', '1',  1),
    ('sub_category', '2',  2),
    ('sub_category', '4',  3);

-- ===============================================
-- 8. report_log 报备结果回传日志表（v0.8）
-- ===============================================
DROP TABLE IF EXISTS report_log;
CREATE TABLE report_log (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    demand_id       INTEGER NOT NULL,                -- 关联 demand.id
    batch_no        TEXT NOT NULL,                   -- 冗余存，便于按批次查
    old_status      TEXT,                            -- pending/in_progress/success/failed (操作前)
    new_status      TEXT NOT NULL,                   -- pending/in_progress/success/failed (操作后)
    failure_reason  TEXT,                            -- 失败原因（仅 failed 时）
    business_time   DATETIME,                        -- 报备时间（来自业务系统/运营商）
    changed_by      TEXT NOT NULL,                   -- 操作人姓名
    changed_at      DATETIME NOT NULL,               -- 录入时间（系统当前时间）
    operator_user_id INTEGER,                        -- 操作人 user.id
    import_source   TEXT NOT NULL DEFAULT 'manual',  -- 'manual' / 'tool:batch_xxx'
    remark          TEXT                             -- 备注（可选）
);

CREATE INDEX idx_report_log_demand ON report_log(demand_id, changed_at);
CREATE INDEX idx_report_log_batch ON report_log(batch_no, changed_at);

-- ===============================================
-- 7. channel_eco 三方/生态合作通道资源表（v0.7）
-- ===============================================
DROP TABLE IF EXISTS channel_eco;
CREATE TABLE channel_eco (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    source          TEXT NOT NULL,                -- 来源：生态/自营
    resource_type   TEXT NOT NULL,                -- 资源类型：短信/视信/5G/专用
    channel_status  TEXT NOT NULL,                -- 通道状态：启用/停用
    channel_id      TEXT NOT NULL UNIQUE,         -- 通道ID（生态方内部 ID）
    channel_name    TEXT NOT NULL,                -- 通道名（含 106XX 主端口号）
    gateway_account TEXT NOT NULL,                -- 网关账号
    carrier         TEXT NOT NULL,                -- 运营商
    supplier        TEXT NOT NULL,                -- 供应商
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_channel_eco_carrier ON channel_eco(carrier);
CREATE INDEX idx_channel_eco_supplier ON channel_eco(supplier);
