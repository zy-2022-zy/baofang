# 第 1 块：提需求 + 自动拆行 + 警告式初判

> **范围**：5 阶段流程的前 1.5 阶段（提需求 → 系统初判 → 待处理）
> **目标**：A 需求方能在系统里填一份需求单，提交后系统按规则自动拆成多条子需求，过程中如果匹配到历史驳回记录就给警告。
> **不含**：驳回表的人工维护入口、报备中状态、报备结果回填（后续块做）

---

## 1. 这块要解决什么

| 痛点 | 这块的解法 |
|------|-----------|
| 需求方在群里乱发 Excel，格式不统一 | 8 条业务线对应 4 套字段模板，提交时强校验 |
| 一个需求跨多个运营商/区域要手抄多份 | 提交后按 (运营商 × 区域 × 端口) 自动拆成 N 条子需求 |
| 之前被运营商驳回的签名，换个皮又来一次 | 提交前查驳回表，命中就 ⚠️ 警告 |
| 需求方不知道"为啥不能这么填" | 字段级错误提示（必填/格式/枚举） |

---

## 2. 数据模型（4 张表）

### 2.1 `demand` 需求主表（拆行后的子需求粒度）

**一条记录 = 一个子需求（已经被拆过的最小单位）**

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `id` | int | ✅ | 主键 |
| `batch_no` | string | ✅ | 批次号，同一次提交的所有子需求共享一个号（格式：`B{YYYYMMDD}-{6位序号}`） |
| `parent_id` | int | ❌ | 如果是子需求，记录父需求 id；如果是首次提交的"逻辑总单"，则为空 |
| `biz_line` | enum | ✅ | 业务线：`realname`/`yinliu`/`yinliu_support`/`mianbai`/`video_sms`/`5g_msg`/`third_party_baobei`/`third_party_mianbai` |
| `biz_line_template` | enum | ✅ | 字段模板：`A_实名`/`B_免白`/`C_引流`/`D_引流支撑`（根据业务线自动选） |
| `carrier` | enum | ✅ | 运营商：`mobile`/`unicom`/`telecom` |
| `region` | string | ✅ | 区域，如"广东"、"海南"、"重庆" |
| `port` | string | ✅ | 端口号（子端口） |
| `signature` | string | ✅ | 短信签名 |
| `biz_content` | string | ❌ | 业务说明/用途 |
| `extra_fields` | json | ❌ | 业务线专属字段（4 套模板差异字段塞这里） |
| `status` | enum | ✅ | 状态：`pending`/`initial_checked`/`assigned`/`reporting`/`done`/`failed` |
| `initial_check_warned` | bool | ✅ | 是否被初判警告过 |
| `initial_check_reason` | string | ❌ | 警告原因 |
| `force_submitted` | bool | ✅ | 是否强制提交（带警告时勾选） |
| `submitter_id` | int | ✅ | 提交人（关联 user 表） |
| `submitter_dept` | enum | ✅ | 提交人部门 |
| `created_at` | datetime | ✅ | 创建时间 |
| `updated_at` | datetime | ✅ | 更新时间 |

### 2.2 `demand_batch` 需求批次表（拆行前的"总单"）

**一条记录 = 一次提单行为（可能拆出 N 条子需求）**

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `id` | int | ✅ | 主键 |
| `batch_no` | string | ✅ | 同 `demand.batch_no` |
| `title` | string | ✅ | 需求标题（提单时填） |
| `biz_line` | enum | ✅ | 业务线（影响子需求展开方式） |
| `submitter_id` | int | ✅ | 提交人 |
| `submitter_dept` | enum | ✅ | 提交人部门 |
| `split_rule` | enum | ✅ | 拆行规则：`by_carrier`/`by_region`/`by_carrier_region_port` |
| `split_count` | int | ✅ | 拆出多少条子需求 |
| `raw_payload` | json | ✅ | 提交时的原始表单数据（用于回溯/审计） |
| `has_warning` | bool | ✅ | 拆出的子需求里是否至少有一条被警告 |
| `force_submitted` | bool | ✅ | 提单时是否带警告强制提交 |
| `created_at` | datetime | ✅ | 创建时间 |

### 2.3 `rejection_log` 驳回日志（初判查询依据）

**一条记录 = 一次"运营商驳回"事件**

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `id` | int | ✅ | 主键 |
| `signature` | string | ✅ | 短信签名 |
| `carrier` | enum | ✅ | 运营商 |
| `port` | string | ❌ | 端口（不填则按"签名+运营商"全匹配） |
| `rejection_reason` | string | ✅ | 驳回原因 |
| `source` | enum | ✅ | 来源：`auto_from_excel`/`manual` |
| `reporter_id` | int | ❌ | 上报人（D 处理方） |
| `rejection_date` | date | ✅ | 驳回日期 |
| `created_at` | datetime | ✅ | 入库时间 |

> **第 1 块先 read-only**：这块只读这张表，不做新增/编辑入口。**第 2 块**才加人工维护界面 + Excel 自动入库。

### 2.4 `user` 用户表（最小可用版）

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `id` | int | ✅ | 主键 |
| `name` | string | ✅ | 姓名 |
| `dept` | enum | ✅ | 部门 |
| `role` | enum | ✅ | 角色：`A_requester`/`D_handler`/`E_admin` |
| `created_at` | datetime | ✅ | 创建时间 |

> 第 1 块**先用 5 个 mock 用户**做登录切换演示，不做完整账号体系。**第 8 块**才上正式账号。

---

## 3. 4 套字段模板

> 业务线 → 字段模板的映射见第 3 步的 8 条业务线表。**前端的表单根据 `biz_line` 动态渲染对应模板的字段。**

### 模板 A：实名（实名 / 视频短信 / 5G 消息 / 三方报备）
共用字段：`signature` `port` `carrier` `region` `biz_line` `biz_content`（用途）

### 模板 B：免白（免白 / 三方免白）
共用字段：`signature` `port` `carrier` `region` `biz_line` `biz_content`
专属字段：`company_name`（实名企业名称）、`expiry_date`（免白失效日期）

### 模板 C：引流
共用字段：`signature` `port` `carrier` `region` `biz_line` `biz_content`
专属字段：`yinliu_sub_port`（引流子端口）、`extension_code`（扩展码）

### 模板 D：引流支撑
共用字段：`signature` `port` `carrier` `region` `biz_line` `biz_content`
专属字段：`support_target_port`（被支撑的引流端口）、`relation_type`（关联类型）

---

## 4. 拆行规则

提单时填写的内容如果是**多条同结构数据**（比如一次要报 3 个端口），按以下规则之一拆：

| `split_rule` | 拆行粒度 | 适用场景 |
|--------------|---------|---------|
| `by_carrier` | 1 条输入 → N 条子需求（按运营商） | "这个签名要报移动+联通+电信" |
| `by_region` | 1 条输入 → N 条子需求（按区域） | "这个签名要报广东+海南+重庆" |
| `by_carrier_region_port` | 1 条输入 → N 条子需求（笛卡尔积） | "这个签名要报 3 运营商 × 3 区域 = 9 条" |

**提单表单设计**：
- 用户在前端用**动态表格**录入 N 行（每行：运营商 + 区域 + 端口 + 签名）
- 提交时根据 `split_rule` 决定是否进一步拆（如已按行录入，选 `none`）
- 系统自动生成 `batch_no` + 逐条 `demand` 记录

---

## 5. 警告式初判流程

```
A 提交需求单
    ↓
[后端] 按 batch_no 逐条 demand 查 rejection_log
    ↓
匹配规则（任一命中即警告）：
  1. signature + carrier 完全匹配
  2. signature + carrier + port 完全匹配
  3. signature 全运营商驳回（port 为空 + 该 signature 在某 carrier 下有任意驳回）
    ↓
命中 → demand.initial_check_warned = true
     → demand.initial_check_reason = "{rejection_date} 该签名曾因 '{reason}' 被{carrier}驳回"
     → demand_batch.has_warning = true
未命中 → 静默通过
    ↓
返回结果给前端
    ↓
【前端分支】
  全部无警告 → 弹"提交成功"，跳到列表
  有警告 → 弹"⚠️ X 条子需求匹配到历史驳回记录"，列表标红，要求勾选"我已知晓风险，强制提交"才能继续
```

**关键设计**：
- ⚠️ **警告不阻断流程**，但 `force_submitted=true` 会留痕（后续可审计）
- `pending` 状态暂时只有 `pending`/`initial_checked` 两个子状态，对应**未初判**/**已初判**
- 真正的"待处理"（分配给 D）属于第 2 块范围

---

## 6. 状态机（第 1 块涉及部分）

```
[pending] ---提交后立刻---> [initial_checked]
                              ↓
                     (第 2 块：分配给 D)
                              ↓
                          [assigned]
                              ↓
                     (D 启动报备)
                              ↓
                          [reporting]
                              ↓
                     (D 上传报备结果)
                              ↓
                     [done] / [failed]
```

**第 1 块覆盖**：`pending` → `initial_checked`，其他状态**先禁用**（前端不显示操作按钮）。

---

## 7. API 设计（最小集）

| Method | Path | 用途 | 请求/响应 |
|--------|------|------|----------|
| GET | `/api/templates?biz_line=xxx` | 拿对应业务线的字段模板 | 响应：模板字段定义 JSON |
| POST | `/api/demands` | 提需求（自动拆行 + 初判） | 请求：表单 + split_rule / 响应：batch_no + N 条子需求 + warning 列表 |
| GET | `/api/demands?batch_no=xxx` | 查某批次的所有子需求 | 响应：demand 列表 |
| GET | `/api/users/me` | 当前登录用户（mock） | 响应：user 信息 |
| GET | `/api/batches` | 我的批次列表 | 响应：demand_batch 列表（按部门权限过滤） |

**权限规则（第 1 块简化版）**：
- 提交：任意已登录用户
- 查看批次：客服部/运营部/实名部/管理员看全部；战略客户部/资源采购部只看本部门

---

## 8. 前端页面（3 个）

### 8.1 提需求页 `/submit`
- 顶部：业务线下拉（8 选 1）
- 中部：动态字段表单（根据业务线 + 模板动态渲染）
- 表格区：可增删行的"运营商-区域-端口-签名"录入表
- 底部：拆行规则选择 + 提交按钮
- 提交后：弹警告窗（若有）→ 跳批次详情页

### 8.2 批次详情页 `/batch/:batch_no`
- 头部：批次号、提交人、提交时间、拆出 N 条
- 表格：该批次所有子需求，命中警告的行**整行标红**，悬浮显示警告原因
- 底部：返回按钮

### 8.3 我的批次列表 `/batches`
- 表格：批次号 / 业务线 / 子需求数 / 是否带警告 / 提交时间 / 状态
- 行点击 → 进批次详情

---

## 9. 验收标准

- [ ] 8 个业务线都能选，表单字段随业务线动态切换
- [ ] 一条输入（3 行运营商+区域+端口）能正确拆成 3 条子需求
- [ ] 按运营商笛卡尔积拆能正确算出行数
- [ ] 驳回表里有 `signature="XX"` + `carrier="mobile"` 的记录 → 提单时如果含相同签名+运营商 → 该子需求被标红警告
- [ ] 弹窗里能看清警告原因（驳回日期、原因、运营商）
- [ ] 勾选"强制提交"后能正常入库，且 `force_submitted=true` 留痕
- [ ] 战略客户部账号登录后，批次列表只显示本部门提交的
- [ ] 实名部/客服部/运营部账号登录后，批次列表显示全部
- [ ] 提交后批次号格式 `B20260610-000001`

---

## 10. 不在这块范围

明确**不做**的（避免范围蔓延）：
- ❌ 驳回表的人工维护界面（第 2 块）
- ❌ 报备结果 Excel 上传 + 自动分析（第 3 块）
- ❌ D 处理方分配/操作界面（第 2 块）
- ❌ 8 条业务线所有字段的精细枚举值（先用文本框，第 4 块再加枚举校验）
- ❌ 真实账号体系/SSO（第 8 块）
- ❌ 历史数据导入工具（第 7 块）
- ❌ 通知/消息推送（第 6 块）
- ❌ 权限到字段级的精细控制（第 8 块）

---

## 11. 开发节奏

| 天 | 交付物 |
|----|--------|
| Day 1 | Flask 后端骨架（4 张表 + 3 个核心 API）+ SQLite 初始化脚本 + mock 数据（含 3 条驳回记录） |
| Day 2 | 单文件 HTML 前端（提需求页 + 批次详情页 + 列表页） |
| Day 3 | 联调：提需求 → 拆行 → 初判警告 → 列表展示，全流程跑通 |
| Day 4 | 修 bug + 给真真看 demo + 收反馈 |

**Day 4 收尾时**，我会把第 1 块 demo + 本文档的更新版一起给你。
