需求管理系统 (baobei-demand-hub) · 源码包

来源：workspace/projects/baobei-demand-hub/

## 文件清单

| 文件 | 行数 | 内容 |
|---|---|---|
| app.py | 2434 | Flask 后端全部（41 个 API + 业务工具）|
| schema.sql | 219 | 8 张表建表语句（含字段注释）|
| index.html | 3837 | 单文件前端（含 HTML 面板 + 全部 JS）|
| scripts/*.py | 3 个 | Excel 导入脚本：通道清单 / 三方通道资源 / 驳回日志 |
| docs/*.md | 7 个 | 设计文档（数据模型 / 修改计划 / 存量表分析）|

## 后端 app.py 结构速览

| 行号 | 内容 |
|---|---|
| L1-60 | 头部说明 + 路径配置（DB_PATH 等）|
| L108-125 | generate_batch_no() 批次号生成（v0.21.1 MAX+1）|
| L126-160 | 自动批次名 + 端口字段解析（去括号）|
| L164-180 | 完整子端口 3 种合法值校验 |
| L205-212 | admin 鉴权（role=E_admin）|
| L213-490 | 用户 / 通道 / 业务配置 / 三方通道 / 驳回日志 API |
| L502-810 | **POST /api/demands 提交主逻辑**（拆行 + 警告 + need_confirm）|
| L808-950 | 作废 / 批量作废 / 撤回报备 |
| L948-1300 | 报备结果回填 / 批量回填 / 导入 / 回传历史 / 状态汇总 |
| L1623-1700 | 批次列表 + 端口分配回传 apply |
| L1860-2090 | sync-from-expansion 工具拆行同步 |
| L2247+ | admin 校验 / docx→PDF / 承诺函代理 |

## 跑起来

```bash
cd workspace/projects/baobei-demand-hub
python3 -u backend/app/app.py      # 5000 端口
```
前端改 frontend/index.html 即可，Flask 直接 serve、无需重启。DB 在 backend/data/baobei.db（不要删）。

---

生成时间：2026-11-23