# Xenta Field CRM (Static)

## 中文说明

### 1) 系统定位
Xenta Field CRM 是一个纯前端静态 CRM（HTML/CSS/JavaScript），可部署在 GitHub Pages。
系统用于销售线索管理、客户跟进记录、周报导出与本地备份。

### 2) 当前技术架构
- 前端：`index.html` + `styles.css` + `app.js`
- 周报导出：`js/excelExport.js`（SheetJS）
- 存储方式：浏览器 `localStorage`（无后端）
- 主存储 Key：`xenta_field_crm_v1`

### 3) 数据存储与安全策略
- 数据以结构 `{ version, lastUpdated, clients }` 保存到 `localStorage`
- 系统版本：`2.0.0`
- 启动时支持旧 key 迁移（向后兼容）
- 保存前自动备份，最多保留最近 5 份
- 保存后做校验，失败自动回滚到最新备份

### 4) 核心功能
- 新建 / 编辑 / 删除客户
- 客户字段支持：公司、联系人、职位、国家、阶段、评分相关字段、沟通记录等
- Email 按钮弹窗显示邮箱并支持一键复制（不再触发 mailto）
- JSON 导入导出（导入为按 `id` 合并，不覆盖全部）
- Weekly Excel 导出（中文字段，含公司、联系人、产品、销售管理、沟通历史）

### 5) 导入导出规则
- JSON 导出文件名：`Xenta_CRM_Backup_YYYYMMDD.json`
- Excel 导出文件名：`Xenta_Weekly_Report_YYYYMMDD.xlsx`
- 导入策略：
  - `id` 已存在 -> 更新该客户
  - `id` 不存在 -> 新增客户

### 6) 使用方式
1. 直接打开 `index.html` 使用（本地）
2. 或部署到 GitHub Pages 在线使用
3. 建议定期执行 JSON 导出做离线备份

### 7) 重要说明
- 当前是本地浏览器存储方案，不是云端数据库
- 更换浏览器/设备不会自动同步数据
- 跨设备建议使用：导出 JSON -> 在目标设备导入 JSON

### 8) 后续可升级方向
- 接入后端（Node.js + MongoDB / REST API）实现多端实时同步
- 增加账号权限、团队协作、审计日志
- 增加自动云备份与冲突合并策略

---

## English Documentation

### 1) Project Overview
Xenta Field CRM is a static, frontend-only CRM (HTML/CSS/JavaScript), suitable for GitHub Pages deployment.
It is designed for lead tracking, client follow-up records, weekly reporting, and local backup workflows.

### 2) Current Architecture
- Frontend: `index.html` + `styles.css` + `app.js`
- Weekly export: `js/excelExport.js` (SheetJS)
- Storage: Browser `localStorage` (no backend)
- Primary storage key: `xenta_field_crm_v1`

### 3) Data Persistence & Safety
- Data is stored as `{ version, lastUpdated, clients }` in `localStorage`
- System version: `2.0.0`
- Legacy key migration is supported on startup
- Auto backup before save (keeps latest 5 backups)
- Save verification with automatic rollback on failure

### 4) Core Features
- Create / update / delete client records
- Rich client fields: company, contact, position, country, stage, scoring-related fields, communication logs
- Email action opens a modal and supports copy-to-clipboard (no mailto trigger)
- JSON import/export (import uses merge-by-id, not full overwrite)
- Weekly Excel export with Chinese business columns and merged communication history

### 5) Import/Export Rules
- JSON export filename: `Xenta_CRM_Backup_YYYYMMDD.json`
- Excel export filename: `Xenta_Weekly_Report_YYYYMMDD.xlsx`
- Import behavior:
  - Existing `id` -> update record
  - New `id` -> append record

### 6) How to Use
1. Open `index.html` directly for local usage
2. Or deploy to GitHub Pages for web access
3. Regular JSON export is recommended for backup

### 7) Important Notes
- This is currently a local-browser storage solution, not a cloud database
- Data is not auto-synced across browsers/devices
- For cross-device usage: export JSON -> import JSON on target device

### 8) Future Upgrade Path
- Integrate backend (Node.js + MongoDB / REST API) for multi-device sync
- Add user roles, collaboration, and audit logs
- Add automatic cloud backups and conflict-resolution strategy
