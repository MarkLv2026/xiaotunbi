# 小豚BI 代码推送指南

## 当前状态

本地代码已完成以下改动（共2个commit待推送）：

### Commit 1: `ad8673e` - 推广分析+时间段对比优化
- 推广分析 Tab：日度明细表增加"费率"列
- 推广分析 Tab：新增"单品推广分析"板块（TOP10花费图+ROI图+数据表+CSV下载）
- 时间段对比 Tab：新增推广数据对比（7个KPI卡片+对比详情表）

### Commit 2: `e761d3c` - 表格美化+智能诊断推广分析
- 推广分析 Tab：店铺/渠道/单品推广矩阵表格美化（HTML居中样式）
- 推广分析 Tab：三个表格增加同比数据列（花费同比、直接ROI同比、CPC同比、转化率同比）
- 时间段对比 Tab：销售对比和推广对比分开展示，加独立标题
- 智能诊断 Tab：新增"推广数据诊断 & 联动分析"板块（5个KPI卡片+6类推广诊断建议）

---

## 推送方案（选一种）

### 方案A：GitHub 网页直接上传（最稳，推荐）

由于当前网络环境 GitHub 443 被阻断，建议：

1. **连接手机热点** 或 **公司外网/VPN**
2. 打开 https://github.com/MarkLv2026/xiaotunbi/edit/main/app.py
3. 把本地 `app.py` 内容全选复制粘贴上去
4. 提交信息写：`feat: 推广分析表格美化+同比；时间段对比拆分；智能诊断增加推广联动分析`
5. 点 Commit changes

> Streamlit Cloud 会自动检测提交并重新部署（约1-3分钟）

### 方案B：使用 GitHub Desktop（图形化）

1. 下载 GitHub Desktop: https://desktop.github.com/
2. 登录你的 GitHub 账号
3. Clone `MarkLv2026/xiaotunbi`
4. 把本地 `app.py` 复制覆盖过去
5. 在 GitHub Desktop 中填写提交信息，点 Commit & Push

### 方案C：生成新 Token 命令行推送

1. 打开 https://github.com/settings/tokens/new
2. 勾选 `repo` 权限，生成新 token
3. 连接手机热点
4. 在本地仓库目录运行：
   ```bash
   git push https://TOKEN@github.com/MarkLv2026/xiaotunbi.git main
   ```

---

## Streamlit Cloud 休眠解决方案

### 问题原因
Streamlit Cloud 免费版在 7 天无访问后会进入休眠状态，首次访问需要点击 "Yes, get this app back up!" 唤醒。

### 解决方案

#### 方案1：UptimeRobot 自动唤醒（推荐，免费）

1. 注册 https://uptimerobot.com/（免费版即可）
2. 添加 New Monitor：
   - Monitor Type: HTTP(s)
   - Friendly Name: `小豚BI Keep Alive`
   - URL: `https://xiaotunbi-tmfhdkek237cxntwknq6ny.streamlit.app/`
   - Monitoring Interval: `5 minutes`（免费版最短）
3. 保存即可

> UptimeRobot 每5分钟访问一次，保持应用活跃

#### 方案2：自有服务器/云函数定时触发

如果有服务器或云函数（如腾讯云函数、阿里云函数计算），设置定时任务每10分钟访问一次 URL。

已提供 `keep_alive.py` 脚本，可配合 cron 使用：
```bash
*/10 * * * * /usr/bin/python3 /path/to/keep_alive.py
```

#### 方案3：手动唤醒（临时）

每次访问前点击 "Yes, get this app back up!" 按钮，等待 10-30 秒即可。

---

## 跨网络访问优化

如果不同网络（如公司内网 vs 家庭网络）访问有困难：

1. **确认 URL 可访问性**：`https://xiaotunbi-tmfhdkek237cxntwknq6ny.streamlit.app/` 是全球可访问的，不需要翻墙
2. **如果某些网络打不开**：可能是 DNS 或防火墙问题，尝试：
   - 切换 4G/5G 热点
   - 使用公共 DNS（如 8.8.8.8 或 114.114.114.114）
3. **长期方案**：可考虑部署到国内平台（如腾讯云 CloudBase、阿里云函数计算），但需额外配置

---

## 文件清单

推送前确认以下文件已更新：

| 文件 | 说明 |
|------|------|
| `app.py` | 主程序（核心改动） |
| `.streamlit/config.toml` | Streamlit 配置 |
| `keep_alive.py` | 唤醒脚本（可选） |
| `PUSH_GUIDE.md` | 本指南 |
