# 官渡排表工具使用说明

## 方式1：可执行文件（推荐给伙伴）

### macOS用户

1. **下载文件**
   - 文件：`guandu-app-macos.zip`
   - 大小：约17MB

2. **解压并运行**
   ```bash
   # 解压
   unzip guandu-app-macos.zip
   
   # 进入目录
   cd dist
   
   # 添加执行权限（首次需要）
   chmod +x guandu-app
   
   # 运行
   ./guandu-app
   ```

3. **访问应用**
   - 自动打开浏览器：http://localhost:5001
   - 或手动输入网址

4. **停止服务**
   - 终端按 `Ctrl+C`

### Windows用户

⚠️ 当前打包版本仅支持macOS。Windows用户可选择：
- 方式2：本地Python运行
- 方式3：访问在线版本（部署后）

---

## 方式2：本地Python运行

### 环境要求
- Python 3.9+
- pip

### 安装步骤

```bash
# 1. 进入项目目录
cd ~/.qclaw/workspace-agent-29b6e205

# 2. 安装依赖
pip3 install flask flask-cors openpyxl

# 3. 运行应用
python3 app.py

# 4. 打开浏览器
# http://localhost:5001
```

---

## 方式3：在线版本（Render部署）

### 部署步骤

1. **注册Render账户**
   - 访问：https://render.com
   - 使用GitHub账号登录

2. **创建Web Service**
   - 点击 "New +" → "Web Service"
   - 连接GitHub仓库：`Kelley99/guandu-scheduler`
   - Region: Singapore（最近）
   - Branch: main
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
   - Instance Type: Free

3. **等待部署**
   - 首次部署约2-3分钟
   - 完成后获得永久链接：`https://guandu-scheduler.onrender.com`

4. **分享链接**
   - 直接发给伙伴访问
   - 无需安装任何软件

---

## 功能说明

### 主要功能
- ✅ 上传凌霄数据统计表（Markdown格式）
- ✅ 自动分配官渡场次（82人 → 3组）
- ✅ 导出分配结果（Excel格式）
- ✅ 支持手动调整

### 数据文件
- 统计表：`knowledge-base/凌霄数据统计表26.3.30.md`
- 官渡表：`knowledge-base/凌霄官渡26.md`

---

## 常见问题

### Q: macOS提示"无法打开，因为无法验证开发者"
A: 右键点击 → 打开 → 仍要打开

### Q: 端口5001被占用
A: 修改 `app.py` 最后一行的端口号

### Q: 上传文件失败
A: 确保文件格式为 `.md` 或 `.xlsx`

---

## 技术支持

- GitHub仓库：https://github.com/Kelley99/guandu-scheduler
- 问题反馈：在GitHub提Issue
