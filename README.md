# 员工每日工作量统计系统 - Web 版本

## 🌐 在线演示

完全基于浏览器的工作量统计系统，无需安装任何软件！

## ✨ 特性

- ✅ **纯前端实现**：无需后端服务器，完全在浏览器中运行
- ✅ **跨平台**：支持 Windows、Mac、Linux，任何有浏览器的设备
- ✅ **隐私安全**：数据完全在本地处理，不上传到服务器
- ✅ **实时处理**：即时生成报表，立即下载
- ✅ **美观界面**：现代化设计，响应式布局
- ✅ **智能计算**：基于时间戳的 EWH 计算，自动计算所有 UPH

## 📦 功能

1. **上传表格**：支持 Picking 和 Packing Excel 文件
2. **填写信息**：日期、工作时间、Preshipment 数据
3. **自动计算**：
   - EWH（有效工作时长）- 基于操作时间戳
   - Packing UPH = pack ÷ Packing_EWH
   - Picking UPH = pick ÷ Picking_EWH
   - Preship UPH = Preshipment ÷ 总EWH
4. **生成报表**：自动生成 Excel 文件并下载

## 🚀 部署到 GitHub Pages

### 方法一：使用 GitHub 网页界面

1. **创建 GitHub 仓库**
   - 登录 GitHub
   - 点击 "New repository"
   - 输入仓库名（例如：`work-report-system`）
   - 选择 "Public"
   - 点击 "Create repository"

2. **上传文件**
   - 在仓库页面点击 "Add file" → "Upload files"
   - 将以下文件拖入：
     - `index.html`
     - `style.css`
     - `app.js`
   - 点击 "Commit changes"

3. **启用 GitHub Pages**
   - 进入仓库的 Settings
   - 左侧菜单点击 "Pages"
   - 在 "Source" 下拉菜单中选择 "main" 分支
   - 选择 "/ (root)" 文件夹
   - 点击 "Save"
   - 等待几分钟，页面会显示访问链接

4. **访问网站**
   - 链接格式：`https://你的用户名.github.io/仓库名/`
   - 例如：`https://username.github.io/work-report-system/`

### 方法二：使用命令行

```bash
# 1. 初始化 Git 仓库
cd web
git init

# 2. 添加文件
git add .
git commit -m "Initial commit: Work Report System"

# 3. 连接到 GitHub 仓库
git remote add origin https://github.com/你的用户名/仓库名.git

# 4. 推送代码
git branch -M main
git push -u origin main

# 5. 在 GitHub 网站上启用 Pages（参考方法一的步骤3）
```

## 📖 使用说明

### 本地运行

1. 直接双击 `index.html` 文件
2. 或者使用简单的 HTTP 服务器：

```bash
# Python 3
python -m http.server 8000

# Node.js (需要先安装 http-server)
npx http-server
```

然后访问 `http://localhost:8000`

### 在线使用

1. 访问部署好的 GitHub Pages 链接
2. 上传 Picking 和 Packing 表格文件
3. 填写日期、工作时间等信息
4. （可选）添加 Preshipment 员工数据
5. 点击"生成每日工作报表"按钮
6. 自动下载生成的 Excel 文件

## 📋 数据格式要求

### Picking 表格
- **L 列（索引 11）**：拣货员姓名
- **M 列（索引 12）**：操作时间

### Packing 表格
- **V 列（索引 21）**：操作人姓名
- **X 列（索引 23）**：操作时间
- **H 列（索引 7）**：扫描件数（打包数）

## 🔧 技术栈

- **HTML5**：页面结构
- **CSS3**：样式和布局
- **原生 JavaScript**：核心逻辑
- **SheetJS (xlsx.js)**：Excel 文件读写
- **CDN**：使用 CDN 加载 xlsx 库，无需本地依赖

## 🌟 优势

### vs Python 桌面版
- ✅ 无需安装 Python 和依赖包
- ✅ 无需下载软件
- ✅ 跨平台，任何设备都能用
- ✅ 更新方便，只需刷新页面
- ✅ 可以分享链接给其他人使用

### 适用场景
- ✅ 多人使用
- ✅ 不同设备使用
- ✅ 不想安装软件
- ✅ 需要远程访问

## 📱 移动端支持

完全支持手机和平板浏览器：
- 响应式设计
- 触摸友好
- 自适应布局

## 🔒 隐私保护

- ✅ 所有数据处理都在本地浏览器完成
- ✅ 不上传任何数据到服务器
- ✅ 不收集任何个人信息
- ✅ 完全离线可用（首次加载后）

## 🐛 问题反馈

如有问题，请在 GitHub Issues 中反馈。

## 📄 许可证

MIT License

---

**享受无需安装的便捷体验！** 🎉

