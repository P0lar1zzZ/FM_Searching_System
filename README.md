# 防火门一键搜取 - 安装指南

**GitHub 项目：** https://github.com/P0lar1zzZ/FM_Searching_System.git

## 📦 项目结构

```
fm_serching_system/
├── install.bat              # Windows 安装脚本（一键配置环境）
├── generate_icon.py         # 图标生成脚本
├── fm_searching             # 防火门搜取核心脚本（Python）
├── fm_icon.ico             # 桌面快捷方式图标
└── README.md               # 本文件
```

## 🚀 快速开始（3步搞定）

### 第一步：运行安装脚本
1. **下载整个 `fm_serching_system` 文件夹**到你的本地电脑
2. **进入文件夹**，找到 `install.bat`
3. **双击运行** `install.bat`
   - 自动检查 Python（没有会提示下载）
   - 自动安装 pywin32、pandas、openpyxl
   - 自动在**桌面**创建快捷方式 ✨

### 第二步：生成图标（可选，但推荐）
如果想要漂亮的图标，运行：
```bash
python generate_icon.py
```
这会在当前目录生成 `fm_icon.ico`（外圆白色+中心黑色圆）

### 第三步：使用工具
1. 在 AutoCAD 中打开 DWG 文件
2. **双击桌面**的"防火门一键搜取"快捷方式
3. 脚本自动扫描并生成 Excel 报表 📊

---

## 📋 install.bat 做了什么？

| 步骤 | 功能 | 说明 |
|------|------|------|
| 1️⃣ | 检查 Python | 如果没装 Python，会打开官方下载页面 |
| 2️⃣ | 安装依赖 | 使用清华 TUNA 镜像，一键安装 3 个包 |
| 3️⃣ | 下载脚本 | 从 GitHub 下载最新版本（使用 fastgit 加速） |
| 4️⃣ | 创建快捷方式 | 自动在桌面创建精美快捷方式 |

### 第三步详解：自动更新脚本

脚本会自动从 GitHub 仓库下载最新版本：
- **GitHub 地址**：https://github.com/P0lar1zzZ/FM_Searching_System.git
- **加速镜像**：fastgit（国内快速访问）
- **备用源**：GitHub 官方源（如加速镜像超时）
- **智能备份**：旧版本自动保存为 `.backup`

即使下载失败，也会保留本地脚本继续使用！

### 安装的包
- **pywin32** - AutoCAD COM 接口调用
- **pandas** - 数据处理和 Excel 导出
- **openpyxl** - Excel 文件支持

### 镜像源
默认使用：**清华大学 TUNA 镜像** https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple

如果超时，脚本会建议你用：
- 阿里云: https://mirrors.aliyun.com/pypi/simple/
- 腾讯云: http://mirrors.cloud.tencent.com/pypi/simple

---

## 🎨 图标说明

运行 `install.bat` 后，脚本会自动查找 `fm_icon.ico`

**图标设计**：
- 🟡 外圆：白色背景 + 灰色边框
- ⚫ 中心：黑色大圆
- ⚪ 中点：白色小点

如果要自定义，运行 `generate_icon.py` 重新生成

---

## ⚠️ 常见问题

### Q1: 提示"Python 未安装"怎么办？
**A:** 脚本会自动打开 Python 官网，请下载 Python 3.8+ 版本，**务必勾选"Add Python to PATH"**

### Q2: pip 安装超时？
**A:** 脚本已使用清华镜像加速。如仍超时，手动运行：
```bash
pip install -i https://mirrors.aliyun.com/pypi/simple/ pywin32 pandas openpyxl
```

### Q3: 快捷方式创建失败？
**A:** 不影响使用。你可以手动在文件夹中运行 `fm_searching`，或用 VBS 手动创建：
```vbs
Set objWS = CreateObject("WScript.Shell")
Set objLink = objWS.CreateShortcut("C:\Users\你的用户名\Desktop\防火门一键搜取.lnk")
objLink.TargetPath = "python.exe"
objLink.Arguments = "D:\fm_serching_system\fm_searching"
objLink.Save
```

### Q4: 脚本运行时提示"找不到 AutoCAD"？
**A:** 确保 AutoCAD 已打开且有活跃文档。重新打开 AutoCAD 后再试

### Q5: 扫不到布局里的防火门规格？
**A:** 脚本已支持布局扫描。如仍有遗漏，检查：
- 图层是否被冻结？
- 标注是否被"炸开"（分解成文字）？
- XRef 参照是否完整附加？

### Q6: 如何手动更新脚本？
**A:** 有三种方式：

**方式 1：自动更新（推荐）**
- 直接再次运行 `install.bat`，它会自动检查并从 GitHub 下载最新版本

**方式 2：Git 克隆**
```bash
git clone https://github.com/P0lar1zzZ/FM_Searching_System.git
```

**方式 3：手动下载单个文件**

从 GitHub 下载这些文件替换本地版本：
- https://raw.githubusercontent.com/P0lar1zzZ/FM_Searching_System/main/fm_searching
- https://raw.githubusercontent.com/P0lar1zzZ/FM_Searching_System/main/generate_icon.py
- https://raw.githubusercontent.com/P0lar1zzZ/FM_Searching_System/main/install.bat

**国内加速（推荐）**
将上面的 URL 中的 `github.com` 替换为 `raw.fastgit.org`，速度会快 10 倍！

---

## 🔧 手动安装依赖（如果 install.bat 失败）

```bash
# 打开 cmd，逐个运行：
pip install pywin32
pip install pandas
pip install openpyxl
```

---

## 📞 技术支持

**GitHub Issues:** https://github.com/P0lar1zzZ/FM_Searching_System/issues

如有问题，请检查：
1. Python 版本 >= 3.8
2. AutoCAD 已打开且文件已保存
3. 外部参照（XRef）文件完整
4. DWG 文件中的防火门规格包含"FM"标识

---

**祝你使用愉快！** 🎉
