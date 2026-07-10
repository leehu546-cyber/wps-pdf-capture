# PDF 截图助手

在 Windows 上配合 WPS 使用的小工具：将 PDF 或 Word 文档的指定页导出为高清 PNG，支持单页截图与批量截图。

## 功能

- **PDF 截图**：通过 PyMuPDF 按页导出
- **Word 截图**：通过 WPS COM 按页导出（无需在 WPS 中逐页翻页）
- **自动识别文档**：读取 WPS 当前打开的文档，或在「搜索目录」下递归查找
- **单页截图**：填写页码后点击「截当前页」
- **批量截图**：填写「从 X 到 Y」后点击「批量截图」
- **按文档归类保存**：`PDFPageCaptures\{文档名}\{文档名}_第009页.png`，同页重复截图会覆盖

## 环境要求

- Windows 10 / 11（64 位）
- 已安装 **WPS**（W365 等）
- 使用 exe 版时**无需安装 Python**

## 下载与安装

在 [Releases](https://github.com/leehu546-cyber/wps-pdf-capture/releases) 下载（若尚未发布，可在本地执行 `build_release.bat` 自行打包）：

| 方式 | 说明 |
|------|------|
| **安装版** | 运行 `PDF截图助手_setup.exe`，按向导安装 |
| **便携版** | 解压 `PDF截图助手_便携版.zip`，双击 `安装.bat` 或 `PDF截图助手.exe` |

## 使用步骤

1. 启动「PDF 截图助手」小窗口（置顶显示）
2. 在「搜索目录」设置存放 docx/pdf 的工作文件夹（默认 `d:\工作`，可修改并自动记住）
3. 用 WPS 打开目标文档，点击「刷新识别」
4. 填写页码：
   - **单页**：填「页码」→ 点绿色「截当前页」
   - **批量**：填「从」「到」→ 点「批量截图」
5. 截图保存在程序目录下的 `PDFPageCaptures\` 文件夹

## 从源码运行

```bash
pip install pymupdf pywin32 pillow
python _wps_pdf_capture_test.py
```

或双击 `启动.bat`（需本机已安装 Python）。

## 打包发布

```bash
build_release.bat
```

生成：

- `release\PDF截图助手_便携版.zip`
- `release\PDF截图助手_setup.exe`（需安装 Inno Setup 6）

## 项目结构

```
├── _wps_pdf_capture_test.py   # 主程序
├── build_release.bat            # 一键打包脚本
├── PDF截图助手.spec             # PyInstaller 配置
├── installer.iss                # Inno Setup 安装包配置
├── 启动.bat                     # 源码启动
├── 安装.bat                     # 便携版简易安装
└── 使用说明.txt
```

## 许可证

MIT
