---
name: windows-qt-app-release
description: 用于开发、优化、打包、记录和发布 Windows Qt/PySide6 桌面程序。适用于用户要求创建或继续开发 Windows Qt 应用、优化表格密集型桌面 UI、沉淀通用 UI 组件配置、使用 PyInstaller 打包 exe 或安装包、生成真实程序 README 动图、编写需求设计文档、上传源码和 GitHub Release 资产等场景。
---

# Windows Qt 应用开发发布

## 工作流程

用于 Windows Qt 桌面程序，尤其适合财务、运营、数据核对、后台工具等表格密集型业务应用。

1. 先梳理业务流程和输入输出，发布前把需求说明和设计方案写入仓库。
2. 检查现有项目结构、数据目录、打包脚本、README、Release 说明和 `.gitignore`。
3. 按项目既有风格实现 UI、数据库和业务逻辑，不做无关重构。
4. 用语法检查、Qt 启动检查和关键流程测试验证改动。
5. README 动图必须来自真实程序窗口截图或录屏，不要用手绘示意图冒充实际界面。
6. 用 PyInstaller 打包主程序 exe；需要给用户安装体验时，再打包安装向导。
7. 更新 README、CHANGELOG、Release Notes、需求设计文档。
8. 上传源码和 Release 资产到 GitHub，排除本地数据库、构建缓存、用户数据。

详细检查清单见 `references/windows-qt-release-checklist.md`。

## UI 设计规则

- 财务和运营工具优先采用“店铺/主体优先”的布局，先选对象，再看汇总和明细。
- 表格界面要密集、清晰、稳定，避免营销页式大卡片和装饰性布局。
- 用户需要调节表格和控制区高度/宽度时，使用 `QSplitter`。
- 同一个表格的分页按钮和行操作按钮放在同一行，距离不要被 stretch 拉太远。
- 筛选条件过多时，单独放一行并允许换行显示。
- 中文业务应用的按钮必须使用明确中文动作，例如 `开始导入`、`保存设置`、`退出`、`清除筛选条件`。
- 大表不要在每次选择筛选条件时立即刷新，先暂存条件，点击“开始筛选”后再查询。

## 性能规则

- 明细大表默认分页加载，不要一次性渲染整月全部数据。
- 排序、筛选、总数统计尽量下推到 SQLite 或数据库层。
- 筛选下拉候选值使用缓存，避免频繁 `distinct` 大表。
- 导入、导出要分批处理，并在状态栏显示处理进度。
- 切换店铺或报表类型时先清空旧数据，等用户点击查询再加载。
- 填充大表前关闭表格更新和排序，填充完再恢复。

## 打包规则

- 主程序使用 PyInstaller，常用参数：`--windowed --onefile --icon app_icon.ico`。
- exe 名称可以是中文，但 GitHub Release 资产名优先使用英文，避免下载链接乱码。
- 如果用户需要“一步步安装、桌面图标、开始菜单、卸载入口”，要提供安装包。
- 发布 zip 中只放安装包、单文件 exe、Release Notes 等必要文件，不放 `data/`、`build/`、`dist/` 缓存和 SQLite 数据库。
- 若覆盖 exe 失败，优先判断旧 exe 是否正在运行或被系统锁定；必要时改用带版本号的新文件名。

## GitHub 发布规则

- 不提交用户数据、本地数据库、导入样例里的敏感数据。
- 发布前更新 README、CHANGELOG、Release Notes、需求设计文档。
- 上传 Release 资产后要再读回资产列表，确认 setup、exe、zip 的文件名和大小。
- 如果 GitHub 对中文资产名处理异常，补传一个英文文件名资产。
- 若没有本机 `git` 或 `gh`，可以使用 GitHub API 创建仓库、提交 tree、创建 Release、上传资产。

## 验证命令

至少运行 Python 语法检查：

```powershell
python -B -c "import ast,pathlib; ast.parse(pathlib.Path('qt_app.py').read_text(encoding='utf-8')); print('AST_OK')"
```

Qt 启动检查可以使用 `QT_QPA_PLATFORM=offscreen`。生成真实 README 截图或 GIF 时，必须使用 Windows 原生窗口模式，否则中文字体可能显示成方块。
