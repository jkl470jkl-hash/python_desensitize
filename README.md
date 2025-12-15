# 中文合同脱敏桌面工具（本地离线）

本项目提供一个面向法务同事的中文合同脱敏工具，完全本地运行，无需联网。支持 `.txt` 和 `.docx` 合同文件，批量拖拽到界面即可脱敏，生成可安全分享的文本、Word 版本以及映射表。

## 小白快速部署步骤（Windows）
1. 安装 Python 3.10+：到 <https://www.python.org/downloads/> 下载 Windows 安装包，全程下一步，记得勾选 **Add python.exe to PATH**。
2. 获取代码：
   - 不会用 Git 可以直接在 GitHub/企业盘下载项目 Zip，右键解压到任意文件夹（如 `D:\ContractRedactor`）。
3. 安装依赖（仅首次）：
   - 在项目目录空白处 **Shift + 右键 → 打开 PowerShell 窗口**，输入：
     ```bash
     pip install -r requirements.txt
     ```
   - 如果网络受限，可提前把离线 whl 包放到同目录，再运行上述命令。
4. 启动图形界面：双击 `run_gui.bat`，或在 PowerShell 执行：
   ```bash
   python redact_gui.py
   ```
   界面支持拖拽文件、选择输出目录、查看日志和进度。默认输出目录为项目下的 `output/`。

## 脱敏规则概览
- 角色主体：甲/乙/丙/丁方、买卖方等 → `[[PARTY_A_1]]` 等占位符
- 公司名称（带常见后缀）→ `[[PARTY_OTHER_1]]`
- 姓名：字段关键词 + 常见姓氏/称谓 → `[[PERSON_1]]`
- 联系方式：手机号、座机号、邮箱
- 账号/证件：银行账号、统一社会信用代码、身份证号
- 地址：字段名 + 地址成分双重校验
- 金额：人民币金额（含符号或大写金额）
- 日期：多种日期格式
- 编号：合同编号、项目编号

占位符与原值的映射会写入 `<文件名>_redaction_map.json`，便于回溯。

## 目录结构示例
```
config/
  address_component_words.txt
  address_field_words.txt
  company_suffix.txt
  person_family_names.txt
  role_words.txt
output/               # 默认输出位置
core_redact.py        # 脱敏核心逻辑 + 自检
redact_gui.py         # Tkinter + tkinterdnd2 图形界面
requirements.txt
README.md
```

## PyInstaller 打包示例
```bash
pyinstaller -F -w redact_gui.py --add-data "config;config" --add-data "output;output"
```
> 打包后可生成 `ContractRedactor.exe`，分发给无需安装 Python 的同事。

## 自检
在代码中提供 `run_self_test()`，可在终端直接运行：
```bash
python core_redact.py
```
终端会输出示例原文、脱敏结果和映射片段，便于快速验证规则。

## 常用文件/脚本
- `run_gui.bat`：双击即可启动 GUI，自动优先使用 `venv\Scripts\python.exe`，找不到 Python 会提示下载路径。
- `requirements.txt`：依赖清单，首次部署运行 `pip install -r requirements.txt` 即可。
- `config/*.txt`：词库，可直接用记事本修改。

## 常见问题
- 拖入 `.doc` 旧格式会提示请先另存为 `.docx`。
- 若无法打开输出目录（非 Windows），请手动前往 `output/` 查看。
- 词库在 `config/` 下，支持用记事本修改，无需改代码。
