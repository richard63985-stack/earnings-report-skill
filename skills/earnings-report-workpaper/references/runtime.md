# Windows 与 macOS 运行

Python 3.11 或更新版本。建议使用 python.org 的 Python 与独立虚拟环境。以下命令均从仓库根目录执行。

Windows PowerShell 7：

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r skills/earnings-report-workpaper/requirements.txt
.\.venv\Scripts\python.exe -m playwright install chromium
```

macOS：

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r skills/earnings-report-workpaper/requirements.txt
.venv/bin/python -m playwright install chromium
```

后续 `python` 指上面虚拟环境中的 Python，不要求激活环境。所有路径支持空格和中文，命令中的路径必须加引号。配置内使用 `/` 和相对路径即可跨平台迁移。

## 依赖与可选设置

- 默认使用 Playwright Chromium；可通过 `REPORT_BROWSER` 指定浏览器可执行文件，不绑定 Edge 或 Chrome 的 Windows 路径。
- 中文字体自动查找 Windows 微软雅黑/宋体、macOS 苹方/黑体；找不到就报错，避免悄悄生成方框字。可设置 `REPORT_FONT` 为自己的中文字体文件。字体不随仓库分发。
- Word/Excel 文件生成不要求安装 Microsoft Office。PDF 预览可安装 LibreOffice，macOS 自动查找 `/Applications/LibreOffice.app/Contents/MacOS/soffice`；其他位置用 `REPORT_SOFFICE` 或 `--soffice` 指定。
- Windows Word COM 仅用于可选导出，要求 Word 和 PowerShell 7；macOS 自动走 LibreOffice，不依赖 COM。机构模板仍需人工在实际 Word 中确认版式。
- HTML 来源应先保存成自包含的本地文件。截图阶段阻止 HTTP 请求并清除原脚本；外部 CSS/图片不会加载，复杂页面宜先保存 PDF 或人工截图。

## 脚本顺序

1. `make_template.py --output "项目/template.docx"`：无正式模板时创建中性草稿模板。用户模板需包含标题、副标题、报告日期、投资评级、评级变动、事件公告、摘要、风险提示等别名的 Word 内容控件；摘要默认需容纳八个段落。未知模板不要直接套填。
2. `fill_report_template.py --markdown "项目/report.md" --template "项目/template.docx" --output "项目/report-v1.docx" --date "日期" --rating "用户确认的评级" --rating-change "用户确认的变动"`。可附 `--export-pdf "项目/report-v1.pdf" --render-pages "项目/preview" --pdf-engine auto`。失败返回非零码，不能当成验收完成。
3. 用户确认后运行 `freeze_word.py "项目/report-v1.docx" --output "项目/frozen-v1.json"`。人工将真实确认记录及哈希填写到配置。
4. `create_visual_workpaper.py --config "项目/workpaper-v1.json" --validate-only`，再去掉 `--validate-only` 生成。已存在的 Excel 拒绝覆盖，应选择新的版本。
5. 视觉检查过程图片以及完整 Word/Excel 页面；用户项目中可保留过程 manifest 作复核。

## 离线演示

`make_demo.py --output-dir "generated/demo"` 创建纯虚构资料、来源 CSV、Word 和底稿配置。
然后运行 `create_visual_workpaper.py --config "generated/demo/workpaper.json"`。
此演示中的确认标志仅供自动测试，不是对真实研报的用户授权。
