# 财报点评与配套底稿 Skill

将中文财报点评写作、Word 成稿与逐句对照 Excel 底稿串成一个流程。支持 Windows 和 macOS，也支持只写研报或只制作底稿。

**流程：官方材料 → 正文与来源草稿 → Word → 用户确认 → Excel 截图底稿。**

包含 Word 内容控件填充、中文正文截图、HTML 关键词高亮、PDF/图片裁剪、Excel 左右对照排版、Word 哈希校验、中性模板和纯虚构演示生成器。真实公司材料、机构模板、历史研报、密钥和个人目录均不包含在仓库内。

英文材料使用“研报正文｜原文截图｜中文辅助译文”三栏，便于审核人员对照；中文材料保留两栏。模型负责忠实翻译，脚本检查缺项并排版，不调用额外的翻译服务。

## 安装

下载或克隆本仓库，把 `skills/earnings-report-workpaper` 整个文件夹复制到 Codex 的 Skills 目录（`CODEX_HOME/skills`，未设置时为用户主目录下的 `.codex/skills`）。必须连同 scripts、references、agents 和 requirements 一起复制。

依赖安装及两套系统的命令见 [运行说明](skills/earnings-report-workpaper/references/runtime.md)。安装后在新任务中使用：

> 使用 $earnings-report-workpaper，根据这些公司材料写财报点评；我确认 Word 后，再制作对应 Excel 底稿。

也可以说：

> 使用 $earnings-report-workpaper，我已有确认后的 Word，只制作对应底稿。

## 范围与限制

- 脚本负责文件生成与结构核验，事实、观点、评级和证据支持度仍需研究人员审阅。
- 默认四个正文 sheet，逐句左右对照；来源 CSV/JSON 为过程文件。
- Word 变化会使底稿配置哈希失效，必须重新确认并同步来源。
- 没有提供正式模板时只能生成中性草稿，不冒充机构正式研报。
- macOS 使用中文系统字体和 Chromium；PDF 导出可使用 LibreOffice。不同 Office 渲染引擎存在排版差异，自动测试不能替代真实模板的人工验收。
- 仓库提供 Windows/macOS/Linux 的 GitHub Actions 测试；具体运行结果见 Actions，不把配置了测试等同于已经通过。

## 开发验证

```bash
python -m unittest discover -s tests -v
```

测试生成虚构来源、Word 和四张 Excel 表，覆盖正文变化、缺失来源及路径异常，默认执行真实浏览器截图。不会调用模型或下载真实研报。
