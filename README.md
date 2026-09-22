# 财报点评与配套底稿 Skill

将中文财报点评、Word 成稿与逐句对应的 Excel 截图底稿串成一个流程。支持 Windows 和 macOS，也可只写正文、给已有 Word 配底稿，或按明确授权批量制作。

**流程：官方资料 → 正文与逐句来源 → 模板内填充 Word → 用户确认或获授权的主控复核 → 锁定版本 → 四表截图底稿 → 检查实际交付件。**

最近更新：[2026-09-22 更新说明](CHANGELOG.md#2026-09-22)。

## 这套 Skill 做什么

- **保留模板**：默认填写标题、日期和正文，保留评级、副标题、作者、侧栏等固定内容；模板刷新遵循本次明确范围。
- **保持正文简洁**：结论先行，事实支持研究判断；核验提醒、计算过程和制作备注放入来源映射或底稿。
- **方便审核**：每个逐句组先列正文、期间、单位、核对数据和完整公式，再并排展示原文截图与忠实中文译文。跨期比较的两期证据在同组相邻展示。
- **同步版本**：锁定 Word 后生成底稿，记录文件哈希；返修使用新文件与新过程目录，最终取件依据复核清单。
- **支持批量**：用户明确授权主控复核后连续制作时，记录授权范围、复核者和准确 Word 版本，不再逐篇重复确认，也不冒充用户逐篇认可。

默认四张正文表：事件、投资要点第一段、投资要点第二段、投资要点第三段。英文或中英混合来源附中文辅助译文；译文、原始证据与计算说明分区放置。

仓库包含通用方法、生成脚本和虚构测试样例；不包含真实公司材料、机构模板、历史研报或登录凭据。

## 安装与更新

下载或克隆本仓库，将 `skills/earnings-report-workpaper` 整个文件夹复制到 Codex 的 Skills 目录：设置了 `CODEX_HOME` 时为其下的 `skills`，否则为用户主目录下的 `.codex/skills`。保留整个目录，包括 `scripts`、`references`、`agents` 和 `requirements.txt`。

更新已有安装前先备份同名 Skill，然后同步完整目录。不要只复制 `SKILL.md`，新规则依赖配套脚本。更新 Skill 不会自动改写或重新生成已有报告。

依赖和系统命令见 [运行说明](skills/earnings-report-workpaper/references/runtime.md)。先复用可用环境，实际需要生成时才补缺少的依赖。安装完成后在新任务中使用。

## 可以直接这样用

**完整流程：**

> 使用 $earnings-report-workpaper，根据这些官方材料写财报点评，沿用我的模板和样稿。我确认 Word 后，再制作对应 Excel 截图底稿。

**已有 Word：**

> 使用 $earnings-report-workpaper，以这份已确认的 Word 为准，只制作底稿。请把跨期数据、完整公式和两期原文证据集中放在同一核对组。

**授权批量：**

> 使用 $earnings-report-workpaper，按这份清单批量制作。我授权主控逐篇复核后连续制作 Word 和底稿；每篇记录准确版本及复核结果，完成后交我审阅。

## 使用入口

| 需要做的事 | 读取 |
| --- | --- |
| 任务分流、核心边界 | [Skill 入口](skills/earnings-report-workpaper/SKILL.md) |
| 写正文、保护模板 | [写作规则](skills/earnings-report-workpaper/references/writing.md) |
| 准备来源、配置与授权记录 | [来源与配置](skills/earnings-report-workpaper/references/evidence.md) |
| 组织四表、截图和译文 | [底稿规范](skills/earnings-report-workpaper/references/workpaper.md) |
| 交付前检查 | [验收规则](skills/earnings-report-workpaper/references/quality.md) |
| 批量、返修或跨电脑交接 | [批量与交接](skills/earnings-report-workpaper/references/batch-handoff.md) |

## 旧项目升级注意

- 未填写集中核对字段的旧配置继续使用旧排版。新写或返修时，每个点同时补齐 `review_inputs` 和 `review_result`；集中核对模式下旧 `evidence.calc` 的计算应迁移到 `point.review_result`。
- `--rating`、`--rating-change`、`--subtitle` 现在是可选项，省略时保留模板；旧模板的特定行高修复须明确使用 `--fix-legacy-row-height`。三个投资要点必须完整，独立投资建议可省略。
- 底稿默认四段按固定顺序匹配 Word；普通逐句模式必须按顺序覆盖整段。旧配置如有漏句，需要补齐对应证据。
- `word_confirmed: true` 只记录真实用户确认。授权主控复核采用独立配置与复核记录，具体字段见来源与配置说明。
- 新版输出 Excel 和过程目录都须另存；已有生成清单的过程目录不能重复使用。模板刷新即使正文不变也会改变 Word 哈希，应建立新旧版本关系。
- 历史交接包如有完整性清单，保留包内原 Skill 快照，另行更新本机 Skill；直接覆盖快照会导致旧包校验失败。

## 验证与限制

在仓库根目录、已准备好依赖和浏览器的环境中运行：

```bash
python -m unittest discover -s tests -v
```

完整测试已包含随Skill分发的回归检查。只运行该检查可用 `python skills/earnings-report-workpaper/scripts/check_regressions.py --browser`。

测试使用无关虚构公司，检查模板保护、制作授权与版本绑定、正文覆盖、旧配置、实际 Excel 单元格和嵌图，以及页面下方的浏览器截图。省略 `--browser` 可运行不启动浏览器的回归检查；使用 `--output-dir "generated/check"` 可保留文件供渲染，选择新的输出目录。

仓库保留 [Windows/macOS/Linux 自动检查](https://github.com/richard63985-stack/earnings-report-skill/actions/workflows/test.yml)。这些检查不启动原生 Microsoft Word/Excel；通过自动检查不代表机构模板、真实报告或 Windows Office 呈现已经验收。事实、观点、译文和来源支持度仍需研究人员核验，所有最终页面仍需查看。
