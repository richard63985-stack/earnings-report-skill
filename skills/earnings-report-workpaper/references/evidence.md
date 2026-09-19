# 来源与配置

来源 CSV 建议字段：sentence_id、section、text、claim_type、source_file、location、quote、translated_fact、basis、formula、secondary_source、status。每条核心句至少一条记录；多来源可复用 sentence_id。status 由研究人员核验，不由生成器自动设为通过。

`freeze_word.py` 导出 Word 的 sha256 和正文段落，供人工选取，不能自动授予确认。

底稿 JSON（相对路径均相对于 company_dir，company_dir 相对于配置文件）：

```json
{
  "company_dir": ".",
  "word_path": "report.docx",
  "word_confirmed": true,
  "word_sha256": "实际确认后的64位哈希",
  "output_xlsx": "report-workpaper-v1.xlsx",
  "process_dir": "process-v1",
  "evidence": {
    "release": {"source": "release.html", "label": "release.html", "terms": ["Revenue increased"]}
  },
  "sections": [{
    "sheet": "事件",
    "body": "与最终 Word 完全一致的段落。",
    "overview_evidence": "release",
    "points": [{"text": "与最终 Word 完全一致的段落。", "evidence_ids": ["release"]}]
  }]
}
```

实际财报点评填齐四个 section。body 也可改为唯一的 body_selector.prefix 或 contains；应检查命中的是目标正文，不是摘要或模板示例。

PDF evidence：source、label、page_index、clip，可加 page（印刷页码）及 calc（计算说明）。图片 evidence：source、label、clip；已人工裁剪高亮的图可用 pre_highlighted=true。原文件始终不改。evidence ID 限英文、数字、下划线、短横线。

英文证据同时填写以下字段（适用于 HTML、PDF 和图片）：

```json
{
  "source_language": "en",
  "source_excerpt": "Gross margin was 70%, up 5 percentage points year over year.",
  "translation_zh": "毛利率为70%，同比提升5个百分点。"
}
```

source_excerpt 对应截图中的证据原文，translation_zh 是忠实翻译，不是研报结论。en/en-US/en-GB 或 mixed 来源必需中文译文；任何译文都必须有对应原文摘录。中文来源可用 zh 并省略译文。脚本不检验翻译语义正确性，交付前由 agent 逐项核对。

points 默认使用 text；整段高亮模式设置 left_mode=full_body_highlight 与 highlights。每一点必须有 evidence_ids，且原文必须出现在对应的 Word 段落中。不要把释义改写混入底稿左侧。
