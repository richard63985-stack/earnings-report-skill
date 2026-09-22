# 来源、授权与配置

来源CSV可用sentence_id、section、text、claim_type、source_file、location、quote、translated_fact、basis、formula、status。核心事实逐句对应一手材料；多来源可复用sentence_id。核验通过由实际审阅者记录，生成器不替代判断。

`freeze_word.py`提取Word哈希和段落，不自动授予制作许可。配置相对路径基于company_dir，company_dir基于配置所在目录。

## 制作许可

默认使用真实用户确认：`word_confirmed: true`，在项目中保存确认原话、准确版本与范围。已获得用户授权“主控复核后连续制作”时，保留 `word_confirmed: false` 并填写：

```json
{
  "generation_authorization": {
    "mode": "delegated_review",
    "authorized": true,
    "user_instruction": "用户本次真实授权原话",
    "scope": "本次明确公司/批次及适用动作",
    "review_record": "process/word_review.json"
  }
}
```

review_record由实际复核者填写 `passed: true`、`reviewer`、`word_sha256`及检查依据；哈希必须与当前Word一致。脚本只校验记录完整性与绑定关系，不能证明授权或审阅真实；作者不能为了通关伪造记录。用户仅要求Word时该授权不存在。

## 新底稿配置

以下只示意一个section，真实配置按顺序填写事件与三个投资要点四段：

```json
{
  "company_dir": "..",
  "word_path": "output/report_v02.docx",
  "word_confirmed": true,
  "word_sha256": "实际锁定Word的64位哈希",
  "output_xlsx": "output/workpaper_v02.xlsx",
  "process_dir": "process/workpaper_v02",
  "evidence": {
    "release": {
      "source": "sources/release.pdf",
      "label": "官方公告：本期及上期数据",
      "page_index": 0,
      "clip": [30, 30, 1100, 160],
      "source_language": "en",
      "source_excerpt": "Revenue: current quarter 126, prior-year quarter 100 (RMB million).",
      "translation_zh": "收入：本季度126、上年同期100，单位人民币百万元。"
    }
  },
  "sections": [{
    "sheet": "事件",
    "body": "本季收入1.26亿元，同比增长26%。",
    "overview_evidence": "release",
    "points": [{
      "text": "本季收入1.26亿元，同比增长26%。",
      "sentence_id": "S1-1",
      "review_title": "本期收入与同比",
      "review_inputs": "来源1：本季126、上年同期100，均为人民币百万元。",
      "review_result": "(126÷100−1)×100%=26%；126百万元=1.26亿元。",
      "evidence_ids": ["release"]
    }]
  }]
}
```

新配置每点填写review_inputs/review_result：直接披露写明属性与来源，计算写全部输入、公式及最后舍入，研究判断写支持事实及归纳属性。组内编号按evidence_ids顺序生成，核对文字与之对应；不把另一组的编号照抄过来。旧evidence.calc仅用于旧排版，新排版的计算统一写point.review_result。

以上数字和裁剪坐标为虚构示例；真实配置按对应页面确定，不能直接复用。

旧配置完全没有review字段时仍支持原排版；一旦使用集中核对字段，所有点都必须填齐。普通point_text模式下所有text按顺序拼接必须覆盖整段，不能只摘几句。整段高亮模式用left_mode=full_body_highlight和highlights，仍需人工核验完整覆盖。body_selector可定位唯一段落；有已知控件时脚本同时核对四段顺序，不能指向另一段后自称匹配。

## 来源字段

PDF：source、label、page_index、clip，可加真实印刷页码page；clip为2倍渲染像素。图片：source、label、clip；确实预高亮时才设pre_highlighted=true。HTML：source、label、唯一且完整的terms。ID只用英文字母、数字、下划线或短横线。label/page记录出处与必要拼接说明，不能冒充原文。

英文或mixed必须给source_excerpt和translation_zh。每条译文对应当前截图，旧配置返修时补齐语言；脚本不判断翻译语义。原文件不改，必要裁剪/高亮在项目过程目录进行并记录。

## 版本

输出Excel和process_dir均按版本另存。已有生成manifest的过程目录不能复用于新版本，防止混入旧截图或覆盖追溯材料。manifest保存Word、配置及输出Excel哈希；最终清单据复核记录确定，不自动推断“最大版本即最终”。
