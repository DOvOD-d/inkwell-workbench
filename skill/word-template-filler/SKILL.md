---
name: word-template-filler
description: Fill any Word (.docx) template, form, task book, report, contract, or school form without placeholders. Use when the user asks to complete/fill a Word document from supplied information, or asks why a Word template cannot be recognized. Reads tables, inline blanks, underlines, bracket prompts, content controls, legacy form fields and headers/footers automatically, then writes values while preserving all original text and layout.
---

# Word 模板填写（无占位符）

## 什么时候用

用户说“帮我填这份 Word / 把资料填进这个模板 / 这份任务书怎么填 / 识别一下这个表”等，并且有 `.docx` 文件时使用。不要要求用户在 Word 里添加 `{{ }}`、书签或任何标记——本技能直接读取原文档结构。

## 工作流（必须按顺序）

> 下面的 `scripts/...` 指本技能目录下的脚本。若当前工作目录不是技能根目录，请换成技能目录的绝对路径（例如 `C:\Users\...\skills\word-template-filler\scripts\analyze.py`）。

1. **分析模板**（只读，不改原文件）：
   ```bash
   python scripts/analyze.py "<原文件.docx>" --out "<工作目录>"
   ```
   - 依赖（python-docx + lxml）已内置在 `scripts/vendor/`，直接用系统 Python 3.13 即可运行；换到其他 Python 版本若报 lxml 错误，再 `python -m pip install -r requirements.txt`。
   - 生成 `<工作目录>/template.docx` 和 `<工作目录>/fields.json`。
   - `fields.json` 里包含 `fields`（key/label/type/group/required/where/context/sample）、`outline`（文档大纲）和 `values`（空值模板）。
   - 先读 `fields.json`，不要凭文件名猜字段。

2. **对齐用户资料与字段**：
   - 优先按 `label` 匹配；同义字段（姓名/学生姓名、学院/所在学院、题目/论文题目）可以复用同一个值。
   - 同一字段在文档里出现多次时，只填一次，生成时会一起替换（key 相同）。
   - `outline` 用来理解栏目上下文；`sample` 是模板里的示例答案，可参考语体，但不要照抄。
   - 长文本字段（multiline）按栏目要求写完整段落；不要编造未提供的人名、单位、日期、数据。
   - **默认所有字段都不强制填写**。只有用户明确要求某项必须填时，才把它当作必填，并一次性列出缺失项确认。
   - 用户没提到、资料里也没有的字段直接留空，不要为了“填满”而编造。

3. **填写**：
   ```bash
   python scripts/fill.py "<工作目录>/fields.json" "<values.json>" --out "<输出.docx>"
   ```
   - `values.json` 是 `{ "字段key": "值" }` 的 JSON 对象；也接受 `{ "values": {...} }`。
   - 键可以用 `key` 或中文 `label`。
   - 默认允许留空；只有用户要求“必填项不能空”时才加 `--strict`。

4. **验收**：
   ```bash
   python scripts/verify.py "<输出.docx>" --values "<values.json>"
   ```
   - 检查填写值是否真的出现在文档里；未填字段可能保留模板原占位提示，默认只提示，不判失败。
   - 用户要求“所有项都必须填”时，给 verify 加 `--strict`。
   - 再人工看一眼页数和分页；输出是普通 .docx，可继续编辑。

## 硬性规则

- **永远不修改原文件**：分析只读；生成写到新路径。
- **不改固定内容**：只写识别出的填写位；标题、说明、签名/日期、表格线、字体、行距、页眉页脚保持不变。
- **不虚构事实**：资料没有的内容写“待补充”或空着，并向用户说明；生成时空值默认保留模板原文/示例，不会清空。
- **不强制填满**：大部分字段是可选；只有用户明确要求的项才当必填。
- **不要输出占位符**：生成的文档里不应出现 `{{ }}`、`【请填写】` 等提示；如出现，说明该字段未被识别，回到步骤 1 检查。
- **旧 .doc 不支持**：先让用户用 Word 另存为 `.docx`。
- 扫描版 PDF 不在本技能范围；需要先 OCR 或转成可编辑 Word。

## 常见问题

- **字段太多/名称太泛**（如“控件1”“字段2”）：结合 `where` 和 `context` 推断；仍不确定就保留原名并问用户。
- **表格里的同一个标签出现多次**：key 会不同（如“专业”“专业2”），分别填写；确认是否确实要填不同值。
- **页眉/页脚字段**：`where` 会写“页眉/页脚”；同样填写即可，生成时会写入对应部件。
- **内容控件/表单域**：`config.fields[].control` 已记录定位信息；不要手动改它，直接走 `fill.py`。
- **用户只给了一份已填好的样例**：先从样例中抽取文字，按字段 `label` 对齐；不要把样例里的固定说明当填写值。

## 输出给用户

简要报告：识别到多少字段、哪些字段留空/待补充、输出文件路径、是否通过验收。不要贴大段 JSON。
