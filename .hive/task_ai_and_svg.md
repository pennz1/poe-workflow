两个独立改动，请按顺序完成。

---

## 任务 1：AI 辅助元素定位 — pricing_automation.py

**目标**：在 Playwright 浏览器自动化中引入 AI，替代硬编码选择器策略。保留高层结构化流程，仅在"需要理解页面内容"的关键决策点调用 AI。

**改动文件**：pricing_automation.py

### 1.1 新增 PageSnapshot 数据类

在文件前部数据模型区（约第 62 行后）添加：

```python
@dataclass
class PageSnapshot:
    """页面当前状态的简化快照，供 AI 分析。"""
    search_results: List[str] = field(default_factory=list)
    visible_buttons: List[str] = field(default_factory=list)
    dropdown_options: Dict[str, List[str]] = field(default_factory=dict)
    raw_text: str = ""
```

### 1.2 新增 _capture_page_snapshot 方法

在 PricingCalculatorAutomation 类中添加。功能：提取搜索结果项文本（匹配 class 含 product/result/card 的元素）、可见按钮文本、所有 select 下拉选项、body 文本前 3000 字符。返回 PageSnapshot。

### 1.3 新增 _ai_locate 方法

接收 PageSnapshot + goal + context，调用 call_azure_openai 分析页面快照返回操作建议 JSON。prompt 向 AI 描述页面状态（搜索结果、可见按钮、下拉选项），要求返回 JSON：{"action": "click", "target_text": "完整按钮文本"} 或 {"action": "select", "target_option": "应选的选项文本"} 或 {"action": "none", "reason": "原因"}。失败返回 None。

### 1.4 改造 _search_and_add — AI 辅助选择

修改现有 _search_and_add 方法（第 811 行起）。核心改动：
- 当 count == 0 时，调用 AI 获取替代搜索词，重试搜索
- 当 count == 1 时，直接点击（不变）
- 当 count > 1 时，调用 AI 从多个结果中选择最匹配目标服务的那个
- 保留原有 force click 兜底逻辑

### 1.5 新增 configure_service_sku 方法

AI 辅助选择服务 SKU。先调用 _capture_page_snapshot，再调用 _ai_locate 找到最接近 desired_sku 的选项，然后通过 select_option 选中。失败返回 False，不阻塞流程。

### 1.6 更新 _add_resources_and_calibrate 调用处

在 _add_resources_and_calibrate 函数中（约第 1263 行），在 configure_service_region 之后调用 configure_service_sku：
```python
if resolved_sku:
    await automation.configure_service_sku(resolved_sku)
```

### 1.7 验证
```bash
.venv/bin/python -m py_compile pricing_automation.py
.venv/bin/python -m pytest test_pricing_automation.py -v 2>&1 | tail -10
```

---

## 任务 2：去掉售前方案文档中的 SVG 架构图

**目标**：移除解决方案文档中自动生成和插入 SVG 架构图的全部逻辑。

### 2.1 llm/prompts.py

删除第 206-268 行：SVG_SYSTEM_PROMPT 常量和 _extract_svg_from_response 函数（包括 "# SVG 架构图生成" 注释）。

### 2.2 documents/solution.py

1. 删除 generate_svg_architecture 整个函数（第 23-70 行）
2. 修改 create_solution_docx 函数：移除 svg_code 参数，移除 svg_code 分支逻辑（第 104-132 行），直接使用 markdown_to_docx(doc, body_content, body_size=9)
3. 清理无用 import：删除 add_svg_image_to_doc、SVG_SYSTEM_PROMPT、_extract_svg_from_response。注意保留 WD_ALIGN_PARAGRAPH 和 RGBColor（封面页还用）

### 2.3 pipeline.py

1. import 行删除 generate_svg_architecture
2. 删除 SVG 生成调用（第 76-77 行）
3. create_solution_docx 调用去掉 svg_code 参数
4. 删除 result["svg_code"] = svg_code 行

### 2.4 ui/tabs.py

搜索并删除所有 svg_code 相关引用。

### 2.5 验证
```bash
.venv/bin/python -m py_compile poeworkflow-refactor/llm/prompts.py
.venv/bin/python -m py_compile poeworkflow-refactor/documents/solution.py
.venv/bin/python -m py_compile poeworkflow-refactor/pipeline.py
.venv/bin/python -m py_compile poeworkflow-refactor/ui/tabs.py
grep -rn "svg" poeworkflow-refactor/ --include="*.py" | grep -iv "svg_code|svg_architecture|SVG_SYSTEM|_extract_svg|add_svg" || echo "SVG references cleaned"
```

---

## 实施顺序
先做任务 2（SVG 删除简单），再做任务 1（AI 辅助元素定位）。

## Commit
任务 2: `chore: remove SVG architecture diagram from solution document`
任务 1: `feat: AI-assisted element finding in pricing calculator automation`

全部完成后 team report 汇报两个任务的验证结果。
