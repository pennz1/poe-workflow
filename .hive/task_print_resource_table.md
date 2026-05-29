AI 解决方案文档生成后，立即在进度日志中打印第八章资源需求表。

## 改动文件
poeworkflow-refactor/pipeline.py

## 改动 1：新增资源章节提取函数

在 pipeline.py 顶部（generate_solution_artifact 之前，约第 45 行）添加：

```python
def _extract_resource_section(content: str) -> str:
    """从解决方案文档中提取第八章资源需求表。"""
    lines = content.split("\n")
    start = None
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("## 八、") or stripped.startswith("## 8.") or stripped == "## 八、资源架构":
            start = i
            break
    if start is None:
        return ""
    # 从标题行开始收集，直到遇到下一个 ## 标题或结束
    section_lines = []
    for i in range(start, len(lines)):
        line = lines[i]
        stripped = line.strip()
        if i > start and (stripped.startswith("## ") or stripped.startswith("# ")):
            break
        section_lines.append(line)
    return "\n".join(section_lines).strip()
```

## 改动 2：在解决方案生成后立即打印

在 `run_full_auto_poe` 函数中，约第 253 行 `st.session_state[target_key] = solution_artifact["content"]` 之后，添加：

```python
    # 打印第八章资源需求表，方便用户提前查看资源清单
    resource_section = _extract_resource_section(solution_artifact["content"])
    if resource_section:
        progress("---")
        progress("**📋 资源需求清单（第八章）**")
        progress("")
        for line in resource_section.split("\n"):
            progress(line)
        progress("")
        progress("**以上为方案文档中的资源需求，可据此在 Azure Pricing Calculator 手动制作价格表。**")
        progress("---")
```

## 验证
```
.venv/bin/python -m py_compile poeworkflow-refactor/pipeline.py
.venv/bin/python -c "
import sys; sys.path.insert(0,'poeworkflow-refactor')
from pipeline import _extract_resource_section
test_content = '''
## 七、集成架构
some content
## 八、资源架构
### Azure 资源需求
| 服务名称 | 配置规格 (SKU) | 区域 | 核心用途 |
| --- | --- | --- | --- |
| Azure OpenAI | GPT-5.4 | East US 2 | 核心推理 |
## 九、其他
'''
result = _extract_resource_section(test_content)
print(result)
assert 'Azure OpenAI' in result
assert '八、资源架构' in result
assert '九、其他' not in result
print('EXTRACTION OK')
"
.venv/bin/python -m pytest test_new_features.py -v 2>&1 | tail -5
```

## Commit
```
git add poeworkflow-refactor/pipeline.py
git commit -m "feat: print Chapter 8 resource table in logs after solution generation"
```

完成后 team report 汇报结果。
