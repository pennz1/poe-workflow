快速验证本次四个改动和残余文件清理。

## 改动清单
| Commit | 内容 |
|--------|------|
| 4385928 | Azure 登录去掉 window.open，自动复制+优雅手动 UI |
| c1bfe73 | 任务完成时间后增加耗时显示 |
| bd53a8e | Migrate 评估含 Storage/Disk + 存储冗余调优 |
| 4185132 | 清理残余文件 (1.svg, .browser_profile, .codex) |

## 验证步骤

### 1. 全语法编译
```
.venv/bin/python -m py_compile frontend/ui.py && echo "ui OK"
.venv/bin/python -m py_compile poeworkflow-refactor/ui/auto_poe.py && echo "auto_poe OK"
.venv/bin/python -m py_compile poeworkflow-refactor/azure/migrate.py && echo "migrate OK"
.venv/bin/python -m py_compile poeworkflow-refactor/pipeline.py && echo "pipeline OK"
```

### 2. 验证 window.open 已删除
```
grep -n "window.open" frontend/ui.py && echo "FAIL: window.open still present" || echo "PASS: window.open removed"
```

### 3. 验证耗时逻辑存在
```
grep -n "elapsed_str\|elapsed_sec\|auto_poe_start_time" poeworkflow-refactor/ui/auto_poe.py && echo "PASS" || echo "FAIL"
```

### 4. 验证 disk 评估配置
```
grep -n "PremiumSSD\|StandardHDD\|azureStorageRedundancy" poeworkflow-refactor/azure/migrate.py | head -5 && echo "PASS" || echo "FAIL"
```

### 5. 验证残余文件已删除
```
test -f 1.svg && echo "FAIL: 1.svg exists" || echo "PASS: 1.svg removed"
test -d .browser_profile && echo "FAIL: .browser_profile exists" || echo "PASS: .browser_profile removed"
test -d .codex && echo "FAIL: .codex exists" || echo "PASS: .codex removed"
```

### 6. Created on (UTC) 专项验证
运行 .venv/bin/python 脚本：导入 fix_assessment_excel_timestamps，构造最小 Excel（Assessment_Summary 含 Created on (UTC)，Assessment_Properties 含 Performance history start/end time），调用函数后验证 created_val == end_val，且日期在 POV 区间内，时分秒保留。

### 7. 回归测试
```
.venv/bin/python -m pytest test_tier_algorithm.py test_new_features.py -v 2>&1 | tail -10
```

### 8. Streamlit 启动检查
```
.venv/bin/streamlit run poeworkflow-refactor/main.py --server.port 8505 --server.headless true &
sleep 5 && curl -s -o /dev/null -w "%{http_code}" http://localhost:8505
kill %1 2>/dev/null; wait 2>/dev/null
```

每项标注 PASS/FAIL，全部通过后 team report 汇报。
