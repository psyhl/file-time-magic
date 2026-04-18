---
name: file-time-magic
description: |
  文件时间属性修改技能。自动设置文件的创建时间、修改时间、访问时间和编辑时长，确保所有时间逻辑一致且真实可信。
  
  触发场景：
  - 用户要求修改文件时间属性
  - 用户说"这个文件编辑了X小时"
  - 用户说"把时间改成..."或"文件时间要合理"
  - 用户指定创建时间或修改时间
  - 用户提问"能不能改文件编辑时长"
  
metadata:
  openclaw:
    emoji: "⏰"
---

# file-time-magic — 文件时间属性修改技能 v1.1.1

## 核心功能

| 能力 | 说明 |
|------|------|
| **编辑时长设置** | 自动处理编辑时长（带随机误差），并同步设置文件系统时间 |
| **用户指定时间点** | 支持用户指定创建时间、修改时间，系统自动计算编辑时长 |
| **时间逻辑自动计算** | 创建时间、修改时间、访问时间自动保持逻辑一致 |
| **随机性注入** | 分钟有±5误差，秒数随机，看起来更真实 |
| **Office 内部元数据** | docx/pptx/xlsx 同时更新 app.xml（TotalTime）+ core.xml（创建/修改时间） |
| **文件夹支持** | 支持修改文件夹的创建/修改/访问时间 |
| **Windows API** | 创建时间使用原生 SetFileTime，无需 PowerShell |

## ⚠️ 核心原则

| 原则 | 说明 |
|------|------|
| **时间逻辑一致性** | 创建时间 + 编辑时长 ≤ 修改时间 ≤ 访问时间 |
| **随机性必须** | 用户说"2小时"，实际是115-125分钟；秒数随机 |
| **时间上限** | 所有时间不超过当前时间（除非 --force） |
| **真实可信** | 工作时间在合理范围（08:00-22:00） |

## 命令语法

```python
# 基本用法（只指定编辑时长）
python "{SKILL_DIR}/scripts/set_file_time.py" --file "<文件路径>" --edit-duration "2小时"

# 指定创建时间
python "{SKILL_DIR}/scripts/set_file_time.py" --file "<文件路径>" --create-time "2024-01-15 09:30:00" --edit-duration "3小时"

# 指定修改时间
python "{SKILL_DIR}/scripts/set_file_time.py" --file "<文件路径>" --modify-time "2026-04-18 14:00:00" --edit-duration "90分钟"

# 同时指定创建时间和修改时间（自动计算编辑时长）
python "{SKILL_DIR}/scripts/set_file_time.py" --file "<文件路径>" --create-time "2024-01-15 09:00:00" --modify-time "2026-04-18 14:00:00"

# 预览（不实际修改）
python "{SKILL_DIR}/scripts/set_file_time.py" --file "<文件路径>" --edit-duration "2小时" --dry-run
```

**参数说明**

| 参数 | 必填 | 说明 |
|------|------|------|
| `--file` | 是 | 目标文件路径 |
| `--edit-duration` | 否* | 编辑时长，如 "2小时"、"120分钟"、"3h" |
| `--create-time` | 否* | 创建时间，如 "2024-01-15 09:30:00" |
| `--modify-time` | 否* | 修改时间，如 "2026-04-18 14:00:00" |
| `--access-time` | 否 | 访问时间（可选） |
| `--base-time` | 否 | 参考时间（默认当前时间） |
| `--dry-run` | 否 | 只显示计算结果，不实际修改 |
| `--force` | 否 | 强制执行，跳过未来时间确认 |

\* 至少需要指定 `--edit-duration` 或 `--create-time`/`--modify-time` 其中之一

## 使用模式

| 用户输入 | 系统行为 |
|---------|---------|
| 只指定 `--edit-duration` | 从当前时间往前推算所有时间 |
| `--create-time` + `--edit-duration` | 从创建时间往后推算修改时间 |
| `--modify-time` + `--edit-duration` | 从修改时间往前推算创建时间（调整到工作时间） |
| `--create-time` + `--modify-time` | 自动计算编辑时长（时间差） |
| 三时间全指定 | 直接设置所有时间 |

## 时间计算规则

1. **编辑时长随机化**
   - "X小时"：实际 = X×60 ± random(-5, +5) 分钟
   - "X分钟"：实际 = X ± random(-3, +3) 分钟
   - 秒数：random(0, 59)

2. **创建时间**
   - 用户指定时：直接使用（加随机秒）
   - 自动计算时：调整到工作时间（08:00-22:00）

3. **修改时间**
   - 用户指定时：直接使用（加随机秒）
   - 自动计算时：= 创建时间 + 编辑时长 + random(0, 2) 分钟

4. **访问时间**
   - 用户指定时：直接使用
   - 自动计算时：= 修改时间 + random(3, 15) 分钟

## 时长格式支持

| 格式 | 示例 | 结果 |
|------|------|------|
| 中文 | "2小时" | 120 分钟 |
| 中文混合 | "2小时30分" | 150 分钟 |
| 口语化 | "两小时半" | 150 分钟 |
| 小数 | "1.5小时" | 90 分钟 |
| 纯数字 | "90" | 90 分钟 |
| 英文简写 | "3h"、"2h30m" | 180 / 150 分钟 |

## 时间格式支持

| 输入格式 | 示例 |
|---------|------|
| 标准格式 | "2024-01-15 09:30:00" |
| 无秒 | "2024-01-15 09:30" |
| 斜杠 | "2024/01/15 09:30" |
| 月-日 | "01-15 09:30"（默认今年） |
| 仅时间 | "09:30"（默认今天） |

## Office 文件内部结构

对于 `.docx` / `.pptx` / `.xlsx`，脚本同时修改两处：

| XML 文件 | 修改内容 | 影响位置 |
|---------|---------|---------|
| `docProps/app.xml` | `TotalTime`（编辑时长，单位：分钟） | 文件 → 属性 → 详细信息 → 编辑时长 |
| `docProps/core.xml` | `dcterms:created`、`dcterms:modified` | 文件 → 信息 → 属性 → 创建时间、修改时间 |

**为什么两个都要改？**  
如果只改 `app.xml`，Word 属性对话框里的"创建时间"和"修改时间"仍然是旧值，容易露馅。

## 文件夹支持

| 特性 | 文件 | 文件夹 |
|------|------|--------|
| 文件系统时间戳 | ✅ 创建/修改/访问 | ✅ 创建/修改/访问 |
| Office 内部编辑时长 | ✅ 有 (TotalTime) | ❌ 不适用 |
| 修改时间语义 | 最后编辑内容 | 最后添加/删除文件 |

**注意**：文件夹不支持 `--edit-duration`（无内部编辑时长）。

## 处理边界情况

| 情况 | 处理方式 |
|------|---------|
| 文件不存在 | 返回错误 |
| 文件被 Office 占用 | 自动关闭进程后再执行 |
| 编辑时长超过文件存在时间 | 自动调整到合理范围 |
| 创建时间在工作时间外 | 自动调整到最近工作时间（08:00-22:00） |
| 时间超过当前时间 | ⚠️ 警告提示（除非 --force） |
| 用户同时指定 create+modify+duration | 以 create+modify 为准，duration 忽略 |

## 验证修改结果

**PowerShell：**
```powershell
Get-Item "文件路径" | Select-Object Name, CreationTime, LastWriteTime
```

**Office 内部（用 Python 直接读 XML）：**
```python
import zipfile, xml.etree.ElementTree as ET
with zipfile.ZipFile("file.docx") as z:
    # 编辑时长
    app = ET.fromstring(z.read('docProps/app.xml'))
    ns = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
    tt = app.find('{%s}TotalTime' % ns)
    print('TotalTime:', tt.text, '分钟')
    # 创建/修改时间
    core = ET.fromstring(z.read('docProps/core.xml'))
    for el in core:
        if 'created' in el.tag or 'modified' in el.tag:
            print(el.tag.split('}')[1], ':', el.text)
```

## 注意事项

- 修改 Office 文件需要解压和重新打包（过程安全）
- 临时目录用 PID 区分，避免并发冲突
- 创建时间使用 Windows API SetFileTime（无需 PowerShell）
- `--dry-run` 模式适合先验证时间计算结果再执行

## v1.1 更新内容

- 🐛 修复：`tmp_dir` 缩进错误导致脚本崩溃
- 🐛 修复：core.xml（创建/修改时间）未被更新
- 🐛 修复：`parse_duration("2小时")` 返回 115 而非 120
- 🐛 修复：PowerShell 子进程调用在部分环境失败
- 🐛 修复：Windows API SetFileTime UTC 时区转换
- ✨ 新增：用户同时指定 create+modify 时，编辑时长自动从差值计算
- ✨ 新增：Windows 原生 API 设置文件时间（不依赖 PowerShell）
- 📄 更新：输出 JSON 增加 `office_internal` 字段
