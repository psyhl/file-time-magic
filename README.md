# ⏰ file-time-magic

OpenClaw & SkillHub skill — 修改文件时间戳与 Office 编辑时长.

## ✨ 核心能力

| 能力 | 说明 |
|------|------|
| 🔧 编辑时长设置 | 修改 Office 文档（docx/pptx/xlsx）内部 TotalTime 属性 |
| 🕐 文件系统时间 | 同时设置创建时间、修改时间、访问时间 |
| 🎲 随机性注入 | 编辑时长 ±5 分钟误差，秒数随机，避免"太整齐" |
| 📁 文件夹支持 | 修改文件夹的创建时间、修改时间、访问时间 |
| ⚖️ 时间逻辑一致性 | 自动保证 创建时间 ≤ 修改时间 ≤ 访问时间 |
| 🛡️ 时间上限保护 | 所有时间不超过当前系统时间 |

## ⚠️ 平台要求

- **仅支持 Windows**（需要 PowerShell + Python 3）
- Mac/Linux 暂不支持文件夹创建时间修改

## 📦 安装

### 方式 1：通过 SkillHub 安装（推荐）

```
skillhub install file-time-magic
```

### 方式 2：手动安装

将本仓库克隆到 OpenClaw skills 目录：

```bash
git clone https://github.com/psyhl/file-time-magic.git ~/.qclaw/skills/file-time-magic
```

## 🚀 使用示例

### 只指定编辑时长

```bash
python set_file_time.py --file "报告.docx" --edit-duration "2小时"
```

### 指定创建时间 + 编辑时长

```bash
python set_file_time.py --file "报告.docx" --create-time "2024-03-15 09:00" --edit-duration "3小时"
```

### 指定创建和修改时间（自动计算编辑时长）

```bash
python set_file_time.py --file "报告.docx" --create-time "2024-01-15 09:00" --modify-time "2026-04-18 14:00"
```

### 修改文件夹时间

```bash
python set_file_time.py --file "C:\工作材料" --create-time "2024-01-15 09:00" --modify-time "2024-06-20 18:00"
```

### Dry-run 模式（只看不改）

```bash
python set_file_time.py --file "报告.docx" --edit-duration "1.5小时" --dry-run
```

## 🔑 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `--file` | ✅ | 目标文件或文件夹路径 |
| `--edit-duration` | ❌* | 编辑时长，如 "2小时"、"120分钟"、"3h" |
| `--create-time` | ❌* | 创建时间，如 "2024-01-15 09:30:00" |
| `--modify-time` | ❌* | 修改时间，如 "2026-04-18 14:00:00" |
| `--access-time` | ❌ | 访问时间（可选） |
| `--dry-run` | ❌ | 只显示计算结果，不实际修改 |
| `--force` | ❌ | 跳过未来时间确认 |

\* 至少需要指定 `--edit-duration` 或 `--create-time`/`--modify-time` 其中之一

## 🧠 时间计算逻辑

1. **编辑时长随机化**：用户说"2小时"→ 实际 115-125 分钟
2. **工作时间保护**：创建时间自动调整到 08:00-22:00
3. **访问时间推算**：修改时间 + 随机 3-15 分钟
4. **Office 内部同步**：docx/pptx/xlsx 的 TotalTime 属性同步更新

## 📄 License

MIT License

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！
