---
name: disk-smart-rename
displayName: Smart Rename 智能文件重命名助手
version: 1.0.0
description: 一键扫描磁盘并智能重命名乱码文件——PDF/DOCX/PPTX/XLSX 自动提取内容标题，图片读取 EXIF 拍摄时间戳，电子发票按日期归档。零破坏，仅重命名不删除。
license: MIT
tags: [文件整理, 磁盘清理, EXIF, 发票归档, 批量重命名, 效率工具]
---

# Smart Rename — 磁盘文件智能重命名工具

## 功能

对指定根目录递归扫描，智能识别"乱命名文件"并自动重命名：

| 文件类型 | 识别标准 | 命名策略 |
|---|---|---|
| **PDF** | 纯数字 / hash / 过短英文名 | 提取 meta title → 无则读第一页正文 → 再无则跳过 |
| **DOCX** | 同上 | 先读 Title 属性 → 再扫 Heading/Bold 段落 → 再读首段 |
| **PPTX** | 同上 | 前三页第一个有意义文本框 |
| **XLSX/XLS** | 同上 | Sheet 名 → 首行有效单元格 |
| **图片 (JPG/PNG/CR2/HEIC...)** | 纯数字 / hash / 设备前缀(IMG/DSC) / 截图类(screenshot/mmexport/微信) | EXIF DateTimeOriginal → 退到 Image DateTime → 退到文件 mtime |
| **发票类** | 任何格式文档 | 内容含"电子发票/增值税发票" → `电子发票_YYYYMMDD` |

### 命名示例

| 原文件名 | 新文件名 | 依据 |
|---|---|---|
| `240103.pptx` | `2023年度总结.pptx` | PPT 首页标题 |
| `96949f13b854575a4b93df0b946d2e0b.docx` | `"深圳工业奖"评审实施方案.docx` | Word 标题属性 |
| `34626ca3958bc196d31e215d4477ddde.pdf` | `电子发票_20240315.pdf` | 发票内容 + 文件日期 |
| `39ac942f469c196e6c767d39eab7bcac.jpg` | `20250918_181747.jpg` | EXIF 拍摄时间 |
| `screenshot-1711605482343.png` | `20240328_135810.png` | 文件修改时间 |
| `IMG_2026.jpg` | `20261118_141907.jpg` | 设备前缀 → mtime |

### 安全原则（零破坏）

1. **仅重命名，不删除** — 所有文件原地 os.rename
2. **零覆盖** — 同目录重名自动加 `_1 _2 ... _N` 后缀
3. **预览优先** — 默认 preview 模式不改；`--apply` 才执行
4. **幂等** — 多次跑只处理新出现的乱码文件，已命名的自动跳过
5. **完整日志** — 每次执行追加写入 rename_log.txt

## 何时使用

触发词：重命名 / 整理文件名 / 批量改名 / 智能命名 / 文件归档 / EXIF 命名 / 发票识别

## 何时不要用

- 目录重组 / 移动文件 → 用 PowerShell Move-Item
- 删除重复文件 → 用 PowerShell Get-FileHash 去重
- 图片视觉识别 → 需要 AI 模型

## 使用方法

### 1. 单根目录

```bash
python scripts/smart_rename.py -r D:\Work --apply
```

### 2. 多根目录

```bash
python scripts/smart_rename.py -r D:\Work D:\Downloads D:\Desktop --apply
```

### 3. 先预览不改

```bash
python scripts/smart_rename.py -r D:\Work
```

### 依赖

首次运行自动安装：`exifread Pillow PyPDF2 python-docx python-pptx openpyxl`

## 输出

| 文件 | 位置 | 内容 |
|---|---|---|
| rename_plan.csv | 脚本同目录 | 每次完整扫描结果（含 status: RENAME/SKIP/FAILED） |
| rename_log.txt | 脚本同目录 | 每次执行追加日志（OK/FAIL/SKIP/COLLISIONS 计数） |

## 内置跳过目录（不会被扫描）

`.git` / `node_modules` / `.wechat` / `Tencent Files` / `__pycache__`

## 已知限制

- 非标准格式的 PDF/DOCX（损坏或加密）无法提取内容 → 安全跳过
- PNG/GIF 等无 EXIF 图片 → 退到文件 mtime（可能不准）
- 文件名含 Windows 非法字符（如控制字节）的 PDF → 安全跳过，无法自动修复

## 技术实现要点

1. **跨平台路径兼容** — 用 `os.path.expanduser` + 绝对路径解析，支持 Windows/macOS/Linux
2. **PDF 脏数据容错** — 跳过 PyPDF2 的 XRef/Warning 异常，单页失败不中断整体流程
3. **PPT 模板名污染** — 黑名单过滤 "PowerPoint Presentation" / "Title Slide Layout" / "Sheet0-3" 等无价值 meta title
4. **图片批量冲突** — 同一天拍摄的几百张照片，秒级时间戳可能重复，自动 `_N` 后缀避免覆盖
5. **依赖自动安装** — 用 `ensure_libs()` 在首次运行时 pip install，不要求预装环境

## 目录结构

```
disk-smart-rename/
├── SKILL.md                   # SkillHub 入口
├── README.md                  # 说明文档
└── scripts/
    └── smart_rename.py        # 重命名核心脚本
```
