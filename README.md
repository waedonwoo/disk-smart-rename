# disk-smart-rename

> 一键让你的磁盘目录从"乱码地狱"变得井井有条

## 简介

Smart Rename 是一款面向个人与企业用户的**磁盘文件智能重命名工具**，同时也是可直接导入 SkillHub 的 AI Skill。

自动识别并批量处理那些让人头疼的乱码文件：
- 纯数字/hash 文件名 → 智能提取文档标题
- 无意义图片名 → 自动读取 EXIF 拍摄时间戳
- 电子发票/增值税发票 → 统一命名为 `电子发票_YYYYMMDD`
- 截图类 → 时间戳自动归类

**特点**：
- 🛡️ **零破坏安全** — 仅重命名，不删除任何文件；同目录重名自动加后缀
- 🔄 **幂等设计** — 已命名良好的文件自动跳过，定期跑一次就能保持整洁
- 🪶 **开箱即用** — 首次运行自动安装依赖，跨平台 (Windows/macOS/Linux)
- 📋 **完整日志** — 每次执行留存 rename_plan.csv + rename_log.txt

## 快速开始

### 方式 1：导入 SkillHub

1. 登录 [SkillHub](https://skillhub.cn)
2. 「导入 Skill」→ 选「GitHub 仓库」
3. 填写仓库地址：`https://github.com/waedonwoo/disk-smart-rename`
4. 一键安装使用

### 方式 2：命令行直接使用

```bash
# 预览（不改文件）
python scripts/smart_rename.py -r D:\Work

# 执行重命名
python scripts/smart_rename.py -r D:\Work D:\Downloads --apply
```

### 依赖

首次运行自动安装：

```
exifread, Pillow, PyPDF2, python-docx, python-pptx, openpyxl
```

## 支持的文件类型

| 类别 | 格式 | 策略 |
|---|---|---|
| 文档 | PDF, DOCX, PPTX, XLSX, XLS | 提取 meta title / 正文标题 / 首行内容 |
| 图片 | JPG, PNG, CR2, CR3, NEF, ARW, DNG, HEIC, BMP, GIF, WEBP | EXIF DateTimeOriginal → Image DateTime → 文件 mtime |

## 安全保障

- ✅ 仅重命名，不删除
- ✅ 同目录重名自动加 `_1 _2 _N` 后缀，零覆盖
- ✅ 预览模式先看后改
- ✅ 完整执行日志留存
- ✅ 不联网、不上传任何文件到服务器
