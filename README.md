# 音乐课教案生成器

自动生成标准格式 Word 教案（德兴市二中音乐电子集体备课模板）。

---

## 项目结构

```
music_slides/
├── generate_docx.py          # 通用教案生成脚本（唯一需要运行的脚本）
├── lessons/                  # 每节课的数据文件（JSON）
│   ├── 第二单元/
│   ├── 第三单元/
│   └── ...
├── 八年级下/                 # 生成的 .docx 教案输出目录
│   ├── 第二单元/
│   └── ...
├── examples 八上/            # 参考示例（原始教案 docx）
├── 八下 培训用书.pdf          # OCR 用的教材扫描件
├── ocr_pdf.py                # OCR 辅助脚本（生成 Swift 二进制）
└── generate_lesson_plans.py  # 旧脚本（已废弃，内容已迁移到 JSON）
```

---

## 快速开始：制作一本新书的教案

### 第一步：建立目录

```bash
mkdir -p "lessons/九年级上/第一单元"
mkdir -p "九年级上/第一单元"
```

### 第二步：为每节课创建 JSON 数据文件

在 `lessons/九年级上/第一单元/` 下新建 JSON 文件，文件名即课名（如 `茉莉花.json`）。

**JSON 格式模板：**

```json
{
  "title": "《茉莉花》",
  "output_path": "/Users/lin/Documents/Code/music_slides/九年级上/第一单元/《茉莉花》.docx",
  "knowledge_goal": "了解……",
  "ability_goal": "能够……",
  "emotion_goal": "感受……",
  "focus": "……",
  "difficulty": "……",
  "intro": "（一）导入\n1. 情境创设：\n    ……",
  "new_lesson": "（二）……\n1. ……",
  "consolidate": "3. ……",
  "summary": "完整演唱……\n拓展……",
  "board": "板书设计  《茉莉花》\n……"
}
```

**字段说明：**

| 字段 | 对应教案位置 | 必填 |
|------|-------------|------|
| `title` | 课题行 | ✅ |
| `output_path` | 输出文件路径（不填则在 JSON 同目录生成） | 建议填 |
| `knowledge_goal` | 素养目标 → 知识目标 | ✅ |
| `ability_goal` | 素养目标 → 能力目标 | ✅ |
| `emotion_goal` | 素养目标 → 情感目标 | ✅ |
| `focus` | 教学重点 | ✅ |
| `difficulty` | 教学难点 | ✅ |
| `intro` | 教学过程 → 导入（第一格） | ✅ |
| `new_lesson` | 教学过程 → 创设情境/引入课题 | ✅ |
| `consolidate` | 教学过程 → 探究新知/讲授新课 | ✅ |
| `summary` | 教学过程 → 巩固应用 | ✅ |
| `board` | 板书设计（最后一行，跨全列） | ✅ |

### 第三步：生成 docx

```bash
# 生成单个课的教案
python3 generate_docx.py lessons/九年级上/第一单元/茉莉花.json

# 批量生成整个单元
python3 generate_docx.py lessons/九年级上/第一单元/

# 批量生成整本书
python3 generate_docx.py lessons/九年级上/
```

---

## 格式规范

- **字体**：正文宋体 + Times New Roman（中英文混排自动适配）
- **标题**：黑体加粗 14pt
- **正文**：11pt
- **页边距**：上下 2cm，左右 2.5cm
- **表格**：9 列 14 行，与德兴市二中集体备课模板完全一致

---

## 读取教材内容（OCR 扫描 PDF）

如果教材是扫描版 PDF（非文字版），使用 macOS 自带 Vision 框架 OCR：

### 编译 OCR 工具（仅需一次）

```bash
python3 ocr_pdf.py
# 编译完成后生成 /tmp/ocr_pdf 二进制文件
```

### 运行 OCR

```bash
# OCR 第 1-30 页
/tmp/ocr_pdf "/path/to/教材.pdf" 1 30 > /tmp/ocr_output.txt

# 查看结果
cat /tmp/ocr_output.txt
```

输出格式为：
```
=== 页面 1 ===
（识别出的文字内容）

=== 页面 2 ===
……
```

> **注意**：`/tmp/ocr_pdf` 重启电脑后会消失，需重新执行 `python3 ocr_pdf.py` 编译。

---

## 已完成的教案

### 八年级下

| 单元 | 课程 | 文件 |
|------|------|------|
| 第二单元 | 《但愿人长久》《沁园春·雪》《蜀道难》《秋之歌》 | 八年级下/第二单元/ |
| 第三单元 | 《念故乡》《第九（自新大陆）交响曲》《沃尔塔瓦河》《大海与辛巴达的船》 | 八年级下/第三单元/ |
| 第四单元 | 《红河谷》《拉库卡拉查》《凯皮拉的小火车》《化装舞会》《飞驰的鹰》《故乡的亲人》 | 八年级下/第四单元/ |
| 第五单元 | 《春江花月夜》（演唱）《春江花月夜》（器乐）《雨打芭蕉》《小放驴》《欢乐歌》《老鼠娶亲》 | 八年级下/第五单元/ |
| 第六单元 | 《军民大生产》《延安颂》《在太行山上》《我为祖国献石油》《愿亲人早日养好伤》《公仆赞》《生死不离》 | 八年级下/第六单元/ |

---

## 修改已有教案

直接编辑对应的 JSON 文件，然后重新运行生成命令：

```bash
python3 generate_docx.py lessons/第六单元/生死不离.json
```

---

## 环境要求

- Python 3.9+（推荐 `/Applications/Xcode.app/Contents/Developer/usr/bin/python3`）
- python-docx：`pip3 install python-docx`
- OCR 功能需要 macOS 11+（Big Sur 及以上，内置 Vision 框架）
