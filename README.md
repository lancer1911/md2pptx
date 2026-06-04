# MD2PPTX

将结构化 Markdown 转换为 PowerPoint (`.pptx`) 的本地工具。  
本项目适合将法律分析、技术说明、项目方案、培训材料、会议汇报等内容快速整理为可编辑的 PPT。

转换器会读取一个 Markdown 文件，并基于指定的 PowerPoint 模板生成新的演示文稿。生成结果主要由普通 PowerPoint 文本框、表格、图片和 shape 组成，便于后续手工修改。

---

## 主要功能

当前版本支持以下内容：

- 封面页：使用一级标题 `#`
- 正文页：使用二级标题 `##`
- 小标题：支持 `###` 到 `######`
- 普通段落
- 无序列表、嵌套列表
- 编号列表
- 行内格式：
  - `**加粗**`
  - `*斜体*`
  - `***加粗斜体***`
  - `~~删除线~~`
  - `` `行内代码` ``
  - `[链接文本](https://example.com)`
- 引用块 `>`，并带左侧竖线
- Markdown 表格
- fenced code block，例如 ```python
- 本地图片引用，例如 `![说明](images/example.png)`
- Speaker notes
- 仿 SmartArt 图形
- 基于模板的标题、正文区域、页脚和装饰元素复用
- 模板最后一页自动作为结束页追加，例如 `Thank you!` 页面

---

## SmartArt 支持

本项目支持一种“仿 SmartArt”语法。  
它并不会生成 Office 原生 SmartArt 对象，而是根据 Markdown 自动生成可编辑的 PowerPoint 普通 shape、连接线和文本框。

优点是：

- 生成后可在 PowerPoint 中直接编辑文字、颜色、大小和位置
- 不依赖 Office 原生 SmartArt XML
- 适合用 Markdown 描述流程图、组织架构、矩阵、时间线等常见商务图形
- 连接线、节点框、文字均为可编辑对象

### SmartArt 基本语法

```markdown
:::smartart type="process" style="chevron" theme="blue"
- 证据收集
- 侵权比对
- 提起诉讼
- 庭审听证
- 裁决结案
:::
```

### 支持的 SmartArt 类型

| type | style | 适用场景 |
| --- | --- | --- |
| `process` | `chevron` | 阶段、步骤、流程、项目推进 |
| `process` | `numbered` | 编号步骤、执行清单、操作流程 |
| `cycle` | - | 闭环、循环管理、持续监控 |
| `hierarchy` | - | 组织架构、团队分工、树形关系 |
| `matrix` | - | 风险矩阵、二维分类、优先级矩阵 |
| `pyramid` | - | 层级递进、优先级、价值分层 |
| `timeline` | - | 时间安排、项目节点、诉讼进度 |

### 颜色主题

目前常用主题包括：

- `blue`
- `cyan`

示例：

```markdown
:::smartart type="cycle" theme="cyan"
- 市场监控
- 侵权分析
- 启动程序
- 跟踪进展
- 复盘优化
:::
```

### hierarchy 层级图

`hierarchy` 使用缩进表示上下级关系。建议每一级缩进使用两个空格，不要混用 Tab。

```markdown
:::smartart type="hierarchy" theme="blue"
- 专利维权项目组
  - 诉讼组
    - 诉状起草
    - 庭审准备
  - 无效组
    - 现有技术检索
    - 无效请求撰写
  - 赔偿组
    - 销售数据取证
    - 利润损失核算
:::
```

### matrix 矩阵图

`matrix` 每一项可使用 `/` 分隔多行文字。

```markdown
:::smartart type="matrix" theme="cyan"
- 高影响 / 低概率 / 重点防范
- 高影响 / 高概率 / 立即应对
- 低影响 / 低概率 / 定期观察
- 低影响 / 高概率 / 接受管理
:::
```

### timeline 时间线

`timeline` 建议使用 `时间 | 说明` 的形式。

```markdown
:::smartart type="timeline" theme="blue"
- 第1周 | 完成诉前评估与样品公证
- 第3周 | 递交起诉状
- 第2月 | 提交无效宣告请求
- 第4月 | 一审庭审
- 第6月+ | 等待裁决或启动和解
:::
```

### SmartArt 排版说明

当前版本已针对 SmartArt 做了以下优化：

- SmartArt 会按照 Markdown 出现顺序排版
- 如果 SmartArt 前面有正文、表格、代码块或图片，图形会自动下移
- 节点文字支持左右、上下居中
- `hierarchy` 中连接线会绘制在节点框下方，不会压住文字和节点框
- 长段落自动换行后，后续表格和 SmartArt 会根据估算高度下移，减少遮挡

建议每页只放一个 SmartArt。  
如果图形节点较多，建议拆分为多页。

---

## macOS 使用方法

### 1. 安装依赖

在终端进入项目目录：

```bash
cd /path/to/MD2PPTX
```

首次使用时运行：

```bash
bash setup.sh
```

该脚本会：

- 检查 Python 版本，要求 Python 3.10+
- 创建本地虚拟环境 `env`
- 安装 `requirements.txt` 中的依赖

也可以手动安装：

```bash
python3 -m venv env
source env/bin/activate
pip install -r requirements.txt
```

### 2. 转换 Markdown 为 PPTX

使用默认模板 `template.pptx`：

```bash
bash run.sh sample.md
```

指定输出文件：

```bash
bash run.sh sample.md output.pptx
```

指定自定义模板：

```bash
bash run.sh sample.md output.pptx my_template.pptx
```

当使用自定义模板时，程序会自动调用 `inspect_template.py` 生成对应的 `.json` 布局配置；如果配置文件已经存在，会直接复用。

---

## Windows 使用方法

### 1. 安装 Python

建议安装 Python 3.10 或更高版本。安装时请勾选：

```text
Add python.exe to PATH
```

如果系统安装了 Python Launcher，`setup.bat` 会优先使用：

```bat
py -3
```

否则会使用：

```bat
python
```

### 2. 安装依赖

在资源管理器中进入项目目录，双击：

```text
setup.bat
```

或者在 Windows Terminal / 命令提示符中运行：

```bat
setup.bat
```

该脚本会：

- 检查 Python 版本
- 创建本地虚拟环境 `env`
- 安装 `requirements.txt` 中的依赖

### 3. 转换 Markdown 为 PPTX

使用默认模板 `template.pptx`：

```bat
run.bat sample.md
```

指定输出文件：

```bat
run.bat sample.md output.pptx
```

指定自定义模板：

```bat
run.bat sample.md output.pptx my_template.pptx
```

也可以直接双击：

```text
convert_sample_windows.cmd
```

它会在必要时先执行 `setup.bat`，然后将 `sample.md` 转换为示例 PPTX。

### 4. PowerShell 中的用法

在 PowerShell 中建议写成：

```powershell
.\setup.bat
.\run.bat .\sample.md .\output.pptx
```

---

## 命令参数说明

`run.sh` 与 `run.bat` 的参数保持一致：

```text
run input.md [output.pptx] [template.pptx]
```

| 参数 | 必填 | 说明 |
| --- | --- | --- |
| `input.md` | 是 | 输入 Markdown 文件 |
| `output.pptx` | 否 | 输出 PPTX 文件；不填时默认输出到输入文件同目录 |
| `template.pptx` | 否 | PPT 模板；不填时优先使用脚本同目录下的 `template.pptx` |

macOS：

```bash
bash run.sh input.md output.pptx template.pptx
```

Windows：

```bat
run.bat input.md output.pptx template.pptx
```

---

## 跨平台文件说明

| 文件 | macOS / Linux | Windows | 说明 |
| --- | --- | --- | --- |
| `setup.sh` | ✅ | - | 创建虚拟环境并安装依赖 |
| `run.sh` | ✅ | - | 执行 Markdown 到 PPTX 的转换 |
| `setup.bat` | - | ✅ | Windows 下创建虚拟环境并安装依赖 |
| `run.bat` | - | ✅ | Windows 下执行转换 |
| `convert_sample_windows.cmd` | - | ✅ | Windows 双击测试示例转换 |
| `requirements.txt` | ✅ | ✅ | Python 依赖清单 |



### 自定义模板 JSON 识别说明

`inspect_template.py` 会按照项目模板约定识别：

- 第 1 页：封面样式参考；
- 第 2 页：正文页样式参考；
- 最后一页：感谢页 / 结束页，不参与正文区域检测。

如果自动生成的 `.json` 中 `content_l/content_t/content_w/content_h` 明显偏移，正文就会整体偏右或偏下。  
此时应以模板第 2 页的正文框为准手动修正 JSON，或重新运行新版 `inspect_template.py`。

## 模板约定：自动追加最后一页

当前版本会把 **模板文件的最后一页** 作为固定结束页，并在每次生成 PPTX 时自动追加到演示文稿末尾。正文页仍然使用原来的空白页生成逻辑，不会继承模板中的标题/正文占位符。

推荐模板结构：

| 模板页 | 用途 |
| --- | --- |
| 第 1 页 | 封面页样式参考 |
| 第 2 页 | 正文页样式参考 |
| 最后一页 | 结束页 / 感谢页，例如 `Thank you!` |

转换时，程序会：

1. 读取模板的版式、标题区域、正文区域、页脚和装饰元素；
2. 使用空白演示文稿生成 Markdown 对应的正文页，避免继承模板占位符；
3. 单独复制模板最后一页的图形和文字；
4. 将复制得到的结束页追加到最终 PPTX 的最后。

因此，Markdown 文件中不需要再手动写“Thank you”页。  
如果希望更换结束页样式，只需要修改 `template.pptx` 的最后一页即可。


## Markdown 写作规则

### 封面页

一级标题 `#` 会生成封面页。

```markdown
# 专利维权项目工作方案
Lancer1911 | 2026
```

### 正文页

二级标题 `##` 会生成普通内容页。

```markdown
## 项目整体流程

从证据收集到最终裁决，五个阶段环环相扣。
```

### 小标题

支持 `###` 到 `######`。

```markdown
### 一、核心结论
#### 1. 证据基础
##### 补充说明
```

### 表格

```markdown
| 阶段 | 主要工作 | 输出 |
| --- | --- | --- |
| 诉前评估 | 初步比对 | 分析备忘录 |
| 证据固定 | 公证购买 | 公证书 |
```

### 引用块

```markdown
> 关键结论：证据固定通常比直接提出高额赔偿请求更重要。
> 连续多行引用会合并为同一个引用块。
```

### 代码块

````markdown
```python
from pathlib import Path
print(Path("sample.md").exists())
```
````

### 图片

```markdown
![示例图片](images/example.png)
```

建议使用相对路径，并将图片放在 Markdown 文件同级目录或子目录中。

---

## Speaker Notes

可以在页面底部使用如下格式加入演讲备注：

```markdown
> **Speaker note:** 这里写给演讲者看的备注。
```

备注不会作为普通正文展示在页面主体中，而是写入 PowerPoint 的 notes 区域。

---

## 示例文件

项目中包含示例文件：

- `sample.md`：覆盖常见 Markdown 标记和全部 SmartArt 类型
- `sample_en.md`：英文示例
- `template.pptx`：示例模板
- `layout_example.json`：布局配置示例

建议先运行：

```bash
python md_to_pptx.py sample.md template.pptx sample_output.pptx
```

确认环境和模板均正常后，再处理正式材料。

---

## 推荐工作流

1. 准备或选择一个 `template.pptx`
2. 根据 `INSTRUCTION_md_format.md` 将材料整理为 Markdown
3. 使用 `sample.md` 检查语法参考
4. 运行 `md_to_pptx.py`
5. 打开生成的 PPTX 进行人工微调
6. 如需从现有 PPT/PDF 反向整理为 Markdown，可参考 `INSTRUCTION-convert_existing_presentation_to_MD.md`

---

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `md_to_pptx.py` | 主程序 |
| `README.md` | 项目说明 |
| `INSTRUCTION_md_format.md` | 将材料整理为 MD2PPTX Markdown 的写作规范 |
| `INSTRUCTION-convert_existing_presentation_to_MD.md` | 将现有 PPT/PDF 反向整理为 Markdown 的规范 |
| `sample.md` | 全功能中文示例 |
| `sample_en.md` | 英文示例 |
| `template.pptx` | 示例模板；最后一页会自动作为结束页追加 |
| `layout_example.json` | 布局配置示例 |
| `requirements.txt` | Python 依赖 |
| `run.sh` / `run.bat` | 运行脚本 |
| `setup.sh` / `setup.bat` | 初始化脚本 |

---

## 注意事项

- 本项目重点是生成结构清晰、可编辑、可继续加工的 PPTX，而不是逐像素复刻复杂设计稿。
- SmartArt 是由普通 shape 组合而成，不是 Office 原生 SmartArt。
- 过长段落、过大表格、过多节点的 SmartArt 仍建议拆页。
- 不建议使用 HTML 标签控制格式，例如 `<br>`、`<span>`、`<div>`。
- 不建议依赖脚注、Mermaid、LaTeX 公式等复杂 Markdown 扩展。
- 如页面出现拥挤，优先拆成多页，而不是强行缩小字号。

---

## 常见问题

### 1. 生成的 SmartArt 可以编辑吗？

可以。SmartArt 是普通 PowerPoint shape、连接线和文本框组成的，可以直接在 PowerPoint 中移动、改字、改颜色、改大小。

### 2. 这是 Office 原生 SmartArt 吗？

不是。它是“仿 SmartArt”的可编辑 shape 组合，兼容性和可控性更好。

### 3. 为什么复杂页面仍可能溢出？

程序会尽量估算正文、表格、代码块、图片和 SmartArt 的高度，但 PPT 排版与字体渲染受模板、系统字体和 PowerPoint 版本影响。正式材料建议每页控制信息量。

### 4. 表格太大怎么办？

建议拆成多页，并在每一页重复表头。

### 5. hierarchy 图线条为什么不会压住框？

当前版本采用两步绘制：先计算节点和连接线位置，先画连接线，再画节点框和文字，因此节点框会显示在线条之上。

---


### 6. Windows 下提示找不到 Python 怎么办？

请确认已安装 Python 3.10+，并且安装时勾选了 `Add python.exe to PATH`。  
也可以安装 Python Launcher，然后重新运行：

```bat
setup.bat
```

脚本会优先尝试 `py -3`，失败后再尝试 `python`。

### 7. Windows 下路径中有空格可以吗？

可以。`run.bat` 已对输入文件、输出文件和模板路径加引号处理。  
例如：

```bat
run.bat "C:\Users\Me\Documents\sample file.md" "C:\Users\Me\Desktop\output file.pptx"
```


### 8. 如何修改自动追加的感谢页？

直接编辑 `template.pptx` 的最后一页即可。程序每次生成时都会保留并追加模板最后一页。  
Markdown 中不需要单独写感谢页。

### 9. 如果不想自动追加感谢页怎么办？

当前版本默认追加模板最后一页。  
如果某个项目不需要结束页，可以使用一个最后一页为空白或极简内容的模板；也可以在代码中关闭该逻辑。

## License

MIT
