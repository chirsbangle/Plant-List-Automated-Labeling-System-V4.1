# Plant-List-Automated-Labeling-System-V4.1

A professional Word VBA utility for automatically annotating Chinese plant names with standardized Latin names from an Excel glossary. The system supports first-occurrence-only labeling, smart auditing of existing bracketed names, mismatch correction, highlight avoidance, token-based rendering, and scientific formatting for botanical nomenclature.

---

## Project Overview

Plant-List-Automated-Labeling-System-V4.1 is a Word VBA tool designed for the automatic annotation of Chinese plant names with standardized Latin names from an Excel glossary. It is intended for botanical checklists, floristic records, taxonomic notes, academic manuscripts, and other structured plant-related documents.

The system is built to improve consistency, reduce repetitive manual work, and support scientific writing workflows in Word documents. It can recognize existing annotations, avoid duplicate labeling, skip highlighted content, detect incorrect bracketed Latin names, and append correct forms according to the glossary.

===========================================================================

Main Features

Automatic annotation of Chinese plant names using an Excel glossary

First-occurrence-only processing to avoid repeated labeling

Smart audit of existing bracketed Latin names

Detection of mismatched Latin names already present in the document

Automatic correction workflow:

existing incorrect bracket is marked in red

correct Latin name is appended in a new bracket

Yellow-highlight avoidance

Token-based rendering to prevent document layout shifts

Scientific formatting for botanical Latin names

Support for family, genus, and species levels

Expected Output Behavior
1. If no bracketed Latin name exists

The tool adds a new standard annotation:

车前科(Plantaginaceae)
2. If the existing bracketed Latin name is already correct

The tool accepts it and skips that item.

Example:

车前科（Plantaginaceae）
3. If the existing bracketed Latin name is incorrect

The tool will:

mark the incorrect bracket in red

append the correct glossary value in a new bracket

Example:

Original text:

谷精草科（Plantaginaceae）

Correct value from Excel:

Eriocaulaceae

Result:

谷精草科（Plantaginaceae）（Eriocaulaceae）

The incorrect bracket remains visible for review, while the correct standardized form is appended after it.

Required Files

The VBA project normally includes the following files:

Module1.bas

frmProgress.frm

frmProgress.frx

These files should be imported into the Word VBA project.

Excel Glossary Format

The Excel glossary must follow this six-column structure:

A = Chinese family name

B = Latin family name

C = Chinese genus name

D = Latin genus name

E = Chinese species name

F = Latin species name

The first row should be the header row. Actual data should begin from row 2.

Example
A	B	C	D	E	F

车前科	Plantaginaceae	车前属	Plantago	车前	Plantago asiatica
谷精草科	Eriocaulaceae	谷精草属	Eriocaulon	谷精草	Eriocaulon buergerianum

Installation
Step 1. Open Microsoft Word

Open the Word document or template in which you want to install the macro.

Step 2. Open the VBA Editor

Press:

Alt + F11
Step 3. Import the VBA files

In the VBA project panel:

Right-click the target project

Choose Import File...

Import:

Module1.bas

frmProgress.frm

Make sure frmProgress.frx is in the same folder when importing the form.

Step 4. Confirm imported components

After import, confirm that the project contains:

Module1

frmProgress

Required Form Configuration

The progress form must be named exactly:

frmProgress

It must contain the following controls:

lblStatus

ProgressBarBack

ProgressBarFore

These names must match exactly.

Recommended form roles

lblStatus: displays the current processing step

ProgressBarBack: background bar

ProgressBarFore: foreground progress bar

Recommended form initialization code

Place this in the code page of frmProgress:

Option Explicit

Private Sub UserForm_Initialize()
    Me.lblStatus.Caption = "准备开始..."
    Me.ProgressBarFore.Width = 0
End Sub
How to Use
Step 1. Prepare the Excel glossary

Make sure the glossary file follows the required A-F structure.

Step 2. Open the Word document

Open the document that contains the Chinese plant names you want to process.

Step 3. Confirm the active document

The macro works on the current active Word document. Make sure the correct file is selected.

Step 4. Run the macro

Run the main entry procedure:

AddLatinNames_Turbo_V4_1_Final
Step 5. Select the glossary file

When prompted, choose the Excel glossary file.

Step 6. Wait for processing

The system will:

initialize the progress form

scan the document

audit existing bracketed Latin names

lock first valid matches using internal tokens

render the final formatted output

Step 7. Review the results

After completion:

correctly existing annotations remain unchanged

new standard annotations are inserted

incorrect existing brackets are marked in red

corrected glossary values are appended in new brackets

Detailed Processing Logic

The macro follows a multi-stage workflow.

Stage 1. Audit

The system scans the document for Chinese plant names and checks their surrounding context.

Stage 2. Match Decision

For each detected name, the system determines whether:

there is no bracket after the name

there is already a correct bracket

there is already an incorrect bracket

the content is inside yellow-highlighted text

the current occurrence should be skipped due to nested-name conflicts

Stage 3. Token Locking

Instead of directly inserting formatted text during the search phase, the system temporarily replaces matched names with internal tokens.

This reduces the risk of:

repeated replacement

layout shift

nested recognition errors

broken formatting in long documents

Stage 4. Rendering

The internal tokens are replaced with final visible output:

中文(Latin) for normal new annotations

中文（错误Latin）（正确Latin） style correction for mismatch cases

Formatting Rules
Chinese text

Chinese text is kept in standard Chinese font formatting.

Latin names

Latin names are displayed in Times New Roman.

Italic rules

For genus- and species-level names, the Latin text is italicized according to botanical conventions.

Non-italic exceptions

The following markers are automatically forced to non-italic formatting:

var.

subsp.

f.

×

ssp.

This helps preserve more appropriate scientific formatting in botanical names.

First-Occurrence Rule

The system processes only the first valid occurrence of each plant name in the document.

This rule is used to:

prevent duplicate annotation

reduce clutter in repeated text

keep academic writing cleaner

maintain consistency in plant lists and manuscripts

If a term appears many times, only the first valid occurrence is processed.

Highlight Avoidance Rule

Any plant name appearing in yellow-highlighted text is skipped.

This is useful when the document contains:

sections under review

regions intentionally excluded from processing

passages where manual handling is preferred

Bracket Audit Rule

The system does not simply skip all text that already has brackets.

Instead, it checks whether the content inside the bracket matches the Excel glossary.

It also tolerates spacing differences such as:

车前科(Plantaginaceae)
车前科 (Plantaginaceae)
车前科    (Plantaginaceae)
车前科（Plantaginaceae）

If the bracketed content matches the glossary, it is accepted.

If the bracketed content does not match the glossary:

the old bracket is preserved

its text is marked red

the correct glossary value is appended in a new bracket

Recommended Workflow

For safe and reliable use, the following workflow is recommended:

prepare or update the Excel glossary

create a backup copy of the original Word document

import the VBA module and form files into Word

run the macro

review any red brackets

confirm corrected cases manually if needed

save the processed document as a new version

This workflow is especially suitable for taxonomic writing, floristic records, and publication preparation.

Common Problems and Solutions
Problem: nothing happens when running the macro

Possible causes:

the wrong macro entry point was executed

the progress form is missing

the control names do not match the code

the wrong document is active

Check:

form name = frmProgress

controls = lblStatus, ProgressBarBack, ProgressBarFore

macro name = AddLatinNames_Turbo_V4_1_Final

Problem: the progress form appears but the bar does not move

Possible causes:

ProgressBarBack does not exist

ProgressBarFore is not correctly configured

control names do not match exactly

Problem: no names are annotated

Possible causes:

the Excel file is empty or incorrectly structured

the Word document does not contain matching Chinese names

all relevant text is highlighted yellow

existing brackets are already being accepted as correct

Problem: incorrect brackets are not corrected

Possible causes:

the bracketed content uses an unexpected structure

the document contains hidden characters or formatting interruptions

the glossary entry contains spelling inconsistencies

Intended Use Cases

This tool is suitable for:

botanical checklists

floristic inventories

taxonomic notes

academic manuscripts

plant resource catalogues

educational botanical reference documents

Chinese-to-Latin annotation workflows in Word

Version

Current public repository version: V4.1

License / Usage Note

This project is presented as a scientific research utility for academic and technical document processing. Please ensure that the glossary data used for annotation is accurate and curated before applying the tool to formal documents.

Author

Chris Bangle

Repository Structure
Plant-List-Automated-Labeling-System-V4.1/
├─ Module1.bas
├─ frmProgress.frm
├─ frmProgress.frx
└─ README.md


主要功能

本工具提供以下核心功能：

按 Excel 名录库自动匹配中文植物名称与标准拉丁名

每个植物名仅处理首次出现

自动识别文中已有括号拉丁名

若已有括号内容与 Excel 一致，则承认并跳过

若已有括号内容与 Excel 不一致，则：

将原括号内容标红

在其后追加正确拉丁名括号

自动跳过黄色高亮内容

使用令牌占位机制避免长文档中替换错位

自动执行植物学命名排版规范

输出效果示例
1. 文中原本没有括号

原文：

车前科

处理后：

车前科(Plantaginaceae)
2. 文中已有正确括号

原文：

车前科（Plantaginaceae）

处理后：

不做修改

系统承认为已正确标注

3. 文中已有错误括号

原文：

谷精草科（Plantaginaceae）

Excel 中正确值为：

Eriocaulaceae

处理后：

谷精草科（Plantaginaceae）（Eriocaulaceae）

同时：

原来的 （Plantaginaceae） 会变成红色

后追加的 （Eriocaulaceae） 为标准格式

文件组成

项目通常包含以下文件：

Module1.bas

frmProgress.frm

frmProgress.frx

README.md

其中：

Module1.bas 为主程序模块

frmProgress.frm 和 frmProgress.frx 为进度窗体

README.md 为项目说明文档

Excel 名录库格式要求

Excel 文件必须采用以下六列结构：

A 列：科中文名

B 列：科拉丁名

C 列：属中文名

D 列：属拉丁名

E 列：种中文名

F 列：种拉丁名

第一行应为表头，数据从第 2 行开始填写。

示例表格
A	B	C	D	E	F

车前科	Plantaginaceae	车前属	Plantago	车前	Plantago asiatica
谷精草科	Eriocaulaceae	谷精草属	Eriocaulon	谷精草	Eriocaulon buergerianum

安装方法
第一步：打开 Word

打开你需要安装宏的 Word 文档或模板。

第二步：打开 VBA 编辑器

按下：

Alt + F11

进入 VBA 编辑器。

第三步：导入 VBA 文件

在左侧工程窗口中：

右键目标工程

选择 Import File...（导入文件）

依次导入：

Module1.bas

frmProgress.frm

注意：导入窗体时，请确保 frmProgress.frx 与 frmProgress.frm 在同一文件夹中。

第四步：确认工程结构

导入完成后，工程中应能看到：

Module1

frmProgress

窗体要求

进度窗体名称必须严格为：

frmProgress

窗体中必须包含以下控件：

lblStatus

ProgressBarBack

ProgressBarFore

这些控件名称必须与代码完全一致，否则程序可能无法正常显示进度。

建议控件用途

lblStatus：显示当前处理状态

ProgressBarBack：进度条背景

ProgressBarFore：进度条前景

建议窗体初始化代码

将以下代码放入 frmProgress 的代码页中：

Option Explicit

Private Sub UserForm_Initialize()
    Me.lblStatus.Caption = "准备开始..."
    Me.ProgressBarFore.Width = 0
End Sub
使用流程（详细版）
第一步：准备 Excel 名录库

请先准备一个符合 A-F 六列格式要求的 Excel 文件。

第二步：打开待处理的 Word 文档

打开需要自动标注植物名的 Word 文档。

第三步：确认活动文档正确

本宏处理的是 当前活动 Word 文档，所以请确保当前打开的是正确文件。

第四步：运行主宏

在 VBA 编辑器中运行主入口过程，例如：

AddLatinNames_Turbo_V4_1_Final
第五步：选择 Excel 文件

程序启动后，会弹出文件选择框，请选择名录 Excel 文件。

第六步：等待系统处理

程序将自动执行以下步骤：

初始化进度窗体

扫描文档中的植物名

检查已有括号内容

判断是否需要新增、跳过或修正

使用令牌占位避免错位

渲染最终格式

第七步：检查结果

处理完成后，请重点检查：

新增的标准标注是否正确

红色括号是否为原来的错误内容

后追加的正确括号是否与 Excel 一致

黄色高亮部分是否已被避让

系统处理逻辑（详细说明）

本系统采用多阶段流程处理文档内容。

阶段一：审计

程序扫描文档中的中文植物名称，并检查其后是否已有括号内容。

阶段二：匹配判断

系统根据实际情况分为几类：

名称后没有括号 → 新增标准拉丁名

名称后已有正确括号 → 承认并跳过

名称后已有错误括号 → 标红旧括号并追加正确括号

名称位于黄色高亮中 → 直接跳过

名称属于长短嵌套关系中的短词 → 暂不处理，优先长词

阶段三：令牌锁定

系统并不是立刻将匹配到的文字替换成最终格式，而是先替换为内部令牌。
这样做的目的是避免：

文本位移

连续替换错位

长文档中格式混乱

“属名/种名”嵌套误识别

阶段四：最终渲染

最后再统一把令牌渲染为最终可见内容，例如：

车前科(Plantaginaceae)

谷精草科（Plantaginaceae）（Eriocaulaceae）

排版规则
中文部分

中文部分保持普通中文字体格式。

拉丁名部分

拉丁名部分会自动设置为：

Times New Roman

斜体规则

对于属级和种级的拉丁名，系统会自动应用斜体。

不斜体例外

以下植物学命名中的标记会自动设置为非斜体：

var.

subsp.

f.

×

ssp.

这样可更符合植物学文献的常见排版规范。

首次出现原则

本系统默认只处理每个植物名在文档中的首次有效出现。

这样设计的目的在于：

防止同一个植物名在全文反复重复标注

让正文更简洁

更适合学术写作与名录整理

即使同一名称在全文出现多次，通常也只处理第一次。

黄色高亮避让规则

若目标植物名位于黄色高亮文本中，系统会自动跳过。

该功能适合以下情况：

某些段落正在人工校对

某些区域不希望被自动修改

有些内容需要手工判断处理

当前版本中，黄色高亮被视为明确避让区域。

括号审计规则

系统不会简单地“看到括号就跳过”，而是会进一步检查括号中的内容。

同时，系统支持植物名与括号之间存在空格的情况，例如：

车前科(Plantaginaceae)
车前科 (Plantaginaceae)
车前科    (Plantaginaceae)
车前科（Plantaginaceae）

只要括号内容与 Excel 中标准值一致，系统就会承认该处已经正确标注。

如果括号内容与 Excel 不一致，则：

原括号会保留

原括号字体标红

正确值将作为新括号追加在后面

推荐使用流程

为了更安全地处理正式文档，建议按以下顺序操作：

整理并检查 Excel 名录库

先备份原始 Word 文档

导入 VBA 模块与窗体

运行宏

检查所有红色括号

必要时手工确认修正结果

将处理后的文档另存为新版本

这套流程特别适合：

植物志资料整理

学术论文写作

名录校核

分类学记录整理

常见问题与解决方法
问题 1：运行后没有任何反应

可能原因：

运行了错误的宏入口

当前活动文档不正确

frmProgress 窗体缺失

窗体控件名称与代码不一致

请检查：

窗体名称是否为 frmProgress

控件名称是否为：

lblStatus

ProgressBarBack

ProgressBarFore

运行的过程名是否为：

AddLatinNames_Turbo_V4_1_Final

问题 2：窗体出现了，但进度条不动

可能原因：

ProgressBarBack 不存在

ProgressBarFore 未正确配置

控件名称拼写不一致

问题 3：没有任何植物名被标注

可能原因：

Excel 文件为空

Excel 列序不符合要求

Word 文档中没有匹配到这些中文名

对应内容位于黄色高亮中

文中原有括号已被系统判断为有效

问题 4：错误括号没有被修正

可能原因：

文本中括号结构比较特殊

中间夹杂隐藏字符

Excel 中标准值存在拼写差异

文档格式过于复杂，导致定位不稳定

适用场景

本工具适合用于：

植物名录整理

植物志资料编写

分类学说明文稿

学术论文植物名标注

资源调查与统计文档

教学参考资料制作

中文植物名与拉丁名对应批处理

当前版本

当前公开版本：V4.1

使用说明补充

本项目属于科研辅助性质工具。
在正式出版、投稿或归档前，请务必再次人工核对：

Excel 名录数据是否准确

红色括号修正是否合理

属名、种名与分类层级是否匹配

自动化工具可以显著提升效率，但最终学术准确性仍需人工把关。

作者

Chris Bangle

仓库结构示例
Plant-List-Automated-Labeling-System-V4.1/
├─ Module1.bas
├─ frmProgress.frm
├─ frmProgress.frx
└─ README.md
