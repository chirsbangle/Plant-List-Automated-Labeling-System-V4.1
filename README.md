# Plant-List-Automated-Labeling-System-V4.1

A professional Word VBA utility for automatically annotating Chinese plant names with standardized Latin names from an Excel glossary. The system supports first-occurrence-only labeling, smart auditing of existing bracketed names, mismatch correction, highlight avoidance, token-based rendering, and scientific formatting for botanical nomenclature.

---

## Project Overview

Plant-List-Automated-Labeling-System-V4.1 is a Word VBA tool designed for the automatic annotation of Chinese plant names with standardized Latin names from an Excel glossary. It is intended for botanical checklists, floristic records, taxonomic notes, academic manuscripts, and other structured plant-related documents.

The system is built to improve consistency, reduce repetitive manual work, and support scientific writing workflows in Word documents. It can recognize existing annotations, avoid duplicate labeling, skip highlighted content, detect incorrect bracketed Latin names, and append correct forms according to the glossary.

---

## Copyright Notice

```text
===========================================================================
系统名称：植物名录自动化标注系统 (Turbo Pro V4.1)
版权所有：(C) 2024-2026 [Chirs Bangle]
软件性质：专业学术辅助工具 (Scientific Research Utility)
===========================================================================
【本次更新日志 - 定制修改】
1. 紧凑模式：中文名与拉丁名之间不再加空格，改为“中文(Latin)”。
2. 增强审计：智能忽略植物名后的空格，准确识别已有拉丁名。
3. 校核模式：若已有括号内容与 Excel 不一致，则原括号标红，并追加正确括号。
===========================================================================
【核心技术亮点】
1. 审计过滤 -> 扫描 -> 锁定 -> 渲染：完整四阶段可视化反馈。
2. 智能审计 (Smart Content Audit): 自动比对括号内容，识别并承认已有标注。
3. 长度优先抢占 (Length-Priority): 彻底解决“种名”与“属名”嵌套识别难题。
4. 物理占位令牌 (Tokenization): 隔离排版，确保超长文档渲染不产生位移。
5. 专业排版引擎: 自动处理 var./subsp./f./× 等植物志特定不斜体规范。
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
