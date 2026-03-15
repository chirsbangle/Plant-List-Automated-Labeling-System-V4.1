# Plant-List-Automated-Labeling-System-V4.1
Word VBA tool for automatically annotating Chinese plant names with Latin names from an Excel glossary.
Option Explicit

' ===========================================================================
' 系统名称：植物名录自动化标注系统 (Turbo Pro V4.1)
' 版权所有：(C) 2024-2026 [Chris Bangle]
' 软件性质：专业学术辅助工具 (Scientific Research Utility)
' ===========================================================================
' 【本次更新日志 - 定制修改】
' 1. 紧凑模式：中文名与拉丁名之间不再加空格，改为“中文(Latin)”。
' 2. 增强审计：智能忽略植物名后的空格，准确识别已有拉丁名。
' 3. 校核模式：若已有括号内容与 Excel 不一致，则原括号标红，并追加正确括号。
' ===========================================================================
' 【核心技术亮点】
' 1. 审计过滤 -> 扫描 -> 锁定 -> 渲染：完整四阶段可视化反馈。
' 2. 智能审计 (Smart Content Audit): 自动比对括号内容，识别并承认已有标注。
' 3. 长度优先抢占 (Length-Priority): 彻底解决“种名”与“属名”嵌套识别难题。
' 4. 物理占位令牌 (Tokenization): 隔离排版，确保超长文档渲染不产生位移。
' 5. 专业排版引擎: 自动处理 var./subsp./f./× 等植物志特定不斜体规范。
' ===========================================================================

Type PlantEntity
    ChineseName As String
    LatinName As String
    Rank As Integer     ' 1=科, 2=属, 3=种
    Length As Integer
    TokenID As String
    IsMatched As Boolean
    NeedAppend As Boolean
End Type

' ==========================================================
' 1. 主入口程序
' ==========================================================
Sub AddLatinNames_Turbo_V4_1_Final()
    Dim doc As Document: Set doc = ActiveDocument
    Dim plantList() As PlantEntity
    Dim pCount As Long, excelPath As String
    
    Dim promptMsg As String
    promptMsg = "欢迎使用植物名录自动化标注系统 V4.1" & vbCrLf & _
                "------------------------------------------------" & vbCrLf & _
                "【版权信息】 (C) 2024-2026 [Chris Bangle]" & vbCrLf & vbCrLf & _
                "【Excel 库标准列序要求】" & vbCrLf & _
                "A: 科中 | B: 科拉" & vbCrLf & _
                "C: 属中 | D: 属拉" & vbCrLf & _
                "E: 种中 | F: 种拉" & vbCrLf & vbCrLf & _
                "【功能说明】" & vbCrLf & _
                "1. 每个物种仅处理首次出现" & vbCrLf & _
                "2. 输出格式：中文(拉丁名)" & vbCrLf & _
                "3. 自动跳过黄色高亮内容" & vbCrLf & _
                "4. 已有括号时会校核内容是否与 Excel 一致" & vbCrLf & _
                "5. 若不一致：原括号标红，并追加正确括号" & vbCrLf & vbCrLf & _
                "是否立即开始？"
    
    If MsgBox(promptMsg, vbYesNo + vbQuestion, "Turbo Pro V4.1") = vbNo Then Exit Sub
    
    ' --- B. 文件加载 ---
    excelPath = PickExcelFile()
    If excelPath = "" Then Exit Sub
    LoadAndSortData excelPath, plantList, pCount
    
    If pCount = 0 Then
        MsgBox "Excel 中没有读取到有效数据。", vbExclamation
        Exit Sub
    End If
    
    ' --- C. 进度窗体 ---
    On Error Resume Next
    frmProgress.Show vbModeless
    frmProgress.lblStatus.Caption = "准备开始..."
    frmProgress.ProgressBarFore.Width = 0
    DoEvents
    On Error GoTo 0
    
    ' --- D. 阶段一：全局锁定 / 校核 ---
    Application.ScreenUpdating = False
    
    Dim globalDoneDict As Object: Set globalDoneDict = CreateObject("Scripting.Dictionary")
    Dim auditCount As Long: auditCount = 0
    Dim lockCount As Long: lockCount = 0
    Dim amendCount As Long: amendCount = 0
    
    Dim i As Long
    For i = 1 To pCount
        If i Mod 20 = 0 Or i = pCount Then
            On Error Resume Next
            frmProgress.lblStatus.Caption = "审计与锁定：" & plantList(i).ChineseName & "  (" & i & "/" & pCount & ")"
            frmProgress.ProgressBarFore.Width = (i / pCount) * (frmProgress.ProgressBarBack.Width * 0.5)
            DoEvents
            On Error GoTo 0
        End If
        
        If Not globalDoneDict.Exists(plantList(i).ChineseName) Then
            Dim rng As Range: Set rng = doc.Content
            With rng.Find
                .ClearFormatting
                .Text = plantList(i).ChineseName
                .Forward = True
                .Wrap = wdFindStop
                
                Do While .Execute
                    Dim status As Integer
                    status = FinalAudit_V4_1(rng, plantList(i), doc)
                    
                    If status = 1 Then
                        ' 没有现成括号，可以正常新加
                        rng.Text = plantList(i).TokenID
                        plantList(i).IsMatched = True
                        plantList(i).NeedAppend = False
                        globalDoneDict.Add plantList(i).ChineseName, True
                        lockCount = lockCount + 1
                        Exit Do
                        
                    ElseIf status = 2 Then
                        ' 已有正确标注，直接承认
                        plantList(i).IsMatched = False
                        plantList(i).NeedAppend = False
                        globalDoneDict.Add plantList(i).ChineseName, True
                        auditCount = auditCount + 1
                        Exit Do
                        
                    ElseIf status = 3 Then
                        ' 已有括号但内容不一致：已标红，后续只追加正确括号
                        rng.Text = plantList(i).TokenID
                        plantList(i).IsMatched = True
                        plantList(i).NeedAppend = True
                        globalDoneDict.Add plantList(i).ChineseName, True
                        amendCount = amendCount + 1
                        Exit Do
                        
                    Else
                        rng.Collapse wdCollapseEnd
                    End If
                Loop
            End With
        End If
    Next i
    
    ' --- E. 阶段二：格式化渲染 ---
    Dim renderIdx As Long: renderIdx = 0
    
    For i = 1 To pCount
        If plantList(i).IsMatched Then
            If renderIdx Mod 10 = 0 Or i = pCount Then
                On Error Resume Next
                frmProgress.lblStatus.Caption = "正在渲染格式..." & "  (" & i & "/" & pCount & ")"
                frmProgress.ProgressBarFore.Width = (frmProgress.ProgressBarBack.Width * 0.5) + _
                                                    (i / pCount) * (frmProgress.ProgressBarBack.Width * 0.5)
                DoEvents
                On Error GoTo 0
            End If
            
            Dim rToken As Range: Set rToken = doc.Content
            With rToken.Find
                .ClearFormatting
                .Text = plantList(i).TokenID
                .Forward = True
                .Wrap = wdFindStop
                
                If .Execute Then
                    renderIdx = renderIdx + 1
                    
                    If plantList(i).NeedAppend Then
                        ' 已存在错误括号，追加正确括号
                        rToken.Text = plantList(i).ChineseName & "§APPEND§"
                        AppendCorrectLatinAfterExistingBracket rToken, plantList(i)
                    Else
                        ' 正常新标注
                        rToken.Text = plantList(i).ChineseName & "(" & plantList(i).LatinName & ")"
                        ApplyFormatting_V4_1 rToken, plantList(i)
                    End If
                End If
            End With
        End If
    Next i
    
    Application.ScreenUpdating = True
    On Error Resume Next
    Unload frmProgress
    On Error GoTo 0
    
    MsgBox "【任务顺利完成】" & vbCrLf & vbCrLf & _
           "智能识别跳过：" & auditCount & " 处" & vbCrLf & _
           "新增标准标注：" & (renderIdx - amendCount) & " 处" & vbCrLf & _
           "发现并修正不一致：" & amendCount & " 处", vbInformation
End Sub

' ==========================================================
' 2. 深度审计逻辑
' 返回值：
' 0 = 跳过
' 1 = 可以锁定并新建标准括号
' 2 = 已有正确标注，承认并跳过
' 3 = 已有错误括号，已标红，后续追加正确括号
' ==========================================================
Function FinalAudit_V4_1(ByRef rng As Range, ByRef info As PlantEntity, ByRef doc As Document) As Integer
    ' 1. 避让黄色高亮
    If rng.HighlightColorIndex = wdYellow Then
        FinalAudit_V4_1 = 0
        Exit Function
    End If
    
    ' 2. 读取后缀文本
    Dim trailRange As Range
    Set trailRange = doc.Range(rng.End, IIf(rng.End + 80 > doc.Content.End, doc.Content.End, rng.End + 80))
    
    Dim trailText As String
    trailText = trailRange.Text
    
    ' 3. 忽略植物名后的空格
    Dim cleanTrail As String
    cleanTrail = LTrim(trailText)
    
    If Len(cleanTrail) > 0 Then
        Dim firstChar As String
        firstChar = Left(cleanTrail, 1)
        
        If firstChar = "(" Or firstChar = "（" Then
            Dim closeChar As String
            closeChar = IIf(firstChar = "(", ")", "）")
            
            Dim endPos As Long
            endPos = InStr(cleanTrail, closeChar)
            
            If endPos > 1 Then
                Dim existingLatin As String
                existingLatin = Mid(cleanTrail, 2, endPos - 2)
                
                If LCase(Trim(existingLatin)) = LCase(Trim(info.LatinName)) Then
                    ' 已有正确标注
                    FinalAudit_V4_1 = 2
                    Exit Function
                Else
                    ' 括号存在但内容不一致：把括号内容标红
                    ColorExistingBracketRed rng, doc
                    FinalAudit_V4_1 = 3
                    Exit Function
                End If
            Else
                FinalAudit_V4_1 = 0
                Exit Function
            End If
        End If
    End If
    
    ' 4. 嵌套避让
    If Len(trailText) > 0 Then
        Dim nextC As String
        nextC = Left(trailText, 1)
        
        If info.Rank = 3 And (nextC = "属" Or nextC = "科") Then
            FinalAudit_V4_1 = 0
            Exit Function
        End If
        
        If info.Rank = 2 And nextC = "科" Then
            FinalAudit_V4_1 = 0
            Exit Function
        End If
    End If
    
    FinalAudit_V4_1 = 1
End Function

' ==========================================================
' 3. 将已有错误括号内容标红
' ==========================================================
Sub ColorExistingBracketRed(ByRef nameRng As Range, ByRef doc As Document)
    Dim trailRange As Range
    Set trailRange = doc.Range(nameRng.End, IIf(nameRng.End + 80 > doc.Content.End, doc.Content.End, nameRng.End + 80))
    
    Dim rawText As String
    rawText = trailRange.Text
    
    Dim offsetSpaces As Long
    offsetSpaces = Len(rawText) - Len(LTrim(rawText))
    
    Dim cleanText As String
    cleanText = LTrim(rawText)
    
    If Len(cleanText) = 0 Then Exit Sub
    
    Dim openChar As String
    openChar = Left(cleanText, 1)
    
    If openChar <> "(" And openChar <> "（" Then Exit Sub
    
    Dim closeChar As String
    closeChar = IIf(openChar = "(", ")", "）")
    
    Dim endPos As Long
    endPos = InStr(cleanText, closeChar)
    If endPos <= 0 Then Exit Sub
    
    Dim bracketStart As Long
    Dim bracketEnd As Long
    
    bracketStart = nameRng.End + offsetSpaces
    bracketEnd = bracketStart + endPos
    
    Dim rBad As Range
    Set rBad = doc.Range(bracketStart, bracketEnd)
    rBad.Font.Color = wdColorRed
End Sub

' ==========================================================
' 4. 对正常新增的 中文(Latin) 应用格式
' ==========================================================
Sub ApplyFormatting_V4_1(ByRef rng As Range, ByRef info As PlantEntity)
    rng.Font.Name = "宋体"
    rng.Font.Italic = False
    
    Dim latStart As Long: latStart = rng.Start + Len(info.ChineseName) + 1
    Dim latEnd As Long: latEnd = rng.End - 1
    
    If latEnd > latStart Then
        Dim rLat As Range: Set rLat = rng.Document.Range(latStart, latEnd)
        rLat.Font.Name = "Times New Roman"
        
        If info.Rank >= 2 Then
            rLat.Font.Italic = True
            
            Dim v
            For Each v In Array("var.", "subsp.", "f.", "×", "ssp.")
                If InStr(rLat.Text, v) > 0 Then
                    Dim rPart As Range: Set rPart = rLat.Duplicate
                    With rPart.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .Text = v
                        .Replacement.Text = v
                        .Replacement.Font.Italic = False
                        .Forward = True
                        .Wrap = wdFindStop
                        .Format = True
                        .Execute Replace:=wdReplaceAll
                    End With
                End If
            Next v
        End If
    End If
End Sub

' ==========================================================
' 5. 对已有错误括号的条目，追加正确括号
' 结果示例：谷精草科（Plantaginaceae）（Eriocaulaceae）
' ==========================================================
Sub AppendCorrectLatinAfterExistingBracket(ByRef rng As Range, ByRef info As PlantEntity)
    Dim doc As Document
    Set doc = rng.Document
    
    ' 把占位文本先还原成中文名
    rng.Text = info.ChineseName
    
    ' 在名称后面查找第一个已有括号，并在其后追加正确括号
    Dim trailRange As Range
    Set trailRange = doc.Range(rng.End, IIf(rng.End + 120 > doc.Content.End, doc.Content.End, rng.End + 120))
    
    Dim rawText As String
    rawText = trailRange.Text
    
    Dim offsetSpaces As Long
    offsetSpaces = Len(rawText) - Len(LTrim(rawText))
    
    Dim cleanText As String
    cleanText = LTrim(rawText)
    
    If Len(cleanText) = 0 Then
        ' 兜底：如果没找到旧括号，就直接补标准括号
        Dim rAppend0 As Range
        Set rAppend0 = doc.Range(rng.End, rng.End)
        rAppend0.Text = "(" & info.LatinName & ")"
        ApplyLatinOnlyFormatting rAppend0, info
        Exit Sub
    End If
    
    Dim openChar As String
    openChar = Left(cleanText, 1)
    
    If openChar <> "(" And openChar <> "（" Then
        ' 兜底：没括号则直接补
        Dim rAppend1 As Range
        Set rAppend1 = doc.Range(rng.End, rng.End)
        rAppend1.Text = "(" & info.LatinName & ")"
        ApplyLatinOnlyFormatting rAppend1, info
        Exit Sub
    End If
    
    Dim closeChar As String
    closeChar = IIf(openChar = "(", ")", "）")
    
    Dim endPos As Long
    endPos = InStr(cleanText, closeChar)
    If endPos <= 0 Then
        Dim rAppend2 As Range
        Set rAppend2 = doc.Range(rng.End, rng.End)
        rAppend2.Text = "(" & info.LatinName & ")"
        ApplyLatinOnlyFormatting rAppend2, info
        Exit Sub
    End If
    
    Dim insertPos As Long
    insertPos = rng.End + offsetSpaces + endPos
    
    Dim rAppend As Range
    Set rAppend = doc.Range(insertPos, insertPos)
    rAppend.Text = "(" & info.LatinName & ")"
    ApplyLatinOnlyFormatting rAppend, info
End Sub

' ==========================================================
' 6. 仅对追加出来的新括号应用格式
' ==========================================================
Sub ApplyLatinOnlyFormatting(ByRef rng As Range, ByRef info As PlantEntity)
    rng.Font.Name = "宋体"
    rng.Font.Italic = False
    
    Dim rLat As Range
    Set rLat = rng.Document.Range(rng.Start + 1, rng.End - 1)
    
    If rLat.End > rLat.Start Then
        rLat.Font.Name = "Times New Roman"
        
        If info.Rank >= 2 Then
            rLat.Font.Italic = True
            
            Dim v
            For Each v In Array("var.", "subsp.", "f.", "×", "ssp.")
                If InStr(rLat.Text, v) > 0 Then
                    Dim rPart As Range
                    Set rPart = rLat.Duplicate
                    With rPart.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .Text = v
                        .Replacement.Text = v
                        .Replacement.Font.Italic = False
                        .Forward = True
                        .Wrap = wdFindStop
                        .Format = True
                        .Execute Replace:=wdReplaceAll
                    End With
                End If
            Next v
        End If
    End If
End Sub

' ==========================================================
' 7. 加载与辅助
' ==========================================================
Sub LoadAndSortData(path As String, ByRef list() As PlantEntity, ByRef count As Long)
    Dim xl As Object, wb As Object, sh As Object
    Set xl = CreateObject("Excel.Application")
    Set wb = xl.Workbooks.Open(path, ReadOnly:=True)
    Set sh = wb.Sheets(1)
    
    Dim lastRow As Long: lastRow = sh.UsedRange.Rows.Count
    ReDim list(1 To lastRow * 3)
    
    count = 0
    Dim i As Long, k As Integer
    Dim cols: cols = Array(1, 3, 5)
    
    For i = 2 To lastRow
        For k = 0 To 2
            Dim cn As String: cn = Trim(CStr(sh.Cells(i, cols(k)).Value))
            Dim ln As String: ln = Trim(CStr(sh.Cells(i, cols(k) + 1).Value))
            
            If cn <> "" And ln <> "" Then
                count = count + 1
                list(count).ChineseName = cn
                list(count).LatinName = ln
                list(count).Rank = k + 1
                list(count).Length = Len(cn)
                list(count).TokenID = "§" & Choose(k + 1, "K", "G", "S") & Format(count, "00000") & "§"
                list(count).IsMatched = False
                list(count).NeedAppend = False
            End If
        Next k
    Next i
    
    wb.Close False
    xl.Quit
    
    Dim j As Long, temp As PlantEntity
    For i = 1 To count - 1
        For j = i + 1 To count
            If list(j).Length > list(i).Length Then
                temp = list(i)
                list(i) = list(j)
                list(j) = temp
            End If
        Next j
    Next i
End Sub

Function PickExcelFile() As String
    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "选择 Excel 库"
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx;*.xls"
        If .Show = -1 Then PickExcelFile = .SelectedItems(1)
    End With
End Function
