Attribute VB_Name = "模块10"
Sub 角标()
    Dim doc As Document
    Dim p As Paragraph
    Dim text As String
    Dim startPos As Long
    Dim endPos As Long
    Dim midStr As String
    
    Dim currentLevel As Integer
    Dim currentLevelNumber As String
    Dim currentLevelText As String
    
    ' 用于精确定位和获取页码的Range对象
    Dim rngChar As Range
    Dim pageNum As Long
    
    Dim resultDoc As Document
    Dim counter As Long
    
    Set doc = ActiveDocument
    Set resultDoc = Documents.Add ' 新建一个文档用于输出结果
    
    ' 添加表头
    With resultDoc.content
        .text = "序号" & vbTab & "角标编号" & vbTab & "最小级别章节号" & vbTab & "章节标题" & vbTab & "所在页码" & vbCr
        .Paragraphs(1).Range.Font.Bold = True
        .Paragraphs(1).Range.InsertParagraphAfter
    End With
    
    counter = 0
    currentLevel = 0
    currentLevelNumber = "无"
    currentLevelText = "文档开头(无标题)"
    
    ' 遍历所有段落
    For Each p In doc.Paragraphs
        
        ' 第一步：检查段落是否为标题，并更新当前上下文（最小级别章节）
        If p.OutlineLevel >= 1 And p.OutlineLevel <= 9 Then
            currentLevel = p.OutlineLevel
            
            ' 获取自动编号（如 1.2.3），如果未启用则为空
            currentLevelNumber = p.Range.ListFormat.ListString
            If currentLevelNumber = "" Then
                ' 尝试提取手动编号
                currentLevelNumber = Trim(Left(p.Range.text, InStr(p.Range.text, " ") - 1))
            End If
            
            ' 清理标题文本
            currentLevelText = Replace(p.Range.text, vbCr, "")
            If IsNumeric(Left(currentLevelText, 1)) Then
               Dim i As Integer
               For i = 1 To Len(currentLevelText)
                   If Mid(currentLevelText, i, 1) = " " Then
                       currentLevelText = Mid(currentLevelText, i + 1)
                       Exit For
                   End If
               Next i
            End If
        End If
        
        ' 第二步：在段落文本中寻找 [数字] 的角标结构
        text = p.Range.text ' 获取段落文本，可能包含段落末尾的 vbCr
        
        ' 寻找第一个左括号 [
        startPos = InStr(1, text, "[")
        
        ' 循环查找直到段落结束（修正了语法：改为 Do While...Loop 配合 Exit Do）
        Do While startPos > 0
            ' 找到对应的右括号 ]
            endPos = InStr(startPos + 1, text, "]")
            
            If endPos > 0 Then
                ' 截取中间内容
                midStr = Mid(text, startPos + 1, endPos - startPos - 1)
                
                ' 判断中间是否为纯数字
                If IsNumeric(midStr) Then
                    counter = counter + 1
                    
                    ' 通过Range对象获取该角标所在的物理页码
                    Set rngChar = p.Range
                    ' 定位Range的起始位置（注意：InStr是从1开始，需要 -1 偏移）
                    rngChar.Start = p.Range.Start + (startPos - 1)
                    ' 定位Range的结束位置（包含右括号 ]
                    rngChar.End = p.Range.Start + endPos
                    ' 获取所在的最后一页页码
                    pageNum = rngChar.Information(wdActiveEndPageNumber)
                    
                    ' 输出到新文档：序号、角标编号、所属章节号、所属章节标题、所在页码
                    With resultDoc.content
                        .InsertAfter counter & vbTab & "[" & midStr & "]" & vbTab & currentLevelNumber & vbTab & currentLevelText & vbTab & "第 " & pageNum & " 页" & vbCr
                    End With
                End If
                
                ' 在当前段落中，从右括号之后继续寻找下一个左括号
                startPos = InStr(endPos + 1, text, "[")
            Else
                ' 没找到对应的右括号，防止死循环，直接退出当前段落循环
                Exit Do
            End If
        Loop
        
    Next p

    ' 提示完成
    MsgBox "扫描完成！共找到 " & counter & " 个形如 [数字] 的角标。页码及章节信息已输出到新文档。", vbInformation, "提取成功"
    
    ' 清理对象
    Set rngChar = Nothing
End Sub
