Attribute VB_Name = "模块9"
Sub ListCrossRef_OnlyPage()
    Dim fld As Field
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim outText As String
    Dim bmkName As String
    Dim codeTxt As String
    Dim refRng As Range
    Dim pageNo As String
    
    Application.ScreenUpdating = False
    '表头
    outText = "序号" & vbTab & "书签名称" & vbTab & "当前文献编号" & vbTab & "首次出现页码" & vbCrLf
    Dim seq As Long
    seq = 0
    
    '遍历全文所有域
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldRef Then
            codeTxt = fld.code.Text
            '筛选Ref交叉引用域
            If InStr(1, codeTxt, "_Ref", vbTextCompare) > 0 Then
                bmkName = Trim(Mid(codeTxt, InStr(1, codeTxt, "_Ref"), 20))
                If Not seen.Exists(bmkName) Then
                    seen(bmkName) = True
                    seq = seq + 1
                    
                    fld.Update
                    Set refRng = fld.Result
                    
                    '获取页码
                    pageNo = refRng.Information(wdActiveEndAdjustedPageNumber)
                    
                    '拼接输出
                    outText = outText & seq & vbTab & bmkName & vbTab & _
                              fld.Result.Text & vbTab & pageNo & vbCrLf
                End If
            End If
        End If
    Next fld
    
    '新建文档输出清单
    Documents.Add
    Selection.TypeText outText
    Application.ScreenUpdating = True
    MsgBox "提取完成！新文档为正文首次引用顺序清单"
End Sub

