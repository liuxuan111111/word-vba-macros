Attribute VB_Name = "模块5"
Sub 导出引用位置()
    Dim fld As Field
    Dim outText As String
    Dim bmkName As String
    Dim codeTxt As String
    Dim refRng As Range
    Dim pageNo As String
    Dim seq As Long
    
    Application.ScreenUpdating = False
    outText = "流水序号" & vbTab & "书签名称" & vbTab & "文献编号" & vbTab & "所在页码" & vbCrLf
    seq = 0
    
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldRef Then
            codeTxt = fld.code.Text
            If InStr(1, codeTxt, "_Ref", vbTextCompare) > 0 Then
                bmkName = Trim(Mid(codeTxt, InStr(1, codeTxt, "_Ref"), 20))
                
                fld.Update
                Set refRng = fld.Result
                pageNo = refRng.Information(wdActiveEndAdjustedPageNumber)
                
                seq = seq + 1
                outText = outText & seq & vbTab & bmkName & vbTab & _
                          fld.Result.Text & vbTab & pageNo & vbCrLf
            End If
        End If
    Next fld
    
    Documents.Add
    Selection.TypeText outText
    Application.ScreenUpdating = True
    MsgBox "提取完成！所有引用位置已导出"
End Sub

