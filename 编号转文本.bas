Attribute VB_Name = "NewMacros"
Sub 编号转文本()
'
' 编号转文本 宏
'
'
Dim kgslist As List
For Each kgslist In ActiveDocument.Lists
kgslist.ConvertNumbersToText
Next
End Sub
Sub 图改图片样式()
Attribute 图改图片样式.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.宏2"
'
' 宏2 宏
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.style = ActiveDocument.Styles("图片")
    With Selection.Find
        .Text = "^g"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
