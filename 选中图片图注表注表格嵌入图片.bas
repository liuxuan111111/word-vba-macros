Attribute VB_Name = "模块1"
Sub 选中图片图注表注表格嵌入图片()
    Dim doc As Document
    Dim para As Paragraph
    Dim tbl As Table
    Dim shpInline As InlineShape
    Application.ScreenUpdating = False
    Set doc = ActiveDocument
    
    '清空原有选中
    Selection.Collapse wdCollapseEnd
    
    '1、选中样式：图片、图注、表注
    For Each para In doc.Paragraphs
        If para.style.NameLocal = "图片" Or para.style.NameLocal = "图注" Or para.style.NameLocal = "表注" Then
            para.Range.Select
            '扩展选区，保留之前选中内容
            Selection.Extend
        End If
    Next para
    
    '2、选中全部表格
    For Each tbl In doc.Tables
        tbl.Range.Select
        Selection.Extend
    Next tbl
    
    '3、选中所有嵌入图片
    For Each shpInline In doc.InlineShapes
        If shpInline.Type = wdInlineShapePicture Then
            shpInline.Range.Select
            Selection.Extend
        End If
    Next shpInline
    
    Application.ScreenUpdating = True
    MsgBox "执行完成：已选中【图片、图注、表注】样式段落 + 全部表格 + 全部嵌入图片"
End Sub
