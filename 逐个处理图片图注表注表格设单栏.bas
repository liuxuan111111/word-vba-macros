Attribute VB_Name = "模块3"
Sub 分栏()
    Dim doc As Document
    Dim para As Paragraph
    Dim tbl As Table
    Dim shpFloat As Shape
    Dim usableW As Single
    Application.ScreenUpdating = False
    Set doc = ActiveDocument
    
    '1、全文先统一设置为一栏，重置原有分栏格式
    doc.content.PageSetup.TextColumns.SetCount NumColumns:=1
    
    '2、全文正文整体设置为两栏
    doc.content.PageSetup.TextColumns.SetCount NumColumns:=2
    
    '3、遍历段落，选中样式为【图片】【图注】【表注】的段落，设置整段跨页单栏居中
    For Each para In doc.Paragraphs
        Select Case para.style.NameLocal
            Case "图片", "图注", "表注"
                para.Range.ParagraphFormat.PageBreakBefore = True
                para.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End Select
    Next para
    
    '4、所有表格设置跨页单栏居中
    For Each tbl In doc.Tables
        tbl.Range.ParagraphFormat.PageBreakBefore = True
        tbl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Next tbl
    
    '5、浮动图片页面水平居中
    usableW = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin
    For Each shpFloat In doc.Shapes
        If shpFloat.Type = msoPicture Then
            shpFloat.Left = doc.PageSetup.LeftMargin + (usableW - shpFloat.Width) / 2
        End If
    Next shpFloat
    
    Application.ScreenUpdating = True
    MsgBox "排版完成：正文文字两栏；样式【图片/图注/表注】、表格整体单栏整页居中", vbInformation
End Sub
