Attribute VB_Name = "模块7"
Sub 批量插图注表注并统一图片样式()
    Dim shp As InlineShape
    Dim capRng As Range
    Dim tbl As Table
    Dim tableCapRng As Range
    Dim picCount As Integer
    Dim tableCount As Integer
    Application.ScreenUpdating = False
    picCount = 0
    tableCount = 0
    
    '===================== 第一部分：批量处理所有嵌入图片 =====================
    '1. 给每张图片插入图注，图注应用【图注】段落样式（图注在图片下方）
    For Each shp In ActiveDocument.InlineShapes
        If shp.Type = wdInlineShapePicture Then
            shp.Select
            Selection.InsertCaption _
                Label:="图", _
                Position:=wdCaptionPositionBelow, _
                ExcludeLabel:=False
            Set capRng = Selection.Paragraphs.Last.Range
            capRng.style = ActiveDocument.Styles("图注")
            picCount = picCount + 1
        End If
    Next shp
    
    '2. 查找所有嵌入图片，批量应用样式【图片】（原宏2查找替换逻辑）
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
    
    '3. 处理浮动图片，同步图片样式
    Call SelectAllImages
    If Selection.Type <> wdSelectionNone Then
        ActiveDocument.InlineShapes(1).Select
        Selection.CopyFormat
        Call SelectAllImages
        Selection.PasteFormat
    End If
    
    '===================== 第二部分：批量处理所有表格，表注放在表格上方，应用【表注】样式 =====================
    For Each tbl In ActiveDocument.Tables
        tbl.Range.Select
        '关键修改：Position:=wdCaptionPositionAbove 表注在表格上方
        Selection.InsertCaption _
            Label:="表", _
            Position:=wdCaptionPositionAbove, _
            ExcludeLabel:=False
        '选中刚生成的表注，应用段落样式「表注」
        Set tableCapRng = Selection.Paragraphs.Last.Range
        tableCapRng.style = ActiveDocument.Styles("表注")
        tableCount = tableCount + 1
    Next tbl
    
    Application.ScreenUpdating = True
    MsgBox "执行完成！" & vbCrLf & _
           "图片：共" & picCount & "张，已添加图注并统一图片样式「图片」" & vbCrLf & _
           "表格：共" & tableCount & "个，表注置于表格上方，全部应用段落样式「表注」", vbInformation
End Sub

'你提供的子过程：选中全部浮动图片Shape
Sub SelectAllImages()
    Dim oShape As Shape
    For Each oShape In ActiveDocument.Shapes
        If oShape.Type = msoPicture Then
            oShape.Select Replace:=False
        End If
    Next oShape
End Sub

