Attribute VB_Name = "模块6"
Sub 批量插图注()
    Dim shp As InlineShape
    Dim capRng As Range
    Dim picCount As Integer
    Application.ScreenUpdating = False
    picCount = 0
    
    '========== 第一步：批量给所有嵌入图片插入图注，图注应用段落样式【图注】 ==========
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
    
    '========== 第二步：查找所有嵌入图片，批量应用样式【图片】（你提供的宏2逻辑） ==========
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
    
    '========== 第三步：处理浮动图片，选中后同步样式 ==========
    Call SelectAllImages
    If Selection.Type <> wdSelectionNone Then
        '复制任意一张嵌入图的图片格式，粘贴给浮动图
        ActiveDocument.InlineShapes(1).Select
        Selection.CopyFormat
        Call SelectAllImages
        Selection.PasteFormat
    End If
    
    Application.ScreenUpdating = True
    MsgBox "执行完成！共处理 " & picCount & " 张图片：" & vbCrLf & "1.全部图片添加图注并应用【图注】段落样式" & vbCrLf & "2.嵌入图片自动套用样式「图片」，浮动图片同步统一样式", vbInformation
End Sub
