Attribute VB_Name = "modHTMLEditor"
'
' Color HTML Tags from Rich Textbox Control by Using this
' Two Quick Sub-Routine without using APIs
'
' Created By : Chris Vega [gwapo@models.com]
'

'
' USAGE:
' ====================================
'
'  1. Initialize Color value for the two variables
'
'       HTML_Color         = Color of the Tags
'       Comment_Color      = Color of the Comments
'
'     with the Color you desired as the initial colors
'     that the procedures will be using.

Public HTML_Color As Long
Public Comment_Color As Long


'
'  2. Use the function to Color the RichTextBox passed within
'     this procedure, for better result (without or less flickering)
'     use a temporary RichTextBox Control.
'
'     where eMidPoint must be set as, 0 if you want to reColor the
'     entire RichTextBox Control, or Higher than 0 (normally, the
'     Current Cursor Location) to partially Color RichTextBox
'

Public Sub ColorTags(rText As RichTextBox, eMidPoint)
    If eMidPoint = 0 Then
        ' Entire Content Coloring
        st = 0
        en = Len(rText.Text)
        
        ' unColor Entire RichTextBox
        ClearColors rText, 0
    Else
        ' Partial Coloring
        
        ' search for the nearest "<" character and mark it as
        ' starting point (search direction is negative)
        st = eMidPoint - 100
        If st > 0 Then
            Do While Not st = 0
                If Mid(rText.Text, st, 1) = "<" Then _
                Exit Do
                st = st - 1
            Loop
        End If
        st = st - 1
        If st < 0 Then st = 0
        
        ' search for the nearest ">" character and mark it as
        ' ending point (search direction is positive)
        
        en = InStr(eMidPoint, rText.Text, ">")
        If en = 0 Then en = Len(rText.Text)
        
        ' unColor the Location (reset)
        ClearColors rText, eMidPoint
    End If
    
    ' Infinity Loop
    While True
        ' Use the Find Method of RichTextBox Control for quickier
        ' Searching of Character "<" and ">"
        ltTag = rText.Find(Chr(60), st)
        gtTag = rText.Find(Chr(62), ltTag)
        tgLen = (gtTag - ltTag) + 1
        
        ' Is there no matching Tags to Color found?
        '       Break the Loop
        If (tgLen) = 0 Or _
           ltTag < 0 Or _
           gtTag < 0 Then Exit Sub
        
        ' Otherwise, Do the Coloring
        With rText
            .SelStart = ltTag
            .SelLength = tgLen
            ' Comment Tag Color
            If Left(.SelText, 2) = "<!" Then _
            .SelColor = Comment_Color _
            Else _
            .SelColor = HTML_Color
            ' HTML Tag Color
        End With
        
        ' Increase Starting Point
        st = ltTag + tgLen
        If st > en Then Exit Sub
    Wend
End Sub

'
'  3. This function resembles the above function to reverse the Coloring.
'
'     where eMidPoint must be set as, 0 if you want to unColor the
'     entire RichTextBox Control, or Higher than 0 (normally, the
'     Current Cursor Location) to partially unColor RichTextBox
'

Public Sub ClearColors(rText As RichTextBox, eMidPoint)
    If eMidPoint = 0 Then
        st = 0
        en = Len(rText.Text)
    Else
        st = eMidPoint - 100
        If st > 0 Then
            Do While Not st = 0
                If Mid(rText.Text, st, 1) = "<" Then _
                Exit Do
                st = st - 1
            Loop
        End If
        st = st - 1
        If st < 0 Then st = 0
        
        en = InStr(eMidPoint, rText.Text, ">")
        If en = 0 Then en = Len(rText.Text)
    End If
    
    With rText
        .SelStart = st
        .SelLength = en - st
        .SelColor = &H0
    End With
End Sub


