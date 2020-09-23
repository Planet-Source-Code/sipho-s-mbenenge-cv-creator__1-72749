Attribute VB_Name = "ModCvCreate"
Global nwcCVFields As New Collection
Global nwcDataEntry As New Collection
Global nwcHeaders As New Collection

Global nwcPersonal As New Collection
Global nwcEducational As New Collection
Global nwcTertiary As New Collection
Global nwcOther As New Collection
Global nwcVoluntary As New Collection
Global nwcWork As New Collection
Global nwcReference As New Collection

Global nwcPerUnselected As New Collection
Global nwcEduUnselected As New Collection
Global nwcTerUnselected As New Collection
Global nwcOthUnselected As New Collection
Global nwcVolUnselected As New Collection
Global nwcWorUnselected As New Collection
Global nwcRefUnselected As New Collection

Global nwcPersonalAddress As New Collection
Global nwcEducationalSubjects As New Collection
Global nwcTertiarySubjects As New Collection
Global nwcTertiary1Subjects As New Collection
Global nwcTertiary2Subjects As New Collection
Global nwcVoluntaryAddress As New Collection
Global nwcWorkAddress As New Collection
Global nwcReferenceAddress As New Collection

Global blnPersonal As Boolean
Global blnEducational As Boolean
Global blnTertiary As Boolean
Global blnOther As Boolean
Global blnVoluntary As Boolean
Global blnWork As Boolean
Global blnReference As Boolean

Global blnDesign As Boolean
Global blnCreate As Boolean

Global blnSave As Boolean
Global blnOpen As Boolean
Global blnPrint As Boolean
Global blnClose As Boolean

Global strDesignType As String

Dim CVcreater As Object
Dim intReply As Integer

Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    
    Load frmMain
    frmMain.Show
    
    Load frmCVHeaders
    frmCVHeaders.Show vbModal
    
    Unload frmSplash
End Sub

Sub Personal_Fields()
With frmDisplay
    
    If blnPersonal = True Then
    'To Do: Set the caption of the labels according
    '       to the way the user specified them.
        For a = 0 To nwcPersonal.Count
            If a >= nwcPersonal.Count Then GoTo 100
            With frmDisplay.Label1(a)
            .Caption = nwcPersonal.Item(a + 1)
            .ForeColor = &HC0FFFF
            If .Caption = "Postal Address" Then
            With frmDisplay.cmdAddPAddress
                .Top = frmDisplay.Label1(a).Top + 30
                .Visible = True
            End With
            End If
            End With
            With frmDisplay.Text1(a)
            .Text = "Type your " & nwcPersonal.Item(a + 1) & " here!"
            .Locked = False
            End With
        Next a
        'To Do: Set the default button
        frmDisplay.cmdAddPAddress.Default = True
100:
    Z = 0: Y = 1
    For Z = 17 To 0 Step -1
    If frmDisplay.Text1(Z).Text = "Field Not Selected" Then
    frmDisplay.Label1(Z).Caption = nwcPerUnselected.Item(Y)
    Y = Y + 1
    End If
    Next Z
    
    Else
'        lblPersonal.ForeColor = vbRed
    End If
    
    'To Do: Hide other tabs
    .picPersonal.Visible = True
    .picEducational.Visible = False
    .picTertiary.Visible = False
    .picOther.Visible = False
    .picVoluntary.Visible = False
    .picWork.Visible = False
    .picReference.Visible = False
End With
End Sub

Sub Educational_Fields()
With frmDisplay

If blnEducational = True Then

'To Do: Set the caption of the labels according
'       to the way the user specified them.
    For a = 0 To nwcEducational.Count
        If a >= nwcEducational.Count Then GoTo 200
        With frmDisplay.Label2(a)
        .Caption = nwcEducational.Item(a + 1)
        .ForeColor = &HC0FFFF
        If .Caption = "Subjects" Then
        With frmDisplay.cmdEduSub
            .Top = frmDisplay.Label2(a).Top + 30
            .Visible = True
        End With
        End If
        End With
        With frmDisplay.Text2(a)
        .Text = "Type your " & nwcEducational.Item(a + 1) & " here!"
        .Locked = False
        End With
    Next a
200:
    Z = 0: Y = 1
    For Z = 4 To 0 Step -1
    If frmDisplay.Text2(Z).Text = "Field Not Selected" Then
    frmDisplay.Label2(Z).Caption = nwcEduUnselected.Item(Y)
    Y = Y + 1
    End If
    Next Z
    
Else
'    lblEducational.ForeColor = vbRed
End If

    'To Do: Hide other tabs
    .picEducational.Visible = True
    .picPersonal.Visible = False
    .picTertiary.Visible = False
    .picOther.Visible = False
    .picVoluntary.Visible = False
    .picWork.Visible = False
    .picReference.Visible = False
End With
End Sub

Sub Tertiary_Fields()
With frmDisplay

If blnTertiary = True Then

'To Do: Set the caption of the labels according
'       to the way the user specified them.
    For a = 0 To nwcTertiary.Count
        If a >= nwcTertiary.Count Then GoTo 300
        With frmDisplay.Label3(a)
        .Caption = nwcTertiary.Item(a + 1)
        .ForeColor = &HC0FFFF

'To Do: Set and align the subject command buttons
'       if they are selected.
        If .Caption = "Subjects" Then
        With frmDisplay.cmdTerSub
            .Top = frmDisplay.Label3(a).Top + 30
            .Visible = True
        End With
        End If
        
        If .Caption = "1st Year Subjects Completed" Then
        With frmDisplay.cmdTer1Sub
            .Top = frmDisplay.Label3(a).Top + 30
            .Visible = True
        End With
            .Caption = "1st Year Subjects Co..."
        End If
        
        If .Caption = "2nd Year Subjects Completed" Then
        With frmDisplay.cmdTer2Sub
            .Top = frmDisplay.Label3(a).Top + 30
            .Visible = True
        End With
        .Caption = "2nd Year Subjects Co..."
        End If
'End To Do.
'        End If
        End With
        With frmDisplay.Text3(a)
        .Text = "Type your " & nwcTertiary.Item(a + 1) & " here!"
        .Locked = False
        End With
    Next a
300:
    Z = 0: Y = 1
    For Z = 4 To 0 Step -1
    If frmDisplay.Text3(Z).Text = "Field Not Selected" Then
    frmDisplay.Label3(Z).Caption = nwcTertiary.Item(Y)
    Y = Y + 1
    End If
    Next Z
    
Else
'    lblTertiary.ForeColor = vbRed
End If

    'To Do: Hide other tabs
    .picTertiary.Visible = True
    .picEducational.Visible = False
    .picPersonal.Visible = False
    .picOther.Visible = False
    .picVoluntary.Visible = False
    .picWork.Visible = False
    .picReference.Visible = False
End With
End Sub

Sub Other_Fields()
With frmDisplay
    
If blnOther = True Then

'To Do: Set the caption of the labels according
'       to the way the user specified them.
    For a = 0 To nwcOther.Count
        If a >= nwcOther.Count Then GoTo 400
        With frmDisplay.Label4(a)
        .Caption = nwcOther.Item(a + 1)
        .ForeColor = &HC0FFFF
        If frmDisplay.Label4(a).Caption = "Postal Address" Then
'        cmdAddPAddress.Top = Label1(a).Top + 80
        End If
        End With
        With frmDisplay.Text4(a)
        .Text = "Type your " & nwcOther.Item(a + 1) & " here!"
        .Locked = False
        End With
    Next a
400:
    Z = 0: Y = 1
    For Z = 5 To 0 Step -1
    If frmDisplay.Text4(Z).Text = "Field Not Selected" Then
    frmDisplay.Label4(Z).Caption = nwcOthUnselected.Item(Y)
    Y = Y + 1
    End If
    Next Z
    
Else
'    lblOther.ForeColor = vbRed
End If
    
    'To Do: Hide other tabs
    .picOther.Visible = True
    .picTertiary.Visible = False
    .picEducational.Visible = False
    .picPersonal.Visible = False
    .picVoluntary.Visible = False
    .picWork.Visible = False
    .picReference.Visible = False
End With
End Sub

Sub Voluntary_Fields()
With frmDisplay

If blnVoluntary = True Then

'To Do: Set the caption of the labels according
'       to the way the user specified them.
    For a = 0 To nwcVoluntary.Count
        If a >= nwcVoluntary.Count Then GoTo 500
        With frmDisplay.Label5(a)
        .Caption = nwcVoluntary.Item(a + 1)
        .ForeColor = &HC0FFFF
        If .Caption = "Address of the Company" Then
        With frmDisplay.cmdVolAddress
            .Top = frmDisplay.Label5(a).Top + 30
            .Visible = True
        End With
        .Caption = "Address of the Comp..."
        End If
        If .Caption = "Title and name of contact" Then .Caption = "Title and Name of Con..."
        End With
        With frmDisplay.Text5(a)
        .Text = "Type your " & nwcVoluntary.Item(a + 1) & " here!"
        .Locked = False
        End With
    Next a
500:
    Z = 0: Y = 1
    For Z = 13 To 0 Step -1
    If frmDisplay.Text5(Z).Text = "Field Not Selected" Then
    frmDisplay.Label5(Z).Caption = nwcVolUnselected.Item(Y)
    Y = Y + 1
    End If
    Next Z
    
Else
'    lblVoluntary.ForeColor = vbRed
End If

    ' To Do: Hide other tabs
    .picVoluntary.Visible = True
    .picOther.Visible = False
    .picTertiary.Visible = False
    .picEducational.Visible = False
    .picPersonal.Visible = False
    .picWork.Visible = False
    .picReference.Visible = False
End With
End Sub

Sub Work_Fields()
With frmDisplay

If blnWork = True Then

'To Do: Set the caption of the labels according
'       to the way the user specified them.
    For a = 0 To nwcWork.Count
        If a >= nwcWork.Count Then GoTo 600
        With frmDisplay.Label6(a)
        .Caption = nwcWork.Item(a + 1)
        .ForeColor = &HC0FFFF
        If frmDisplay.Label6(a).Caption = "Address of the Company" Then
        With frmDisplay.cmdWorkAddress
            .Top = frmDisplay.Label6(a).Top + 30
            .Visible = True
        End With
        .Caption = "Address of the Com..."
        End If
        If .Caption = "Title and name of contact" Then .Caption = "Title and Name of Con..."
        End With
        With frmDisplay.Text6(a)
        .Text = "Type your " & nwcWork.Item(a + 1) & " here!"
        .Locked = False
        End With
    Next a
600:
    Z = 0: Y = 1
    For Z = 12 To 0 Step -1
    If frmDisplay.Text6(Z).Text = "Field Not Selected" Then
    frmDisplay.Label6(Z).Caption = nwcWorUnselected.Item(Y)
    If frmDisplay.Label6(Z).Caption = "Title and name of contact" Then frmDisplay.Label6(Z).Caption = "Title and Name of Con..."
    If frmDisplay.Label6(Z).Caption = "Address of the Company" Then frmDisplay.Label6(Z).Caption = "Address of the Com..."
    Y = Y + 1
    End If
    Next Z
    
Else
'    lblWork.ForeColor = vbRed
End If

    'To Do: Hide other tabs
    .picWork.Visible = True
    .picVoluntary.Visible = False
    .picOther.Visible = False
    .picTertiary.Visible = False
    .picEducational.Visible = False
    .picPersonal.Visible = False
    .picReference.Visible = False
End With
End Sub

Sub References_Fields()
With frmDisplay

If blnReference = True Then

'To Do: Set the caption of the labels according
'       to the way the user specified them.
    For a = 0 To nwcReference.Count
        If a >= nwcReference.Count Then GoTo 700
        With frmDisplay.Label7(a)
        .Caption = nwcReference.Item(a + 1)
        .ForeColor = &HC0FFFF
        If frmDisplay.Label7(a).Caption = "Address" Then
        With frmDisplay.cmdRefAddress
            .Top = frmDisplay.Label7(a).Top + 30
            .Visible = True
        End With
        End If
        If .Caption = "Title, name and surname" Then .Caption = "Title, Name and Sur..."
        End With
        With frmDisplay.Text7(a)
        .Text = "Type your " & nwcReference.Item(a + 1) & " here!"
        .Locked = False
        End With
    Next a
700:
    Z = 0: Y = 1
    For Z = 6 To 0 Step -1
    If frmDisplay.Text7(Z).Text = "Field Not Selected" Then
    frmDisplay.Label7(Z).Caption = nwcRefUnselected.Item(Y)
    Y = Y + 1
    End If
    Next Z
    
Else
'    lblReference.ForeColor = vbRed
End If

    'To Do: Hide other tabs
    .picWork.Visible = False
    .picVoluntary.Visible = False
    .picOther.Visible = False
    .picTertiary.Visible = False
    .picEducational.Visible = False
    .picPersonal.Visible = False
    .picReference.Visible = True
End With
End Sub

Sub Mbenenge_Design()

    Dim appWord As Word.Application
    Dim blnNewWord As Boolean

    On Error Resume Next
    Set appWord = GetObject(, "Word.Application")
    If Err.Number = 0 Then
        Debug.Print "Word running: Use currently running Word"
        blnNewWord = False
        Documents.Add DocumentType:=wdNewBlankDocument
    Else
        Debug.Print "Word not running: Create instance of Word"
        Set appWord = New Word.Application
        blnNewWord = True
    End If
    On Error GoTo 0
    ' Start of whatever one needs to do with Word
    With appWord
        If .Documents.Count = 0 Then
            .Documents.Add
        End If
        .Visible = True
    End With

'To Do: Start Word and prepare Word
'       for data entry
' ResumeCreator Macro
' Macro recorded 08/05/2007 by Psyfo
'

Screen.MousePointer = vbHourglass

    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(8), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    Selection.PageSetup.LeftMargin = CentimetersToPoints(1.9)
    Selection.PageSetup.RightMargin = CentimetersToPoints(1.95)
    With Selection.Sections(1)
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleThickThinSmallGap
            .LineWidth = wdLineWidth300pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleThinThickSmallGap
            .LineWidth = wdLineWidth300pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleThickThinSmallGap
            .LineWidth = wdLineWidth300pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleThinThickSmallGap
            .LineWidth = wdLineWidth300pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFrom = wdBorderDistanceFromPageEdge
            .AlwaysInFront = True
            .SurroundHeader = True
            .SurroundFooter = True
            .JoinBorders = False
            .DistanceFromTop = 24
            .DistanceFromLeft = 24
            .DistanceFromBottom = 24
            .DistanceFromRight = 24
            .Shadow = False
            .EnableFirstPageInSection = True
            .EnableOtherPagesInSection = True
            .ApplyPageBordersToAllSections
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleThickThinSmallGap
        .DefaultBorderLineWidth = wdLineWidth300pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.PageSetup.TopMargin = CentimetersToPoints(1.9)
    Selection.PageSetup.BottomMargin = CentimetersToPoints(1.59)


' To Do: Curriculum Vitae Header

Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Name = "Monotype Corsiva"
    With Selection.Font
        .Name = "Monotype Corsiva"
        .Size = 12
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineDouble
        .UnderlineColor = wdColorAutomatic
        .Strikethrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
    End With
    Selection.Font.Size = 28
    Selection.TypeText Text:="C"
    Selection.Font.Size = 18
    Selection.TypeText Text:="urriculum "
    Selection.Font.Size = 28
    Selection.TypeText Text:="V"
    Selection.Font.Size = 18
    Selection.TypeText Text:="itae"
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.TypeParagraph

    Selection.Font.Size = 12
    Selection.Font.Italic = wdToggle
    Selection.Font.Name = "Times New Roman"

    Selection.Font.Italic = wdToggle
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 12
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .Strikethrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
    End With

'================================================================================

' To Do: Personal Details Header

If blnPersonal = True Then
    For a = 0 To nwcPersonal.Count - 1
    With frmDisplay.Text1(a)
        If a <> nwcPersonal.Count And .Text <> "Field Not Selected" Then 'And frmDisplay.Label1(a).Caption = "Postal Address" And frmDisplay.Label1(a).Caption = "Telephone Number" Then
            If frmDisplay.Label1(a).Caption = "Postal Address" Then
                Selection.TypeText Text:=frmDisplay.Label1(a)
                For b = 1 To nwcPersonalAddress.Count
                Selection.TypeText Text:=vbTab & ":" & vbTab & nwcPersonalAddress.Item(b)
                Selection.TypeParagraph
                Next b
                If nwcPersonalAddress.Count < 1 Then
                    Selection.TypeText Text:=vbTab & ":"
                    Selection.TypeParagraph
                End If
                a = a + 1
            End If
        If frmDisplay.Label1(a).Caption = "Telephone Number" Then
            Selection.TypeParagraph
            Selection.TypeText Text:=frmDisplay.Label1(a) & vbTab & ":" & vbTab & frmDisplay.Text1(a)
            Selection.TypeParagraph
            Selection.TypeParagraph
        End If
    End If
    End With
    Next a
End If

If blnPersonal = True Then
    Selection.Font.Name = "Monotype Corsiva"
    Selection.Font.Size = 26
    Selection.TypeText Text:="P"
    Selection.Font.Size = 18
    Selection.TypeText Text:="ersonal "
    Selection.Font.Size = 26
    Selection.TypeText Text:="D"
    Selection.Font.Size = 18
    Selection.TypeText Text:="etails"
    Selection.TypeParagraph
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph
End If

    
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)

'To Do: With the data collected
'       create a resume

If blnPersonal = True Then
    For a = 0 To nwcPersonal.Count - 1
        With frmDisplay.Text1(a)
        If a <> nwcPersonal.Count And .Text <> "Field Not Selected" And frmDisplay.Label1(a).Caption <> "Postal Address" And frmDisplay.Label1(a).Caption <> "Telephone Number" Then
        
        
        Selection.TypeText Text:=frmDisplay.Label1(a) & vbTab & ":" & vbTab & frmDisplay.Text1(a)
        Selection.TypeParagraph
        End If
        End With
    Next a
    Selection.TypeParagraph
End If

'================================================================================


If blnEducational = True Then

    Selection.Font.Name = "Monotype Corsiva"
    Selection.Font.Size = 26
    Selection.TypeText Text:="E"
    Selection.Font.Size = 18
    Selection.TypeText Text:="ducational "
    Selection.Font.Size = 26
    Selection.TypeText Text:="Q"
    Selection.Font.Size = 18
    Selection.TypeText Text:="ualification"
    Application.Run MacroName:="Normal.NewMacros.LineSpacing"
    Selection.TypeParagraph
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend

'To Do: Change Line Spacing to 1
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With

    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

End If



If blnEducational = True Then
    For c = 0 To nwcEducational.Count - 1
        With frmDisplay.Text2(c)
        If c <> nwcEducational.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label2(c).Caption = "Subjects" Then
            Selection.TypeText Text:=frmDisplay.Label2(c)
            For d = 1 To nwcEducationalSubjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcEducationalSubjects.Item(d)
        If d = nwcEducationalSubjects.Count Then
                'To Do: Change Line Spacing to 1.5
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpace1pt5
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
        End If
        Selection.TypeParagraph
        Next d
        If nwcEducationalSubjects.Count < 1 Then
            Selection.TypeText Text:=vbTab & ":"
            Selection.TypeParagraph
        End If
        c = c + 1
        End If
            Selection.TypeText Text:=frmDisplay.Label2(c) & vbTab & ":" & vbTab & frmDisplay.Text2(c)
            Selection.TypeParagraph
        End If
        End If
        End With
    Next c
    Selection.TypeParagraph
End If

'------------------------------------------------------------------------------

If blnTertiary = True Then

    Selection.Font.Name = "Monotype Corsiva"
    Selection.Font.Size = 26
    Selection.TypeText Text:="T"
    Selection.Font.Size = 18
    Selection.TypeText Text:="ertiary "
    Selection.Font.Size = 26
    Selection.TypeText Text:="Q"
    Selection.Font.Size = 18
    Selection.TypeText Text:="ualifications"
    Application.Run MacroName:="Normal.NewMacros.LineSpacing"
    Selection.TypeParagraph
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend

'To Do: Change Line Spacing to 1
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With

    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

End If



If blnTertiary = True Then
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
    For e = 0 To nwcTertiary.Count - 1
        With frmDisplay.Text3(e)
        If e <> nwcTertiary.Count Then
        If .Text <> "Field Not Selected" Then

Select Case frmDisplay.Label3(e).Caption
        Case "Subjects"
            If frmDisplay.Label3(e).Caption = "Subjects" Then
            Selection.TypeText Text:="Subjects"
            For f = 1 To nwcTertiarySubjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcTertiarySubjects.Item(f)
            If f = nwcTertiarySubjects.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next f
            End If
            If nwcTertiarySubjects.Count < 1 Then
                Selection.TypeText Text:=vbTab & ":"
                Selection.TypeParagraph
            End If

        Case "1st Year Subjects Co..."
            If frmDisplay.Label3(e).Caption = "1st Year Subjects Co..." Then
            Selection.TypeText Text:="1st Year Subjects Completed"
            For g = 1 To nwcTertiary1Subjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcTertiary1Subjects.Item(g)
            If g = nwcTertiary1Subjects.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next g
            End If
            If nwcTertiary1Subjects.Count < 1 Then
                Selection.TypeText Text:=vbTab & ":"
                Selection.TypeParagraph
            End If

        Case "2nd Year Subjects Co..."
            If frmDisplay.Label3(e).Caption = "2nd Year Subjects Co..." Then
            Selection.TypeText Text:="2nd Year Subjects Completed"
            For h = 1 To nwcTertiary2Subjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcTertiary2Subjects.Item(h)
            If h = nwcTertiary2Subjects.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next h
            End If
            If nwcTertiary2Subjects.Count < 1 Then
                Selection.TypeText Text:=vbTab & ":"
                Selection.TypeParagraph
            End If

        Case Else
        Selection.TypeText Text:=frmDisplay.Label3(e) & vbTab & ":" & vbTab & frmDisplay.Text3(e)
        Selection.TypeParagraph
End Select

        End If
        End If
        End With
    Next e
    Selection.TypeParagraph
End If

'================================================================================


If blnOther = True Then

    Selection.Font.Name = "Monotype Corsiva"
    Selection.Font.Size = 26
    Selection.TypeText Text:="O"
    Selection.Font.Size = 18
    Selection.TypeText Text:="ther "
    Selection.Font.Size = 26
    Selection.TypeText Text:="Q"
    Selection.Font.Size = 18
    Selection.TypeText Text:="ualifications"
    Application.Run MacroName:="Normal.NewMacros.LineSpacing"
    Selection.TypeParagraph
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend

'To Do: Change Line Spacing to 1
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With

    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

End If



If blnOther = True Then
            'To Do: Change Line Spacing to 1.5
                With Selection.ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpace1pt5
                    .Alignment = wdAlignParagraphLeft
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                End With
    For i = 0 To nwcOther.Count - 1
        With frmDisplay.Text4(i)
        If i <> nwcOther.Count Then
        If .Text <> "Field Not Selected" Then
            Selection.TypeText Text:=frmDisplay.Label4(i) & vbTab & ":" & vbTab & frmDisplay.Text4(i)
            Selection.TypeParagraph
        End If
        End If
        End With
    Next i
    Selection.TypeParagraph
End If

'================================================================================

If blnVoluntary = True Then

    Selection.Font.Name = "Monotype Corsiva"
    Selection.Font.Size = 26
    Selection.TypeText Text:="V"
    Selection.Font.Size = 18
    Selection.TypeText Text:="oluntary "
    Selection.Font.Size = 26
    Selection.TypeText Text:="E"
    Selection.Font.Size = 18
    Selection.TypeText Text:="xperience"
    Application.Run MacroName:="Normal.NewMacros.LineSpacing"
    Selection.TypeParagraph
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend

'To Do: Change Line Spacing to 1
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With

    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

End If




If blnVoluntary = True Then
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
    For j = 0 To nwcVoluntary.Count - 1
        With frmDisplay.Text5(j)
        If j <> nwcVoluntary.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label5(j).Caption = "Address of the Comp..." Then
            Selection.TypeText Text:=frmDisplay.Label5(j)
            For k = 1 To nwcVoluntaryAddress.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcVoluntaryAddress.Item(k)
            If k = nwcVoluntaryAddress.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next k
            j = j + 1
            End If
        Selection.TypeText Text:=frmDisplay.Label5(j) & vbTab & ":" & vbTab & frmDisplay.Text5(j)
        Selection.TypeParagraph
        End If
        End If
        End With
    Next j
    Selection.TypeParagraph
End If

'================================================================================

If blnWork = True Then

    Selection.Font.Name = "Monotype Corsiva"
    Selection.Font.Size = 26
    Selection.TypeText Text:="W"
    Selection.Font.Size = 18
    Selection.TypeText Text:="ork "
    Selection.Font.Size = 26
    Selection.TypeText Text:="E"
    Selection.Font.Size = 18
    Selection.TypeText Text:="xperience"
    Application.Run MacroName:="Normal.NewMacros.LineSpacing"
    Selection.TypeParagraph
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend

'To Do: Change Line Spacing to 1
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With

    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

End If




If blnWork = True Then
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
    For l = 0 To nwcWork.Count - 1
        With frmDisplay.Text6(l)
        If l <> nwcWork.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label6(l).Caption = "Address of the Com..." Then
            Selection.TypeText Text:=frmDisplay.Label6(l)
            For m = 1 To nwcWorkAddress.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcWorkAddress.Item(m)
            If m = nwcWorkAddress.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next m
            l = l + 1
            End If
        Selection.TypeText Text:=frmDisplay.Label6(l) & vbTab & ":" & vbTab & frmDisplay.Text6(l)
        Selection.TypeParagraph
        End If
        End If
        End With
    Next l
    Selection.TypeParagraph
End If

'================================================================================

If blnReference = True Then

    Selection.Font.Name = "Monotype Corsiva"
    Selection.Font.Size = 26
    Selection.TypeText Text:="R"
    Selection.Font.Size = 18
    Selection.TypeText Text:="eference "
    Selection.TypeParagraph
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend

'To Do: Change Line Spacing to 1
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With

    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

End If




If blnReference = True Then
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
    For n = 0 To nwcReference.Count - 1
        With frmDisplay.Text7(n)
        If n <> nwcReference.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label7(n).Caption = "Address" Then
            Selection.TypeText Text:=frmDisplay.Label7(n)
            For o = 1 To nwcReferenceAddress.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcReferenceAddress.Item(o)
            If o = nwcReferenceAddress.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next o
            n = n + 1
            End If
        Selection.TypeText Text:=frmDisplay.Label7(n) & vbTab & ":" & vbTab & frmDisplay.Text7(n)
        Selection.TypeParagraph
        End If
        End If
        End With
    Next n
End If


Screen.MousePointer = vbDefault

MsgBox "New Cv button Pressed!"
End

End Sub

Sub Supa_Design()

    Dim appWord As Word.Application
    Dim blnNewWord As Boolean

    On Error Resume Next
    Set appWord = GetObject(, "Word.Application")
    If Err.Number = 0 Then
        Debug.Print "Word running: Use currently running Word"
        blnNewWord = False
        Documents.Add DocumentType:=wdNewBlankDocument
    Else
        Debug.Print "Word not running: Create instance of Word"
        Set appWord = New Word.Application
        blnNewWord = True
    End If
    On Error GoTo 0
    ' Start of whatever one needs to do with Word
    With appWord
        If .Documents.Count = 0 Then
            .Documents.Add
        End If
        .Visible = True
    End With

'To Do: Start Word and prepare Word
'       for data entry
' ResumeCreator Macro
' Macro recorded 08/05/2007 by Psyfo


Screen.MousePointer = vbHourglass


    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(8), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    Selection.PageSetup.LeftMargin = CentimetersToPoints(1.9)
    Selection.PageSetup.RightMargin = CentimetersToPoints(1.95)

    With Selection.Sections(1)
        With .Borders(wdBorderLeft)
            .ArtStyle = wdArtTwistedLines1
            .ArtWidth = 20
            .ColorIndex = wdAuto
        End With
        With .Borders(wdBorderRight)
            .ArtStyle = wdArtTwistedLines1
            .ArtWidth = 20
            .ColorIndex = wdAuto
        End With
        With .Borders(wdBorderTop)
            .ArtStyle = wdArtTwistedLines1
            .ArtWidth = 20
            .ColorIndex = wdAuto
        End With
        With .Borders(wdBorderBottom)
            .ArtStyle = wdArtTwistedLines1
            .ArtWidth = 20
            .ColorIndex = wdAuto
        End With
        With .Borders
            .DistanceFrom = wdBorderDistanceFromPageEdge
            .AlwaysInFront = True
            .SurroundHeader = True
            .SurroundFooter = True
            .JoinBorders = False
            .DistanceFromTop = 24
            .DistanceFromLeft = 24
            .DistanceFromBottom = 24
            .DistanceFromRight = 24
            .EnableFirstPageInSection = True
            .EnableOtherPagesInSection = True
            .ApplyPageBordersToAllSections
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.PageSetup.TopMargin = CentimetersToPoints(1.9)
    Selection.PageSetup.BottomMargin = CentimetersToPoints(1.59)

'============================

    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Size = 18
    Selection.Font.Name = "Algerian"
    Selection.TypeText Text:="curriculum vitae"
    Selection.TypeParagraph
    Selection.TypeText Text:="of"
    Selection.TypeParagraph
    
    
If blnPersonal = True Then
    For a = 0 To nwcPersonal.Count - 1
        Select Case frmDisplay.Label1(a)
        Case "Name"
        Selection.TypeText Text:=frmDisplay.Text1(a)
        Case "Surname"
        Selection.TypeText Text:=frmDisplay.Text1(a) & " "
        End Select
    Next a
End If

'    Selection.TypeText Text:="sipho stanley mbenenge"
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.MoveUp Unit:=wdLine, Count:=3
    Selection.MoveDown Unit:=wdLine, Count:=3, Extend:=wdExtend
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleThinThickSmallGap
            .LineWidth = wdLineWidth300pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleThickThinSmallGap
            .LineWidth = wdLineWidth300pt
            .Color = wdColorAutomatic
        End With
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleThinThickSmallGap
        .DefaultBorderLineWidth = wdLineWidth300pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.TypeParagraph
'==========================


' Supa_SubHeader Macro
' Macro recorded 07/08/2007 by Psyfo
'
If blnPersonal = True Then

    Selection.Font.Size = 16
    Selection.Font.Name = "Algerian"
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="PERSONAL DETAILS"
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth100pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=2


'To Do: Add Personal Details Data
'       With the data collected
'       create a resume

Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)

    For a = 0 To nwcPersonal.Count - 1
        With frmDisplay.Text1(a)
        If a <> nwcPersonal.Count And .Text <> "Field Not Selected" Then
        
            If frmDisplay.Label1(a).Caption = "Postal Address" Then
                Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
                Selection.TypeText Text:=frmDisplay.Label1(a)
                For b = 1 To nwcPersonalAddress.Count
                Selection.TypeText Text:=vbTab & ":" & vbTab & nwcPersonalAddress.Item(b)
                    If b = nwcPersonalAddress.Count Then
                        'To Do: Change Line Spacing to 1.5
                        Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
                    End If
                    Selection.TypeParagraph
                Next b
                If nwcPersonalAddress.Count < 1 Then
                    Selection.TypeText Text:=vbTab & ":"
                    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
                    Selection.TypeParagraph
                End If
                a = a + 1
            End If
        Selection.TypeText Text:=frmDisplay.Label1(a) & vbTab & ":" & vbTab & frmDisplay.Text1(a)
        Selection.TypeParagraph
        End If
        End With
    Next a
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph
End If
'===========================

'Educational Qualifition

If blnEducational = True Then

    Selection.Font.Size = 16
    Selection.Font.Name = "Algerian"
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="EDUCATIONAL QUALIFICATIONS"
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth100pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=2



    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)

    For c = 0 To nwcEducational.Count - 1
        With frmDisplay.Text2(c)
        If c <> nwcEducational.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label2(c).Caption = "Subjects" Then
            Selection.TypeText Text:=frmDisplay.Label2(c)
            For d = 1 To nwcEducationalSubjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcEducationalSubjects.Item(d)
        If d = nwcEducationalSubjects.Count Then
                'To Do: Change Line Spacing to 1.5
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpace1pt5
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
        End If
        Selection.TypeParagraph
        Next d
        If nwcEducationalSubjects.Count < 1 Then
            Selection.TypeText Text:=vbTab & ":"
            Selection.TypeParagraph
        End If
        c = c + 1
        End If
            Selection.TypeText Text:=frmDisplay.Label2(c) & vbTab & ":" & vbTab & frmDisplay.Text2(c)
            Selection.TypeParagraph
        End If
        End If
        End With
    Next c
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph

End If
'==============================

'Tertiary Qualifications

If blnTertiary = True Then

    Selection.Font.Size = 16
    Selection.Font.Name = "Algerian"
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="TERTIARY QUALIFICATIONS"
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth100pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=2


    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
    For e = 0 To nwcTertiary.Count - 1
        With frmDisplay.Text3(e)
        If e <> nwcTertiary.Count Then
        If .Text <> "Field Not Selected" Then

Select Case frmDisplay.Label3(e).Caption
        Case "Subjects"
            If frmDisplay.Label3(e).Caption = "Subjects" Then
            Selection.TypeText Text:="Subjects"
            For f = 1 To nwcTertiarySubjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcTertiarySubjects.Item(f)
            If f = nwcTertiarySubjects.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next f
            End If
            If nwcTertiarySubjects.Count < 1 Then
                Selection.TypeText Text:=vbTab & ":"
                Selection.TypeParagraph
            End If

        Case "1st Year Subjects Co..."
            If frmDisplay.Label3(e).Caption = "1st Year Subjects Co..." Then
            Selection.TypeText Text:="1st Year Subjects Completed"
            For g = 1 To nwcTertiary1Subjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcTertiary1Subjects.Item(g)
            If g = nwcTertiary1Subjects.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next g
            End If
            If nwcTertiary1Subjects.Count < 1 Then
                Selection.TypeText Text:=vbTab & ":"
                Selection.TypeParagraph
            End If

        Case "2nd Year Subjects Co..."
            If frmDisplay.Label3(e).Caption = "2nd Year Subjects Co..." Then
            Selection.TypeText Text:="2nd Year Subjects Completed"
            For h = 1 To nwcTertiary2Subjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcTertiary2Subjects.Item(h)
            If h = nwcTertiary2Subjects.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next h
            End If
            If nwcTertiary2Subjects.Count < 1 Then
                Selection.TypeText Text:=vbTab & ":"
                Selection.TypeParagraph
            End If

        Case Else
        Selection.TypeText Text:=frmDisplay.Label3(e) & vbTab & ":" & vbTab & frmDisplay.Text3(e)
        Selection.TypeParagraph
End Select

        End If
        End If
        End With
    Next e
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph
End If
'==========================

'Other Qualifications

If blnOther = True Then

    Selection.Font.Size = 16
    Selection.Font.Name = "Algerian"
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="OTHER QUALIFICATIONS"
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth100pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=2


    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)


    For i = 0 To nwcOther.Count - 1
        With frmDisplay.Text4(i)
        If i <> nwcOther.Count Then
        If .Text <> "Field Not Selected" Then
            Selection.TypeText Text:=frmDisplay.Label4(i) & vbTab & ":" & vbTab & frmDisplay.Text4(i)
            Selection.TypeParagraph
        End If
        End If
        End With
    Next i
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph
End If
'==========================

'Voluntary Work

If blnVoluntary = True Then

    Selection.Font.Size = 16
    Selection.Font.Name = "Algerian"
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="VOLUNTARY WORK"
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth100pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=2

    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)

    For j = 0 To nwcVoluntary.Count - 1
        With frmDisplay.Text5(j)
        If j <> nwcVoluntary.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label5(j).Caption = "Address of the Comp..." Then
            Selection.TypeText Text:=frmDisplay.Label5(j)
            For k = 1 To nwcVoluntaryAddress.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcVoluntaryAddress.Item(k)
            If k = nwcVoluntaryAddress.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next k
            j = j + 1
            End If
        Selection.TypeText Text:=frmDisplay.Label5(j) & vbTab & ":" & vbTab & frmDisplay.Text5(j)
        Selection.TypeParagraph
        End If
        End If
        End With
    Next j
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph
End If
'==========================

'Work Experience

If blnWork = True Then

    Selection.Font.Size = 16
    Selection.Font.Name = "Algerian"
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="WORK EXPERIENCE"
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth100pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=2

    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)

    For l = 0 To nwcWork.Count - 1
        With frmDisplay.Text6(l)
        If l <> nwcWork.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label6(l).Caption = "Address of the Com..." Then
            Selection.TypeText Text:=frmDisplay.Label6(l)
            For m = 1 To nwcWorkAddress.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcWorkAddress.Item(m)
            If m = nwcWorkAddress.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next m
            l = l + 1
            End If
        Selection.TypeText Text:=frmDisplay.Label6(l) & vbTab & ":" & vbTab & frmDisplay.Text6(l)
        Selection.TypeParagraph
        End If
        End If
        End With
    Next l
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph
End If
'==========================

'Reference

If blnReference = True Then

    Selection.Font.Size = 16
    Selection.Font.Name = "Algerian"
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="REFERENCE"
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Size = 12
    Selection.Font.Name = "Times New Roman"
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth100pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.MoveDown Unit:=wdLine, Count:=2

    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)

    For n = 0 To nwcReference.Count - 1
        With frmDisplay.Text7(n)
        If n <> nwcReference.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label7(n).Caption = "Address" Then
            Selection.TypeText Text:=frmDisplay.Label7(n)
            For o = 1 To nwcReferenceAddress.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcReferenceAddress.Item(o)
            If o = nwcReferenceAddress.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next o
            n = n + 1
                If nwcReferenceAddress.Count = 0 Then
                Selection.TypeParagraph
                End If
            End If
        Selection.TypeText Text:=frmDisplay.Label7(n) & vbTab & ":" & vbTab & frmDisplay.Text7(n)
        Selection.TypeParagraph
        End If
        End If
        End With
    Next n
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph
End If


End

End Sub

Sub Rafedile_Design()

    Dim appWord As Word.Application
    Dim blnNewWord As Boolean

    On Error Resume Next
    Set appWord = GetObject(, "Word.Application")
    If Err.Number = 0 Then
        Debug.Print "Word running: Use currently running Word"
        blnNewWord = False
        Documents.Add DocumentType:=wdNewBlankDocument
    Else
        Debug.Print "Word not running: Create instance of Word"
        Set appWord = New Word.Application
        blnNewWord = True
    End If
    On Error GoTo 0
    ' Start of whatever one needs to do with Word
    With appWord
        If .Documents.Count = 0 Then
            .Documents.Add
        End If
        .Visible = True
    End With

'To Do: Start Word and prepare Word
'       for data entry
' ResumeCreator Macro
' Macro recorded 08/05/2007 by Psyfo


Screen.MousePointer = vbHourglass


    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(8), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    Selection.PageSetup.LeftMargin = CentimetersToPoints(1.9)
    Selection.PageSetup.RightMargin = CentimetersToPoints(1.95)


    With Selection.Sections(1)
        With .Borders(wdBorderLeft)
            .ArtStyle = wdArtWhiteFlowers
            .ArtWidth = 15
            .ColorIndex = wdAuto
        End With
        With .Borders(wdBorderRight)
            .ArtStyle = wdArtWhiteFlowers
            .ArtWidth = 15
            .ColorIndex = wdAuto
        End With
        With .Borders(wdBorderTop)
            .ArtStyle = wdArtWhiteFlowers
            .ArtWidth = 15
            .ColorIndex = wdAuto
        End With
        With .Borders(wdBorderBottom)
            .ArtStyle = wdArtWhiteFlowers
            .ArtWidth = 15
            .ColorIndex = wdAuto
        End With
        With .Borders
            .DistanceFrom = wdBorderDistanceFromPageEdge
            .AlwaysInFront = True
            .SurroundHeader = True
            .SurroundFooter = True
            .JoinBorders = False
            .DistanceFromTop = 24
            .DistanceFromLeft = 24
            .DistanceFromBottom = 24
            .DistanceFromRight = 24
            .EnableFirstPageInSection = True
            .EnableOtherPagesInSection = True
            .ApplyPageBordersToAllSections
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth100pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.PageSetup.TopMargin = CentimetersToPoints(1.9)
    Selection.PageSetup.BottomMargin = CentimetersToPoints(1.59)
'===========================

    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Name = "Algerian"
    Selection.Font.Size = 18
    With Selection.Font
        .Name = "Algerian"
        .Size = 18
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineDouble
        .UnderlineColor = wdColorAutomatic
        .Strikethrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
    End With
    Selection.TypeText Text:="CURRICULUM VITAE OF "
    
    If blnPersonal = True Then
    For a = 0 To nwcPersonal.Count - 1
        Select Case frmDisplay.Label1(a)
        Case "Name"
        Selection.TypeText Text:=frmDisplay.Text1(a)
        Case "Surname"
        Selection.TypeText Text:=frmDisplay.Text1(a) & " "
        End Select
    Next a
    End If

    Selection.TypeParagraph
    With Selection.Font
        .Name = "Algerian"
        .Size = 18
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .Strikethrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
    End With
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 12
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.TypeParagraph
'===================================
'
' Rafedile_SubHeader Macro
' Macro recorded 14/08/2007 by Psyfo
'

If blnPersonal = True Then

    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Name = "Algerian"
    Selection.Font.Size = 16
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleDouble
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.TypeText Text:="PERSONAL DETAILS"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph
'================================
'To Do: Add Personal Details Data
'       With the data collected
'       create a resume
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)

    For a = 0 To nwcPersonal.Count - 1
        With frmDisplay.Text1(a)
        If a <> nwcPersonal.Count And .Text <> "Field Not Selected" Then
        
            If frmDisplay.Label1(a).Caption = "Postal Address" Then
                Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
                Selection.TypeText Text:=frmDisplay.Label1(a)
                For b = 1 To nwcPersonalAddress.Count
                Selection.TypeText Text:=vbTab & ":" & vbTab & nwcPersonalAddress.Item(b)
                    If b = nwcPersonalAddress.Count Then
                        'To Do: Change Line Spacing to 1.5
                        Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
                    End If
                    Selection.TypeParagraph
                Next b
                If nwcPersonalAddress.Count < 1 Then
                    Selection.TypeText Text:=vbTab & ":"
                    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
                    Selection.TypeParagraph
                End If
                a = a + 1
            End If
        Selection.TypeText Text:=frmDisplay.Label1(a) & vbTab & ":" & vbTab & frmDisplay.Text1(a)
        Selection.TypeParagraph
        End If
        End With
    Next a
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph
End If
'==========================================
'To Do: Add Educational Qualifications Data
'       With the data collected
'       create a resume
If blnEducational = True Then

    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Name = "Algerian"
    Selection.Font.Size = 16
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleDouble
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.TypeText Text:="EDUCATIONAL QUALIFICATIONS"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

'=============================

    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)

    For c = 0 To nwcEducational.Count - 1
        With frmDisplay.Text2(c)
        If c <> nwcEducational.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label2(c).Caption = "Subjects" Then
            Selection.TypeText Text:=frmDisplay.Label2(c)
            For d = 1 To nwcEducationalSubjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcEducationalSubjects.Item(d)
        If d = nwcEducationalSubjects.Count Then
                'To Do: Change Line Spacing to 1.5
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpace1pt5
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
        End If
        Selection.TypeParagraph
        Next d
        If nwcEducationalSubjects.Count < 1 Then
            Selection.TypeText Text:=vbTab & ":"
            Selection.TypeParagraph
        End If
        c = c + 1
        End If
            Selection.TypeText Text:=frmDisplay.Label2(c) & vbTab & ":" & vbTab & frmDisplay.Text2(c)
            Selection.TypeParagraph
        End If
        End If
        End With
    Next c
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph
'==========================================
'To Do: Add Tertiary Qualifications Data
'       With the data collected
'       create a resume
If blnTertiary = True Then

    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Name = "Algerian"
    Selection.Font.Size = 16
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleDouble
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.TypeText Text:="TERTIARY QUALIFICATIONS"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

'=============================

    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
    For e = 0 To nwcTertiary.Count - 1
        With frmDisplay.Text3(e)
        If e <> nwcTertiary.Count Then
        If .Text <> "Field Not Selected" Then

Select Case frmDisplay.Label3(e).Caption
        Case "Subjects"
            If frmDisplay.Label3(e).Caption = "Subjects" Then
            Selection.TypeText Text:="Subjects"
            For f = 1 To nwcTertiarySubjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcTertiarySubjects.Item(f)
            If f = nwcTertiarySubjects.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next f
            End If
            If nwcTertiarySubjects.Count < 1 Then
                Selection.TypeText Text:=vbTab & ":"
                Selection.TypeParagraph
            End If

        Case "1st Year Subjects Co..."
            If frmDisplay.Label3(e).Caption = "1st Year Subjects Co..." Then
            Selection.TypeText Text:="1st Year Subjects Completed"
            For g = 1 To nwcTertiary1Subjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcTertiary1Subjects.Item(g)
            If g = nwcTertiary1Subjects.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next g
            End If
            If nwcTertiary1Subjects.Count < 1 Then
                Selection.TypeText Text:=vbTab & ":"
                Selection.TypeParagraph
            End If

        Case "2nd Year Subjects Co..."
            If frmDisplay.Label3(e).Caption = "2nd Year Subjects Co..." Then
            Selection.TypeText Text:="2nd Year Subjects Completed"
            For h = 1 To nwcTertiary2Subjects.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcTertiary2Subjects.Item(h)
            If h = nwcTertiary2Subjects.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next h
            End If
            If nwcTertiary2Subjects.Count < 1 Then
                Selection.TypeText Text:=vbTab & ":"
                Selection.TypeParagraph
            End If

        Case Else
        Selection.TypeText Text:=frmDisplay.Label3(e) & vbTab & ":" & vbTab & frmDisplay.Text3(e)
        Selection.TypeParagraph
End Select

        End If
        End If
        End With
    Next e
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph
End If

End If
'=======================================
'To Do: Add Other Qualifications Data
'       With the data collected
'       create a resume

If blnOther = True Then

    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Name = "Algerian"
    Selection.Font.Size = 16
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleDouble
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.TypeText Text:="OTHER QUALIFICATIONS"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

'=============================

    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)


    For i = 0 To nwcOther.Count - 1
        With frmDisplay.Text4(i)
        If i <> nwcOther.Count Then
        If .Text <> "Field Not Selected" Then
            Selection.TypeText Text:=frmDisplay.Label4(i) & vbTab & ":" & vbTab & frmDisplay.Text4(i)
            Selection.TypeParagraph
        End If
        End If
        End With
    Next i
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph

End If
'=======================================
'To Do: Add Voluntary Work Data
'       With the data collected
'       create a resume

If blnVoluntary = True Then

    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Name = "Algerian"
    Selection.Font.Size = 16
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleDouble
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.TypeText Text:="VOLUNTARY WORK"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

'=============================

    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)


    For j = 0 To nwcVoluntary.Count - 1
        With frmDisplay.Text5(j)
        If j <> nwcVoluntary.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label5(j).Caption = "Address of the Comp..." Then
            Selection.TypeText Text:=frmDisplay.Label5(j)
            For k = 1 To nwcVoluntaryAddress.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcVoluntaryAddress.Item(k)
            If k = nwcVoluntaryAddress.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next k
            j = j + 1
            End If
        Selection.TypeText Text:=frmDisplay.Label5(j) & vbTab & ":" & vbTab & frmDisplay.Text5(j)
        Selection.TypeParagraph
        End If
        End If
        End With
    Next j
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph

End If
'=======================================
'To Do: Add Work Experience Data
'       With the data collected
'       create a resume

If blnWork = True Then

    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Name = "Algerian"
    Selection.Font.Size = 16
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleDouble
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.TypeText Text:="WORK EXPERIENCE"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

'=============================

    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)


    For l = 0 To nwcWork.Count - 1
        With frmDisplay.Text6(l)
        If l <> nwcWork.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label6(l).Caption = "Address of the Com..." Then
            Selection.TypeText Text:=frmDisplay.Label6(l)
            For m = 1 To nwcWorkAddress.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcWorkAddress.Item(m)
            If m = nwcWorkAddress.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next m
            l = l + 1
            End If
        Selection.TypeText Text:=frmDisplay.Label6(l) & vbTab & ":" & vbTab & frmDisplay.Text6(l)
        Selection.TypeParagraph
        End If
        End If
        End With
    Next l
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph

End If
'=======================================
'To Do: Add Reference Data
'       With the data collected
'       create a resume

If blnReference = True Then

    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Name = "Algerian"
    Selection.Font.Size = 16
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleDouble
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.TypeText Text:="REFERENCE"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeParagraph

'=============================

    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)


    For n = 0 To nwcReference.Count - 1
        With frmDisplay.Text7(n)
        If n <> nwcReference.Count Then
        If .Text <> "Field Not Selected" Then
            If frmDisplay.Label7(n).Caption = "Address" Then
            Selection.TypeText Text:=frmDisplay.Label7(n)
            For o = 1 To nwcReferenceAddress.Count
                'To Do: Change Line Spacing to 1
                    With Selection.ParagraphFormat
                        .LeftIndent = CentimetersToPoints(0)
                        .RightIndent = CentimetersToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .KeepWithNext = False
                        .KeepTogether = False
                        .PageBreakBefore = False
                        .NoLineNumber = False
                        .Hyphenation = True
                        .FirstLineIndent = CentimetersToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    End With
            Selection.TypeText Text:=vbTab & ":" & vbTab & nwcReferenceAddress.Item(o)
            If o = nwcReferenceAddress.Count Then
                    'To Do: Change Line Spacing to 1.5
                        With Selection.ParagraphFormat
                            .LeftIndent = CentimetersToPoints(0)
                            .RightIndent = CentimetersToPoints(0)
                            .SpaceBefore = 0
                            .SpaceBeforeAuto = False
                            .SpaceAfter = 0
                            .SpaceAfterAuto = False
                            .LineSpacingRule = wdLineSpace1pt5
                            .Alignment = wdAlignParagraphLeft
                            .WidowControl = True
                            .KeepWithNext = False
                            .KeepTogether = False
                            .PageBreakBefore = False
                            .NoLineNumber = False
                            .Hyphenation = True
                            .FirstLineIndent = CentimetersToPoints(0)
                            .OutlineLevel = wdOutlineLevelBodyText
                            .CharacterUnitLeftIndent = 0
                            .CharacterUnitRightIndent = 0
                            .CharacterUnitFirstLineIndent = 0
                            .LineUnitBefore = 0
                            .LineUnitAfter = 0
                        End With
            End If
            Selection.TypeParagraph
            Next o
            n = n + 1
                If nwcReferenceAddress.Count = 0 Then
                Selection.TypeParagraph
                End If
            End If
        Selection.TypeText Text:=frmDisplay.Label7(n) & vbTab & ":" & vbTab & frmDisplay.Text7(n)
        Selection.TypeParagraph
        End If
        End If
        End With
    Next n
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.TypeParagraph

End If

End
End Sub
