VERSION 5.00
Begin VB.Form frmCVItems 
   BackColor       =   &H00875B25&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resume Creator - Step 2 of 5"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   12135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Back"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CheckBox chkSelect 
      BackColor       =   &H00875B25&
      Caption         =   "Select &All"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   5640
      Width           =   1815
   End
   Begin VB.ListBox lstCvList 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4380
      ItemData        =   "frmCVItems.frx":0000
      Left            =   3840
      List            =   "frmCVItems.frx":0002
      MultiSelect     =   1  'Simple
      TabIndex        =   9
      Top             =   1125
      Width           =   3255
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0FFFF&
      Default         =   -1  'True
      Height          =   495
      Left            =   7320
      Picture         =   "frmCVItems.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   7320
      Picture         =   "frmCVItems.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ListBox lstSelected 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4380
      ItemData        =   "frmCVItems.frx":0888
      Left            =   8640
      List            =   "frmCVItems.frx":088A
      TabIndex        =   6
      Top             =   1125
      Width           =   3255
   End
   Begin VB.CommandButton cmdDown 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   7320
      Picture         =   "frmCVItems.frx":088C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4965
      Width           =   1095
   End
   Begin VB.CommandButton cmdUp 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   7320
      Picture         =   "frmCVItems.frx":0CCE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4245
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Next"
      Height          =   375
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6225
      Left            =   0
      ScaleHeight     =   6195
      ScaleWidth      =   3570
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   3600
      Begin VB.Label lblPersonal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Details"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   195
         Width           =   1980
      End
      Begin VB.Label lblEducational 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Educational Qualifications"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   1002
         Width           =   3150
      End
      Begin VB.Label lblTertiary 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tertiary Educations"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   14
         Top             =   1809
         Width           =   2340
      End
      Begin VB.Label lblVoluntary 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voluntary Work"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   3423
         Width           =   1845
      End
      Begin VB.Label lblWork 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Work Experience"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   12
         Top             =   4230
         Width           =   2040
      End
      Begin VB.Label lblReference 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   11
         Top             =   5040
         Width           =   1275
      End
      Begin VB.Label lblOther 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Qualifications"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Top             =   2616
         Width           =   2400
      End
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   12120
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   3840
      Picture         =   "frmCVItems.frx":1110
      Top             =   300
      Width           =   480
   End
   Begin VB.Label lblDisplay 
      BackColor       =   &H00875B25&
      BackStyle       =   0  'Transparent
      Caption         =   "Select the fields you require for the Personal Details Data entry"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmCVItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intCvNum, intSelNum As Integer
Dim nwcUpDown As New Collection
Dim nwcCV As New Collection

Private Sub CVFields()

chkSelect.Value = 0

If blnPersonal = True Then
lstCvList.Clear: lstSelected.Clear ': lblPersonal.ForeColor = vbRed
lblDisplay.Caption = "Select the fields you require for the Personal Details Data entry"

With lstCvList
    .AddItem "Surname"
    .AddItem "Name"
    .AddItem "Identity Number"
    .AddItem "Nationality"
    .AddItem "Gender"
    .AddItem "Marital Status"
    .AddItem "Postal Address"
    .AddItem "Telephone Number"
    .AddItem "Cell Number"
    .AddItem "Fax Number"
    .AddItem "E-mail"
    .AddItem "Health"
    .AddItem "Home Language"
    .AddItem "Language Ability"
    .AddItem "Drivers Licence"
    .AddItem "Criminal Record"
    .AddItem "Interest"
    .AddItem "Hobbies"
End With
    Exit Sub
Else
    If nwcPersonal.Count <> 0 Then lblPersonal.ForeColor = vbGreen
End If


If blnEducational = True Then
lstCvList.Clear: lstSelected.Clear ': lblEducational.ForeColor = vbRed
lblDisplay.Caption = "Select the fields you require for the Educational Qualifications Data entry"

With lstCvList
    .AddItem "Name of School"
    .AddItem "Highest Grade Passed"
    .AddItem "Subjects"
    .AddItem "Year"
    .AddItem "Awards / Achievements"
End With
    Exit Sub
Else
    If nwcEducational.Count <> 0 Then lblEducational.ForeColor = vbGreen
End If


If blnTertiary = True Then
lstCvList.Clear: lstSelected.Clear ': lblTertiary.ForeColor = vbRed
lblDisplay.Caption = "Select the fields you require for the Tertiary Qualifications Data entry"

With lstCvList
    .AddItem "Name of Institute"
    .AddItem "Degree"
    .AddItem "Diploma"
    .AddItem "Certificate"
    .AddItem "Course"
    .AddItem "Subjects"
    .AddItem "1st Year Subjects Completed"
    .AddItem "2nd Year Subjects Completed"
    .AddItem "Year"
    .AddItem "Awards / Achievements"
End With
    Exit Sub
Else
    If nwcTertiary.Count <> 0 Then lblTertiary.ForeColor = vbGreen
End If


If blnOther = True Then
lstCvList.Clear: lstSelected.Clear ': lblOther.ForeColor = vbRed
lblDisplay.Caption = "Select the fields you require for the Other Qualification Data entry"

With lstCvList
    .AddItem "Name of Institute"
    .AddItem "Course"
    .AddItem "Training"
    .AddItem "Grade (s) Obtained"
    .AddItem "Certificate"
    .AddItem "Year"
End With
    Exit Sub
Else
    If nwcOther.Count <> 0 Then lblOther.ForeColor = vbGreen
End If


If blnVoluntary = True Then
lstCvList.Clear: lstSelected.Clear ': lblVoluntary.ForeColor = vbRed
lblDisplay.Caption = "Select the fields you require for the Voluntary Work Data entry"

With lstCvList
    .AddItem "Company"
    .AddItem "Address of the Company"
    .AddItem "Voluntered for"
    .AddItem "Position"
    .AddItem "Experience in"
    .AddItem "Period"
    .AddItem "Reason for leaving"
    .AddItem "Title and name of contact"
    .AddItem "Occupation of contact"
    .AddItem "Contact numbers"
    .AddItem "2nd Contact Numbers"
    .AddItem "Fax numbers"
    .AddItem "E-mail"
    .AddItem "Awards / Achievements"
End With
    Exit Sub
Else
    If nwcVoluntary.Count <> 0 Then lblVoluntary.ForeColor = vbGreen
End If


If blnWork = True Then
lstCvList.Clear: lstSelected.Clear ': lblWork.ForeColor = vbRed
lblDisplay.Caption = "Select the fields you require for the Work Experience Data entry"

With lstCvList
    .AddItem "Company"
    .AddItem "Address of the Company"
    .AddItem "Position"
    .AddItem "Duties"
    .AddItem "Period"
    .AddItem "Reason for leaving"
    .AddItem "Title and name of contact"
    .AddItem "Occupation of contact"
    .AddItem "Contact numbers"
    .AddItem "2nd Contact Numbers"
    .AddItem "Fax numbers"
    .AddItem "E-mail"
    .AddItem "Awards / Achievements"
End With
    Exit Sub
Else
    If nwcWork.Count <> 0 Then lblWork.ForeColor = vbGreen
End If


If blnReference = True Then
lstCvList.Clear: lstSelected.Clear ': lblReference.ForeColor = vbRed
lblDisplay.Caption = "Select the fields you require for the Reference Data entry"

With lstCvList
    .AddItem "Title, name and surname"
    .AddItem "Occupation"
    .AddItem "Address"
    .AddItem "Contact numbers"
    .AddItem "2nd Contact Numbers"
    .AddItem "Fax numbers"
    .AddItem "E-mail"
End With
    Exit Sub
Else
    If nwcReference.Count <> 0 Then lblReference.ForeColor = vbGreen
End If

End Sub

Private Sub chkSelect_Click()
If chkSelect.Value = 1 Then
    For a = 0 To lstCvList.ListCount - 1
    lstCvList.Selected(a) = True
    Next a
Else
    For a = 0 To lstCvList.ListCount - 1
    lstCvList.Selected(a) = False
    Next a
End If
End Sub

Private Sub cmdAdd_Click()

'Collect all the data selected
For a = 0 To lstCvList.ListCount - 1
    If lstCvList.Selected(a) = True Then
        nwcCV.Add lstCvList.List(a)
    End If
Next a

If nwcCV.Count = 0 Then
    Exit Sub
End If

'Test how many values selected
If nwcCV.Count < 2 Then
    If lstCvList.ListCount <> 0 And lstCvList.Text <> "" Then
        lstSelected.AddItem lstCvList.Text
        lstCvList.RemoveItem (intCvNum)
    End If
Else
    'If more than one value selected then
    'add all of them using a loop.
    For a = 1 To nwcCV.Count
        lstSelected.AddItem nwcCV.Item(a)
            For b = 0 To lstCvList.ListCount - 1
                If nwcCV.Item(a) = lstCvList.List(b) Then
                    'Remove selected data
                    lstCvList.RemoveItem (b)
                End If
            Next b
    Next a
End If

'Remove the previously selected data
For a = 1 To nwcCV.Count
    nwcCV.Remove (1)
Next a

chkSelect.Value = 0
cmdNext.SetFocus

End Sub

Private Sub cmdBack_Click()
Unload Me
frmCVHeaders.Show
End Sub

Private Sub cmdCancel_Click()
Dim strUserReply As String

strUserReply = MsgBox("Are you sure you want to exit Resume Creator?", vbInformation + vbYesNoCancel, "Quit")
If strUserReply = vbYes Then
    End
End If

End Sub

Private Sub cmdDown_Click()

If lstSelected.ListCount > 1 And lstSelected.Text <> "" And intSelNum <> lstSelected.ListCount - 1 Then
    nwcUpDown.Add lstSelected.ListIndex
    nwcUpDown.Add Val(lstSelected.ListIndex + 1)
    nwcUpDown.Add lstSelected.Text
    nwcUpDown.Add lstSelected.List(intSelNum + 1)
    
    lstSelected.List(nwcUpDown.Item(1)) = nwcUpDown.Item(4)
    lstSelected.List(nwcUpDown.Item(2)) = nwcUpDown.Item(3)
    lstSelected.Selected(intSelNum + 1) = True
    lstSelected.Refresh
    
    'Reset the Values to null
    For a = 1 To nwcUpDown.Count
        nwcUpDown.Remove (1)
    Next a
End If

End Sub

Private Sub cmdNext_Click()

If lstSelected.ListCount = 0 Then
    MsgBox "Warning: No fields where selected." + vbCr + vbCr + "Please select items that you will use to create a CV", vbOKOnly
    Exit Sub
End If


If blnPersonal = True Then
    For a = 0 To lstSelected.ListCount - 1
        nwcPersonal.Add lstSelected.List(a)
    Next a
    blnPersonal = False
    For a = 0 To lstCvList.ListCount - 1
        nwcPerUnselected.Add lstCvList.List(a)
    Next a
    If blnEducational = True Or blnTertiary = True Or blnHobbies = True Or blnOther = True Or blnVoluntary = True Or blnWork = True Or blnReference = True Then
        Call CVFields
        chkSelect.SetFocus
        Exit Sub
    End If
End If

If blnEducational = True Then
    For a = 0 To lstSelected.ListCount - 1
        nwcEducational.Add lstSelected.List(a)
    Next a
    blnEducational = False
    For a = 0 To lstCvList.ListCount - 1
        nwcEduUnselected.Add lstCvList.List(a)
    Next a
    If blnTertiary = True Or blnOther = True Or blnVoluntary = True Or blnWork = True Or blnReference = True Then
        Call CVFields
        chkSelect.SetFocus
        Exit Sub
    End If
End If


If blnTertiary = True Then
    For a = 0 To lstSelected.ListCount - 1
        nwcTertiary.Add lstSelected.List(a)
    Next a
    blnTertiary = False
    For a = 0 To lstCvList.ListCount - 1
        nwcTerUnselected.Add lstCvList.List(a)
    Next a
    If blnHobbies = True Or blnOther = True Or blnVoluntary = True Or blnWork = True Or blnReference = True Then
        Call CVFields
        chkSelect.SetFocus
        Exit Sub
    End If
End If



If blnOther = True Then
    For a = 0 To lstSelected.ListCount - 1
        nwcOther.Add lstSelected.List(a)
    Next a
    blnOther = False
    For a = 0 To lstCvList.ListCount - 1
        nwcOthUnselected.Add lstCvList.List(a)
    Next a
    If blnVoluntary = True Or blnWork = True Or blnReference = True Then
        Call CVFields
        chkSelect.SetFocus
        Exit Sub
    End If
End If


If blnVoluntary = True Then
    For a = 0 To lstSelected.ListCount - 1
        nwcVoluntary.Add lstSelected.List(a)
    Next a
    blnVoluntary = False
    For a = 0 To lstCvList.ListCount - 1
        nwcVolUnselected.Add lstCvList.List(a)
    Next a
    If blnWork = True Or blnReference = True Then
        Call CVFields
        chkSelect.SetFocus
        Exit Sub
    End If
End If


If blnWork = True Then
    For a = 0 To lstSelected.ListCount - 1
        nwcWork.Add lstSelected.List(a)
    Next a
    blnWork = False
    For a = 0 To lstCvList.ListCount - 1
        nwcWorUnselected.Add lstCvList.List(a)
    Next a
    If blnReference = True Then
        Call CVFields
        chkSelect.SetFocus
        Exit Sub
    End If
End If


If blnReference = True Then
    For a = 0 To lstSelected.ListCount - 1
        nwcReference.Add lstSelected.List(a)
    Next a
    blnReference = False
    For a = 0 To lstCvList.ListCount - 1
        nwcRefUnselected.Add lstCvList.List(a)
    Next a
    Call CVFields
    chkSelect.SetFocus
End If


If nwcPersonal.Count <> 0 Then blnPersonal = True
If nwcEducational.Count <> 0 Then blnEducational = True
If nwcTertiary.Count <> 0 Then blnTertiary = True
If nwcOther.Count <> 0 Then blnOther = True
If nwcVoluntary.Count <> 0 Then blnVoluntary = True
If nwcWork.Count <> 0 Then blnWork = True
If nwcReference.Count <> 0 Then blnReference = True


Unload frmCVItems

'Load Cv Creator Data Entry
With frmDisplay
    Load frmDisplay
    .Show
'    .tabFields.Enabled = True
    .tabFields.DeselectAll
    .tabFields.Refresh
End With

frmDisplay.tabFields.Refresh

Personal_Fields
Educational_Fields
Tertiary_Fields
Other_Fields
Voluntary_Fields
Work_Fields
References_Fields

'To Do: Hide other tabs
frmDisplay.picWork.Visible = False
frmDisplay.picVoluntary.Visible = False
frmDisplay.picOther.Visible = False
frmDisplay.picTertiary.Visible = False
frmDisplay.picEducational.Visible = False
frmDisplay.picPersonal.Visible = False
frmDisplay.picReference.Visible = False

'If blnPersonal = True Then Personal_Fields
'If blnEducational = True Then Educational_Fields
'If blnTertiary = True Then Tertiary_Fields
'If blnOther = True Then Other_Fields
'If blnVoluntary = True Then Voluntary_Fields
'If blnWork = True Then Work_Fields
'If blnReference = True Then References_Fields

End Sub

Private Sub cmdRemove_Click()

If lstSelected.ListCount <> 0 And lstSelected.Text <> "" Then
    lstCvList.AddItem lstSelected.Text
    lstSelected.RemoveItem (intSelNum)
End If

chkSelect.Value = 0

End Sub

Private Sub cmdUp_Click()

If lstSelected.ListCount > 1 And lstSelected.Text <> "" And intSelNum > 0 Then
    nwcUpDown.Add lstSelected.ListIndex
    nwcUpDown.Add Val(lstSelected.ListIndex - 1)
    nwcUpDown.Add lstSelected.Text
    nwcUpDown.Add lstSelected.List(intSelNum - 1)
    
    lstSelected.List(nwcUpDown.Item(2)) = nwcUpDown.Item(3)
    lstSelected.List(nwcUpDown.Item(1)) = nwcUpDown.Item(4)
    lstSelected.Selected(intSelNum - 1) = True
    lstSelected.Refresh
    
    'Reset the Values to null
    For a = 1 To nwcUpDown.Count
        nwcUpDown.Remove (1)
    Next a
End If

lstSelected.Refresh

End Sub

Private Sub Form_Load()

If blnPersonal = False Then lblPersonal.ForeColor = vbRed
If blnEducational = False Then lblEducational.ForeColor = vbRed
If blnTertiary = False Then lblTertiary.ForeColor = vbRed
If blnOther = False Then lblOther.ForeColor = vbRed
If blnVoluntary = False Then lblVoluntary.ForeColor = vbRed
If blnWork = False Then lblWork.ForeColor = vbRed
If blnReference = False Then lblReference.ForeColor = vbRed

Call CVFields

End Sub

Private Sub lstCvList_Click()

intCvNum = lstCvList.ListIndex

End Sub

Private Sub lstSelected_Click()

intSelNum = lstSelected.ListIndex

End Sub
