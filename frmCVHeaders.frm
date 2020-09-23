VERSION 5.00
Begin VB.Form frmCVHeaders 
   BackColor       =   &H00875B25&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resume Creator - Step 1 of 5"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   Icon            =   "frmCVHeaders.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8850
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Ok"
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00875B25&
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   645
      ScaleHeight     =   4755
      ScaleWidth      =   7530
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   7560
      Begin VB.CheckBox chkReference 
         BackColor       =   &H00875B25&
         Height          =   195
         Left            =   600
         TabIndex        =   15
         Top             =   4163
         Width           =   255
      End
      Begin VB.CheckBox chkWork 
         BackColor       =   &H00875B25&
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   3518
         Width           =   255
      End
      Begin VB.CheckBox chkVoluntary 
         BackColor       =   &H00875B25&
         Height          =   195
         Left            =   600
         TabIndex        =   13
         Top             =   2873
         Width           =   255
      End
      Begin VB.CheckBox chkOther 
         BackColor       =   &H00875B25&
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   2213
         Width           =   255
      End
      Begin VB.CheckBox chkTertiary 
         BackColor       =   &H00875B25&
         Height          =   195
         Left            =   600
         TabIndex        =   11
         Top             =   1568
         Width           =   255
      End
      Begin VB.CheckBox chkEducational 
         BackColor       =   &H00875B25&
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   923
         Width           =   255
      End
      Begin VB.CheckBox chkPersonal 
         BackColor       =   &H00875B25&
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   278
         Width           =   255
      End
      Begin VB.Label lblBusR 
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
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   1080
         TabIndex        =   8
         Top             =   195
         Width           =   2265
      End
      Begin VB.Label lblDailyP 
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
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   1080
         TabIndex        =   7
         Top             =   840
         Width           =   3630
      End
      Begin VB.Label lblMulti 
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
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   1080
         TabIndex        =   6
         Top             =   1485
         Width           =   2715
      End
      Begin VB.Label lblRepair 
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
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   1080
         TabIndex        =   5
         Top             =   2790
         Width           =   2145
      End
      Begin VB.Label lblOther 
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
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   1080
         TabIndex        =   4
         Top             =   3435
         Width           =   2445
      End
      Begin VB.Label lblTotalP 
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
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   1080
         TabIndex        =   3
         Top             =   4080
         Width           =   1470
      End
      Begin VB.Label lblTotalC 
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
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   1080
         TabIndex        =   2
         Top             =   2130
         Width           =   2760
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   360
      Picture         =   "frmCVHeaders.frx":0442
      Top             =   330
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the fields that you want your CV / RESUME to contain and click the Next button to continue."
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
      Height          =   660
      Left            =   960
      TabIndex        =   16
      Top             =   240
      Width           =   7560
   End
End
Attribute VB_Name = "frmCVHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
Dim intField As Integer

'--------------------------------------------------------------------------------------
For a = 1 To nwcCVFields.Count
nwcCVFields.Remove (1)
Next

If chkPersonal.Value = 1 Then
    nwcCVFields.Add "Personal Details"
    blnPersonal = True
Else
    blnPersonal = False
End If

If chkEducational.Value = 1 Then
    nwcCVFields.Add "Educational Qualifications"
    blnEducational = True
Else
    blnEducational = False
End If

If chkTertiary.Value = 1 Then
    nwcCVFields.Add "Tertiary Education"
    blnTertiary = True
Else
    blnTertiary = False
End If

If chkOther.Value = 1 Then
    nwcCVFields.Add "Other Qualifications"
    blnOther = True
Else
    blnOther = False
End If

If chkVoluntary.Value = 1 Then
    nwcCVFields.Add "Voluntary Work"
    blnVoluntary = True
Else
    blnVoluntary = False
End If

If chkWork.Value = 1 Then
    nwcCVFields.Add "Work Experience"
    blnWork = True
Else
    blnWork = False
End If

If chkReference.Value = 1 Then
    nwcCVFields.Add "Reference"
    blnReference = True
Else
    blnReference = False
End If


If nwcCVFields.Count = 0 Then
    MsgBox "Please select fields needed for data entry"
    Exit Sub
End If
'--------------------------------------------------------------------------------------


Unload Me

Load frmCVItems
frmCVItems.Show

End Sub
