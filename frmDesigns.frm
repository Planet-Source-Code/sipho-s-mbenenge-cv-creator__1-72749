VERSION 5.00
Begin VB.Form frmDesigns 
   BackColor       =   &H00875B25&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resume Creator - Step 4 of 5"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12870
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00875B25&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   8
      ScaleHeight     =   825
      ScaleWidth      =   3600
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4970
      Width           =   3600
      Begin VB.Label lblDesign 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Design Names"
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
         Left            =   345
         TabIndex        =   17
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.PictureBox pctMorolong 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00875B25&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   8
      ScaleHeight     =   825
      ScaleWidth      =   3600
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4140
      Width           =   3600
      Begin VB.Label lblMorolong 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Morolong Design"
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
         Left            =   345
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   240
         Width           =   2040
      End
   End
   Begin VB.PictureBox pctPhungula 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00875B25&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   8
      ScaleHeight     =   825
      ScaleWidth      =   3600
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3320
      Width           =   3600
      Begin VB.Label lblPhungula 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phungula Desing"
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
         Left            =   345
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox pctMakhetha 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00875B25&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   8
      ScaleHeight     =   825
      ScaleWidth      =   3600
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2490
      Width           =   3600
      Begin VB.Label lblMakhetha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Makhetha Design"
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
         Left            =   345
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   240
         Width           =   2115
      End
   End
   Begin VB.PictureBox pctRafedile 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00875B25&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   8
      ScaleHeight     =   825
      ScaleWidth      =   3600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1670
      Width           =   3600
      Begin VB.Label lblRafedile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rafedile Design"
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
         Left            =   345
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.PictureBox pctSupa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00875B25&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   8
      ScaleHeight     =   825
      ScaleWidth      =   3600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   840
      Width           =   3600
      Begin VB.Label lblSupa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supa Design"
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
         Left            =   345
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   1560
      End
   End
   Begin VB.PictureBox pctCVDesigns 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   6120
      ScaleHeight     =   7905
      ScaleWidth      =   6105
      TabIndex        =   5
      Top             =   70
      Width           =   6135
   End
   Begin VB.PictureBox pctMbenenge 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00875B25&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   8
      ScaleHeight     =   825
      ScaleWidth      =   3600
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   20
      Width           =   3600
      Begin VB.Label lblMbenenge 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00875B25&
         BackStyle       =   0  'Transparent
         Caption         =   "Mbenenge Design"
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
         Left            =   345
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.PictureBox pctPersonal 
      Appearance      =   0  'Flat
      BackColor       =   &H00875B25&
      ForeColor       =   &H80000008&
      Height          =   8025
      Left            =   3600
      ScaleHeight     =   7995
      ScaleWidth      =   9210
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9240
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Ok"
      Height          =   375
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8025
      Left            =   0
      ScaleHeight     =   7995
      ScaleWidth      =   3570
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3600
   End
End
Attribute VB_Name = "frmDesigns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdOk_Click()

Unload Me
'If blnDesign = True Then
'    Me.Hide
'Else
'    MsgBox "Please select the designs you want your cv to be formatted like.", vbOKOnly + vbInformation, "Design Type"
'End If
End Sub

Private Sub lblMakhetha_Click()
pctCVDesigns = LoadPicture(App.Path & "\Previews\" & "Makhetha SM.jpg")
strDesignType = "Makhetha"
blnDesign = True
pctMakhetha.BackColor = &H875B25

pctMbenenge.BackColor = vbBlack
pctSupa.BackColor = vbBlack
pctRafedile.BackColor = vbBlack
pctPhungula.BackColor = vbBlack
pctMorolong.BackColor = vbBlack
End Sub

Private Sub lblMbenenge_Click()
pctCVDesigns = LoadPicture(App.Path & "\Previews\" & "Mbenenge SS.jpg")
strDesignType = "Mbenenge"
blnDesign = True
pctMbenenge.BackColor = &H875B25

pctSupa.BackColor = vbBlack
pctRafedile.BackColor = vbBlack
pctMakhetha.BackColor = vbBlack
pctPhungula.BackColor = vbBlack
pctMorolong.BackColor = vbBlack
End Sub

Private Sub lblMorolong_Click()
pctCVDesigns = LoadPicture("")

blnDesign = False
lblMorolong.BackColor = &H875B25

pctMbenenge.BackColor = vbBlack
pctSupa.BackColor = vbBlack
pctRafedile.BackColor = vbBlack
pctMakhetha.BackColor = vbBlack
pctPhungula.BackColor = vbBlack
End Sub

Private Sub lblPhungula_Click()
pctCVDesigns = LoadPicture(App.Path & "\Previews\" & "Sample 1.jpg")
strDesignType = "Phungula"
blnDesign = True
pctPhungula.BackColor = &H875B25

pctMbenenge.BackColor = vbBlack
pctSupa.BackColor = vbBlack
pctRafedile.BackColor = vbBlack
pctMakhetha.BackColor = vbBlack
pctMorolong.BackColor = vbBlack
End Sub

Private Sub lblRafedile_Click()
pctCVDesigns = LoadPicture(App.Path & "\Previews\" & "Rafedile BM.jpg")
strDesignType = "Rafedile"
blnDesign = True
pctRafedile.BackColor = &H875B25

pctMbenenge.BackColor = vbBlack
pctSupa.BackColor = vbBlack
pctMakhetha.BackColor = vbBlack
pctPhungula.BackColor = vbBlack
pctMorolong.BackColor = vbBlack
End Sub

Private Sub lblSupa_Click()
pctCVDesigns = LoadPicture(App.Path & "\Previews\" & "Supa S.jpg")
strDesignType = "Supa"
blnDesign = True
pctSupa.BackColor = &H875B25

pctMbenenge.BackColor = vbBlack
pctRafedile.BackColor = vbBlack
pctMakhetha.BackColor = vbBlack
pctPhungula.BackColor = vbBlack
pctMorolong.BackColor = vbBlack
End Sub
