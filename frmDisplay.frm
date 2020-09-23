VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDisplay 
   BackColor       =   &H00875B25&
   Caption         =   "Cv Creator - Fields"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   14370
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picFields 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00875B25&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   13800
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   0
      Width           =   13800
      Begin MSComctlLib.TabStrip tabFields 
         Height          =   400
         Left            =   0
         TabIndex        =   47
         Top             =   960
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   714
         TabWidthStyle   =   1
         MultiRow        =   -1  'True
         Style           =   2
         Separators      =   -1  'True
         TabMinWidth     =   2646
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   7
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Personal Details"
               Key             =   "Personal"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Educational Qual..."
               Key             =   "Educational"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Tertiary Qual..."
               Key             =   "Tertiary"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Other Qual..."
               Key             =   "Other"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Voluntary Work"
               Key             =   "Voluntary"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Work Experience"
               Key             =   "Work"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "References"
               Key             =   "References"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Baskerville Old Face"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTip 
         BackStyle       =   0  'Transparent
         Caption         =   "Please select the fields that you want your CV / RESUME to contain and click the Next button to continue."
         Height          =   495
         Left            =   1440
         TabIndex        =   161
         Top             =   240
         Width           =   7335
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   360
         Picture         =   "frmDisplay.frx":0000
         Top             =   210
         Width           =   480
      End
   End
   Begin VB.PictureBox picEducational 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00875B25&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8025
      Left            =   0
      ScaleHeight     =   8025
      ScaleWidth      =   13800
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   13800
      Begin VB.CommandButton cmdEduSub 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Subjects"
         Height          =   260
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2810
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   0
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Field Not Selected"
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   1
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Field Not Selected"
         Top             =   1500
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   2
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Field Not Selected"
         Top             =   2760
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   3
         Left            =   3045
         TabIndex        =   25
         Text            =   "Field Not Selected"
         Top             =   4020
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   4
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Field Not Selected"
         Top             =   5280
         Width           =   4575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of School"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   285
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Highest Grade Passed"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   1545
         Width           =   2475
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subjects"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   2805
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   24
         Top             =   4065
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Awards / Achievements"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   26
         Top             =   5325
         Width           =   2670
      End
   End
   Begin VB.PictureBox picPersonal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00875B25&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8025
      Left            =   0
      ScaleHeight     =   8025
      ScaleWidth      =   13800
      TabIndex        =   162
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   13800
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   16
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "Field Not Selected"
         Top             =   6992
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   15
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Field Not Selected"
         Top             =   6560
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   14
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Field Not Selected"
         Top             =   6128
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   13
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Field Not Selected"
         Top             =   5696
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   12
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Field Not Selected"
         Top             =   5264
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   11
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "Field Not Selected"
         Top             =   4832
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   10
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Field Not Selected"
         Top             =   4400
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   9
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Field Not Selected"
         Top             =   3968
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   8
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Field Not Selected"
         Top             =   3536
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   7
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Field Not Selected"
         Top             =   3104
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   6
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Field Not Selected"
         Top             =   2672
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   5
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Field Not Selected"
         Top             =   2240
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   4
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Field Not Selected"
         Top             =   1808
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   3
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Field Not Selected"
         Top             =   1376
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   2
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Field Not Selected"
         Top             =   944
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   1
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Field Not Selected"
         Top             =   512
         Width           =   4575
      End
      Begin VB.CommandButton cmdAddPAddress 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Address"
         Height          =   260
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   2722
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   17
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Field Not Selected"
         Top             =   7440
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   320
         Index           =   0
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Field Not Selected"
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Interest"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   16
         Left            =   240
         TabIndex        =   42
         Top             =   7037
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Criminal Offence"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   15
         Left            =   240
         TabIndex        =   40
         Top             =   6605
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Drivers Licence"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   14
         Left            =   240
         TabIndex        =   38
         Top             =   6173
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Language Ability"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   13
         Left            =   240
         TabIndex        =   36
         Top             =   5741
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Home Language"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   12
         Left            =   240
         TabIndex        =   34
         Top             =   5309
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Health"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   11
         Left            =   240
         TabIndex        =   32
         Top             =   4877
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   10
         Left            =   240
         TabIndex        =   30
         Top             =   4445
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax Number"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   9
         Left            =   240
         TabIndex        =   28
         Top             =   4013
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Number"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   8
         Left            =   240
         TabIndex        =   16
         Top             =   3581
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone Number"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   7
         Left            =   240
         TabIndex        =   14
         Top             =   3149
         Width           =   2205
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Postal"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   6
         Left            =   240
         TabIndex        =   12
         Top             =   2717
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Marital Status"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   2285
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1853
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1421
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Identity Number"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   989
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   145
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hobbies"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   17
         Left            =   240
         TabIndex        =   44
         Top             =   7485
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   557
         Width           =   675
      End
   End
   Begin VB.PictureBox picReference 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00875B25&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8025
      Left            =   0
      ScaleHeight     =   8025
      ScaleWidth      =   13800
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   13800
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   6
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   140
         Text            =   "Field Not Selected"
         Top             =   5280
         Width           =   4575
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   5
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   139
         Text            =   "Field Not Selected"
         Top             =   4440
         Width           =   4575
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   4
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   138
         Text            =   "Field Not Selected"
         Top             =   3600
         Width           =   4575
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   3
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   137
         Text            =   "Field Not Selected"
         Top             =   2760
         Width           =   4575
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   2
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   136
         Text            =   "Field Not Selected"
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   1
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   135
         Text            =   "Field Not Selected"
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   0
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   134
         Text            =   "Field Not Selected"
         Top             =   240
         Width           =   4575
      End
      Begin VB.CommandButton cmdRefAddress 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Address"
         Height          =   260
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   1970
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   6
         Left            =   240
         TabIndex        =   147
         Top             =   5325
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax numbers"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   146
         Top             =   4485
         Width           =   1500
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Contact Numbers"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   145
         Top             =   3645
         Width           =   2505
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact numbers"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   144
         Top             =   2805
         Width           =   1965
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   143
         Top             =   1965
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   142
         Top             =   1125
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title, name and surname"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   141
         Top             =   285
         Width           =   2850
      End
   End
   Begin VB.PictureBox picWork 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00875B25&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8025
      Left            =   0
      ScaleHeight     =   8025
      ScaleWidth      =   13800
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   13800
      Begin VB.CommandButton cmdWorkAddress 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Address"
         Height          =   260
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   890
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   0
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   117
         Text            =   "Field Not Selected"
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   1
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   116
         Text            =   "Field Not Selected"
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   2
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "Field Not Selected"
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   3
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "Field Not Selected"
         Top             =   2040
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   4
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Field Not Selected"
         Top             =   2640
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   5
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "Field Not Selected"
         Top             =   3240
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   6
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   111
         Text            =   "Field Not Selected"
         Top             =   3840
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   7
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   110
         Text            =   "Field Not Selected"
         Top             =   4440
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   8
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   109
         Text            =   "Field Not Selected"
         Top             =   5040
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   9
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   108
         Text            =   "Field Not Selected"
         Top             =   5640
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   10
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   107
         Text            =   "Field Not Selected"
         Top             =   6240
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   11
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   106
         Text            =   "Field Not Selected"
         Top             =   6840
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   12
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "Field Not Selected"
         Top             =   7440
         Width           =   4575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   131
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address of the Company"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   130
         Top             =   885
         Width           =   2775
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   129
         Top             =   1485
         Width           =   930
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Duties"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   128
         Top             =   2085
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   127
         Top             =   2685
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason for leaving"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   126
         Top             =   3285
         Width           =   2100
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title and name of contact"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   6
         Left            =   240
         TabIndex        =   125
         Top             =   3885
         Width           =   2895
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation of contact"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   7
         Left            =   240
         TabIndex        =   124
         Top             =   4485
         Width           =   2490
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact numbers"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   8
         Left            =   240
         TabIndex        =   123
         Top             =   5085
         Width           =   1965
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Contact Numbers"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   9
         Left            =   240
         TabIndex        =   122
         Top             =   5685
         Width           =   2505
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax numbers"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   10
         Left            =   240
         TabIndex        =   121
         Top             =   6285
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   11
         Left            =   240
         TabIndex        =   120
         Top             =   6885
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Awards / Achievements"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   12
         Left            =   240
         TabIndex        =   119
         Top             =   7485
         Width           =   2670
      End
   End
   Begin VB.PictureBox picVoluntary 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00875B25&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8025
      Left            =   0
      ScaleHeight     =   8025
      ScaleWidth      =   13800
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   13800
      Begin VB.CommandButton cmdVolAddress 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Address"
         Height          =   260
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   825
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   0
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "Field Not Selected"
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   1
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "Field Not Selected"
         Top             =   775
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   2
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "Field Not Selected"
         Top             =   1310
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   3
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "Field Not Selected"
         Top             =   1845
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   4
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   84
         Text            =   "Field Not Selected"
         Top             =   2380
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   5
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "Field Not Selected"
         Top             =   2915
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   6
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Field Not Selected"
         Top             =   3450
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   7
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "Field Not Selected"
         Top             =   3985
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   8
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Field Not Selected"
         Top             =   4520
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   9
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "Field Not Selected"
         Top             =   5055
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   10
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "Field Not Selected"
         Top             =   5590
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   11
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Field Not Selected"
         Top             =   6125
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   12
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Field Not Selected"
         Top             =   6660
         Width           =   4575
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   13
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Field Not Selected"
         Top             =   7200
         Width           =   4575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   103
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address of the Company"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   102
         Top             =   820
         Width           =   2775
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voluntered for"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   101
         Top             =   1355
         Width           =   1680
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   100
         Top             =   1890
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Experience in"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   99
         Top             =   2425
         Width           =   1560
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   98
         Top             =   2960
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason for leaving"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   6
         Left            =   240
         TabIndex        =   97
         Top             =   3495
         Width           =   2100
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title and name of contact"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   7
         Left            =   240
         TabIndex        =   96
         Top             =   4030
         Width           =   2895
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation of contact"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   8
         Left            =   240
         TabIndex        =   95
         Top             =   4565
         Width           =   2490
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact numbers"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   9
         Left            =   240
         TabIndex        =   94
         Top             =   5100
         Width           =   1965
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Contact Numbers"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   10
         Left            =   240
         TabIndex        =   93
         Top             =   5635
         Width           =   2505
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax numbers"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   11
         Left            =   240
         TabIndex        =   92
         Top             =   6170
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   12
         Left            =   240
         TabIndex        =   91
         Top             =   6705
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Awards / Achievements"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   13
         Left            =   240
         TabIndex        =   90
         Top             =   7245
         Width           =   2670
      End
   End
   Begin VB.PictureBox picOther 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00875B25&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8025
      Left            =   0
      ScaleHeight     =   8025
      ScaleWidth      =   13680
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   13680
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   0
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   154
         Text            =   "Field Not Selected"
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   1
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   153
         Text            =   "Field Not Selected"
         Top             =   1026
         Width           =   4575
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   2
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   152
         Text            =   "Field Not Selected"
         Top             =   1812
         Width           =   4575
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   3
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   151
         Text            =   "Field Not Selected"
         Top             =   2598
         Width           =   4575
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   4
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   150
         Text            =   "Field Not Selected"
         Top             =   3384
         Width           =   4575
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   5
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   149
         Text            =   "Field Not Selected"
         Top             =   4170
         Width           =   4575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Institute"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   160
         Top             =   285
         Width           =   2010
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   159
         Top             =   1065
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Training"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   158
         Top             =   1860
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grade (s) Obtained"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   157
         Top             =   2640
         Width           =   2190
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Certificate"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   156
         Top             =   3435
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   155
         Top             =   4215
         Width           =   555
      End
   End
   Begin VB.PictureBox picTertiary 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00875B25&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8025
      Left            =   0
      ScaleHeight     =   8025
      ScaleWidth      =   13800
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   13800
      Begin VB.CommandButton cmdTerSub 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Subjects"
         Height          =   260
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   1836
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   0
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Field Not Selected"
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   1
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Field Not Selected"
         Top             =   1013
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   2
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "Field Not Selected"
         Top             =   1786
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   3
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "Field Not Selected"
         Top             =   2559
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   4
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Field Not Selected"
         Top             =   3332
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   5
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Field Not Selected"
         Top             =   4105
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   6
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "Field Not Selected"
         Top             =   4878
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   7
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Field Not Selected"
         Top             =   5651
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   8
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "Field Not Selected"
         Top             =   6424
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   9
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Field Not Selected"
         Top             =   7200
         Width           =   4575
      End
      Begin VB.CommandButton cmdTer1Sub 
         BackColor       =   &H00C0FFFF&
         Caption         =   "1st Year Sub"
         Height          =   260
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4928
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdTer2Sub 
         BackColor       =   &H00C0FFFF&
         Caption         =   "2nd Year Sub"
         Height          =   260
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   5701
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Institute"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   73
         Top             =   285
         Width           =   2010
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Degree"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   72
         Top             =   1058
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diploma"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   71
         Top             =   1831
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Certificate"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   70
         Top             =   2604
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   69
         Top             =   3377
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subjects"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   68
         Top             =   4150
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1st Year Subjects Co..."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   6
         Left            =   240
         TabIndex        =   67
         Top             =   4920
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Year Subjects Co..."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   7
         Left            =   240
         TabIndex        =   66
         Top             =   5700
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   8
         Left            =   240
         TabIndex        =   65
         Top             =   6469
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Awards / Achievements"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   9
         Left            =   240
         TabIndex        =   64
         Top             =   7245
         Width           =   2670
      End
   End
   Begin VB.Menu mnuPAddress 
      Caption         =   "PAddress"
      Visible         =   0   'False
      Begin VB.Menu mnuAddPadd 
         Caption         =   "Add Address"
      End
      Begin VB.Menu Pspacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPadd 
         Caption         =   "View Address"
      End
      Begin VB.Menu Pspacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearPadd 
         Caption         =   "Clear Address"
      End
   End
   Begin VB.Menu mnuESubject 
      Caption         =   "ESubject"
      Visible         =   0   'False
      Begin VB.Menu mnuAddEsub 
         Caption         =   "Add Subject"
      End
      Begin VB.Menu Espacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewEsub 
         Caption         =   "View Subject"
      End
      Begin VB.Menu Espacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveEsub 
         Caption         =   "Remove Subject"
      End
      Begin VB.Menu mnuRemoveAllESub 
         Caption         =   "Remove All Subject (s)"
      End
   End
   Begin VB.Menu mnuTSubject 
      Caption         =   "TSubject"
      Visible         =   0   'False
      Begin VB.Menu mnuAddTsub 
         Caption         =   "Add Tertiary Subject"
      End
      Begin VB.Menu Tspacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTsub 
         Caption         =   "View Tertiary Subject"
      End
      Begin VB.Menu Tspacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteTsub 
         Caption         =   "Delete Tertiary Subject"
      End
      Begin VB.Menu mnuDeleteAllTsub 
         Caption         =   "Delete All Tertiary Subject (s)"
      End
   End
   Begin VB.Menu mnu1stTSubject 
      Caption         =   "1stTSubject"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd1Tsub 
         Caption         =   "Add 1st Year Subject"
      End
      Begin VB.Menu Tspacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView1Tsub 
         Caption         =   "View 1st Year Subject"
      End
      Begin VB.Menu Tspacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete1Tsub 
         Caption         =   "Delete 1st Year Subject"
      End
      Begin VB.Menu mnuDeleteAll1Tsub 
         Caption         =   "Delete All 1st Year Subject (s)"
      End
   End
   Begin VB.Menu mnu2ndTSubject 
      Caption         =   "2ndTSubject"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd2Tsub 
         Caption         =   "Add 2nd Year Subject"
      End
      Begin VB.Menu Tspacer5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView2Tsub 
         Caption         =   "View 2nd Year Subject"
      End
      Begin VB.Menu Tspacer6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete2Tsub 
         Caption         =   "Delete 2nd Year Subject"
      End
      Begin VB.Menu mnuDeleteAll2Tsub 
         Caption         =   "Delete All 2nd Year Subject"
      End
   End
   Begin VB.Menu mnuVCaddress 
      Caption         =   "VCaddress"
      Visible         =   0   'False
      Begin VB.Menu mnuAddVCaddress 
         Caption         =   "Add Company Address"
      End
      Begin VB.Menu Vspacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewVCaddress 
         Caption         =   "View Company Address"
      End
      Begin VB.Menu Vspacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearVCaddress 
         Caption         =   "Clear Company Address"
      End
   End
   Begin VB.Menu mnuWCaddress 
      Caption         =   "WCaddress"
      Visible         =   0   'False
      Begin VB.Menu mnuAddWCaddress 
         Caption         =   "Add Company Address"
      End
      Begin VB.Menu Wspacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewWCaddress 
         Caption         =   "View Company Address"
      End
      Begin VB.Menu Wspacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearWCaddress 
         Caption         =   "Clear Company Address"
      End
   End
   Begin VB.Menu mnuRAddress 
      Caption         =   "RAddress"
      Visible         =   0   'False
      Begin VB.Menu mnuAddRaddress 
         Caption         =   "Add Address"
      End
      Begin VB.Menu Rspacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRaddress 
         Caption         =   "View Address"
      End
      Begin VB.Menu Rspacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearRaddress 
         Caption         =   "Clear Address"
      End
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddPAddress_Click()
Call PopupMenu(mnuPAddress)
End Sub

Private Sub cmdEduSub_Click()
Call PopupMenu(mnuESubject)
End Sub

Private Sub cmdRefAddress_Click()
Call PopupMenu(mnuRAddress)
End Sub

Private Sub cmdTer1Sub_Click()
Call PopupMenu(mnu1stTSubject)
End Sub

Private Sub cmdTer2Sub_Click()
Call PopupMenu(mnu2ndTSubject)
End Sub

Private Sub cmdTerSub_Click()
Call PopupMenu(mnuTSubject)
End Sub

Private Sub cmdVolAddress_Click()
Call PopupMenu(mnuVCaddress)
End Sub

Private Sub cmdWorkAddress_Click()
Call PopupMenu(mnuWCaddress)
End Sub

Private Sub Form_Initialize()
tabFields.DeselectAll
tabFields.Refresh
tabFields.DeselectAll

    With frmMain
        .mnuFile.Visible = True
        .mnuView.Visible = True
        .mnuHelp.Visible = True
    End With

End Sub

Private Sub Form_Load()
    frmDisplay.tabFields.DeselectAll
    tabFields.DeselectAll
    frmDisplay.tabFields.Refresh
End Sub

Private Sub Form_Resize()
    frmDisplay.picFields.Width = frmDisplay.Width
    tabFields.Width = frmDisplay.Width
'    tabFields.DeselectAll
'    picPersonal.Width = frmDisplay.Width
End Sub

Private Sub mnuAdd1Tsub_Click()

For a = 0 To nwcTertiary.Count
With Text3(a)
 '   MsgBox nwcEducational.Count & " - " & a
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "1st Year Subjects Co..." Then
    nwcTertiary1Subjects.Add .Text
    .Text = ""
    .SetFocus
'    Exit Sub
    End If
End With
Next a

End Sub

Private Sub mnuAdd2Tsub_Click()
For a = 0 To nwcTertiary.Count
With Text3(a)
 '   MsgBox nwcEducational.Count & " - " & a
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "2nd Year Subjects Co..." Then
    nwcTertiary2Subjects.Add .Text
    .Text = ""
    .SetFocus
'    Exit Sub
    End If
End With
Next a

End Sub

Private Sub mnuAddEsub_Click()

For a = 0 To nwcEducational.Count
With Text2(a)
 '   MsgBox nwcEducational.Count & " - " & a
    If a >= nwcEducational.Count Then Exit Sub
    If Label2(a).Caption = "Subjects" Then
    nwcEducationalSubjects.Add .Text
    .Text = ""
    .SetFocus
'    Exit Sub
    End If
End With
Next a

End Sub

Private Sub mnuAddPadd_Click()

For a = 0 To nwcPersonal.Count
    With Text1(a)
    If a >= nwcPersonal.Count Then Exit Sub
    If Label1(a).Caption = "Postal Address" Then
    nwcPersonalAddress.Add .Text
    .Text = ""
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuAddRaddress_Click()

For a = 0 To nwcReference.Count
    With Text7(a)
    If a >= nwcReference.Count Then Exit Sub
    If Label7(a).Caption = "Address" Then
    nwcReferenceAddress.Add .Text
    .Text = ""
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuAddTsub_Click()

For a = 0 To nwcTertiary.Count
With Text3(a)
 '   MsgBox nwcEducational.Count & " - " & a
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "Subjects" Then
    nwcTertiarySubjects.Add .Text
    .Text = ""
    .SetFocus
'    Exit Sub
    End If
End With
Next a

End Sub

Private Sub mnuAddVCaddress_Click()

For a = 0 To nwcVoluntary.Count
    With Text5(a)
    If a >= nwcVoluntary.Count Then Exit Sub
    If Label5(a).Caption = "Address of the Comp..." Then
    nwcVoluntaryAddress.Add .Text
    .Text = ""
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuAddWCaddress_Click()

For a = 0 To nwcWork.Count
    With Text6(a)
    If a >= nwcWork.Count Then Exit Sub
    If Label6(a).Caption = "Address of the Com..." Then
    nwcWorkAddress.Add .Text
    .Text = ""
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuClearPadd_Click()

Do While nwcPersonalAddress.Count <> 0
    nwcPersonalAddress.Remove (1)
Loop

For a = 0 To nwcPersonal.Count
    With Text1(a)
    If a >= nwcPersonal.Count Then Exit Sub
    If Label1(a).Caption = "Postal Address" Then
    .Text = "Please enter your Postal Address here!"
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuClearRaddress_Click()

Do While nwcReferenceAddress.Count <> 0
    nwcReferenceAddress.Remove (1)
Loop

For a = 0 To nwcReference.Count
    With Text7(a)
    If a >= nwcReference.Count Then Exit Sub
    If Label7(a).Caption = "Address" Then
    .Text = "Type the Address of the Company here!"
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuClearVCaddress_Click()

Do While nwcVoluntaryAddress.Count <> 0
    nwcVoluntaryAddress.Remove (1)
Loop

For a = 0 To nwcVoluntary.Count
    With Text5(a)
    If a >= nwcVoluntary.Count Then Exit Sub
    If Label5(a).Caption = "Address of the Comp..." Then
    .Text = "Type the Address of the Company here!"
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuClearWCaddress_Click()

Do While nwcWorkAddress.Count <> 0
    nwcWorkAddress.Remove (1)
Loop

For a = 0 To nwcWork.Count
    With Text6(a)
    If a >= nwcWork.Count Then Exit Sub
    If Label6(a).Caption = "Address of the Com..." Then
    .Text = "Type the Address of the Company here!"
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuDelete1Tsub_Click()

On Error GoTo RemoveErrorHandler

Dim intSubNum As Integer


intSubNum = InputBox("Please enter the number of the subject you want to remove", "Educational Qualifications - Subjects", 1)
If intSubNum > 0 Then
    nwcTertiary1Subjects.Remove (intSubNum)
End If

For a = 0 To nwcTertiary.Count
    With Text3(a)
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "1st Year Subjects Co..." Then
    .SetFocus
    End If
End With
Next a

Exit Sub



RemoveErrorHandler:

MsgBox Err.Description

For a = 0 To nwcTertiary.Count
    With Text3(a)
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "1st Year Subjects Co..." Then
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuDelete2Tsub_Click()

On Error GoTo RemoveErrorHandler

Dim intSubNum As Integer


intSubNum = InputBox("Please enter the number of the subject you want to remove", "Educational Qualifications - Subjects", 1)
If intSubNum > 0 Then
    nwcTertiary2Subjects.Remove (intSubNum)
End If

For a = 0 To nwcTertiary.Count
    With Text3(a)
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "2nd Year Subjects Co..." Then
    .SetFocus
    End If
End With
Next a

Exit Sub



RemoveErrorHandler:

MsgBox Err.Description

For a = 0 To nwcTertiary.Count
    With Text3(a)
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "2nd Year Subjects Co..." Then
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuDeleteAll1Tsub_Click()

Do While nwcTertiary1Subjects.Count <> 0
    nwcTertiary1Subjects.Remove (1)
Loop

For a = 0 To nwcTertiary.Count
    With Text3(a)
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "1st Year Subjects Co..." Then
    .Text = "Type your 1st Year Subjects Completed here!"
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuDeleteAll2Tsub_Click()

Do While nwcTertiary2Subjects.Count <> 0
    nwcTertiary2Subjects.Remove (1)
Loop

For a = 0 To nwcTertiary.Count
    With Text3(a)
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "2nd Year Subjects Co..." Then
    .Text = "Type your 2nd Year Subjects Completed here!"
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuDeleteAllTsub_Click()

Do While nwcTertiarySubjects.Count <> 0
    nwcTertiarySubjects.Remove (1)
Loop

For a = 0 To nwcTertiary.Count
    With Text3(a)
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "Subjects" Then
    .Text = "Please enter your Subjects here!"
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuDeleteTsub_Click()

On Error GoTo RemoveErrorHandler

Dim intSubNum As Integer


intSubNum = InputBox("Please enter the number of the subject you want to remove", "Educational Qualifications - Subjects", 1)
If intSubNum > 0 Then
    nwcTertiarySubjects.Remove (intSubNum)
End If

For a = 0 To nwcTertiary.Count
    With Text3(a)
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "Subjects" Then
    .SetFocus
    End If
End With
Next a

Exit Sub



RemoveErrorHandler:

MsgBox Err.Description

For a = 0 To nwcTertiary.Count
    With Text3(a)
    If a >= nwcTertiary.Count Then Exit Sub
    If Label3(a).Caption = "Subjects" Then
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuRemoveAllESub_Click()

Do While nwcEducationalSubjects.Count <> 0
    nwcEducationalSubjects.Remove (1)
Loop

For a = 0 To nwcEducational.Count
    With Text2(a)
    If a >= nwcEducational.Count Then Exit Sub
    If Label2(a).Caption = "Subjects" Then
    .Text = "Please enter your Subjects here!"
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuRemoveEsub_Click()

On Error GoTo RemoveErrorHandler

Dim intSubNum As Integer


intSubNum = InputBox("Please enter the number of the subject you want to remove", "Educational Qualifications - Subjects", 1)
If intSubNum > 0 Then
    nwcEducationalSubjects.Remove (intSubNum)
End If

For a = 0 To nwcEducational.Count
    With Text2(a)
    If a >= nwcEducational.Count Then Exit Sub
    If Label2(a).Caption = "Subjects" Then
    .SetFocus
    End If
End With
Next a

Exit Sub



RemoveErrorHandler:

MsgBox Err.Description

For a = 0 To nwcEducational.Count
    With Text2(a)
    If a >= nwcEducational.Count Then Exit Sub
    If Label2(a).Caption = "Subjects" Then
    .SetFocus
    End If
End With
Next a

End Sub

Private Sub mnuView1Tsub_Click()

Dim strSubjects As String

For a = 1 To nwcTertiary1Subjects.Count
    strSubjects = strSubjects & a & "  -  " & nwcTertiary1Subjects.Item(a) & vbCrLf
Next a

MsgBox strSubjects, vbInformation + vbOKOnly, "Tertiary Qualifications - 1st Year Subjects"

End Sub

Private Sub mnuView2Tsub_Click()

Dim strSubjects As String

For a = 1 To nwcTertiary2Subjects.Count
    strSubjects = strSubjects & a & "  -  " & nwcTertiary2Subjects.Item(a) & vbCrLf
Next a

MsgBox strSubjects, vbInformation + vbOKOnly, "Tertiary Qualifications - 2nd Year Subjects"

End Sub

Private Sub mnuViewEsub_Click()
Dim strSubjects As String

For a = 1 To nwcEducationalSubjects.Count
    strSubjects = strSubjects & a & "  -  " & nwcEducationalSubjects.Item(a) & vbCrLf
Next a

MsgBox strSubjects, vbInformation + vbOKOnly, "Educational Qualifications - Subjects"

End Sub

Private Sub mnuViewPadd_Click()
Dim strAddress As String

For a = 1 To nwcPersonalAddress.Count
    strAddress = strAddress & nwcPersonalAddress.Item(a) & vbCrLf
Next a

MsgBox strAddress, vbInformation + vbOKOnly, "Personal Details - Address"
End Sub

Private Sub mnuViewRaddress_Click()

Dim strAddress As String

For a = 1 To nwcReferenceAddress.Count
    strAddress = strAddress & nwcReferenceAddress.Item(a) & vbCrLf
Next a

MsgBox strAddress, vbInformation + vbOKOnly, "Reference - Address"

End Sub

Private Sub mnuViewTsub_Click()

Dim strSubjects As String


For a = 1 To nwcTertiarySubjects.Count
    strSubjects = strSubjects & a & "  -  " & nwcTertiarySubjects.Item(a) & vbCrLf
Next a

MsgBox strSubjects, vbInformation + vbOKOnly, "Tertiary Qualifications - Subjects"

End Sub

Private Sub mnuViewVCaddress_Click()

Dim strAddress As String

For a = 1 To nwcVoluntaryAddress.Count
    strAddress = strAddress & nwcVoluntaryAddress.Item(a) & vbCrLf
Next a

MsgBox strAddress, vbInformation + vbOKOnly, "Voluntary Work - Address"

End Sub

Private Sub mnuViewWCaddress_Click()

Dim strAddress As String

For a = 1 To nwcWorkAddress.Count
    strAddress = strAddress & nwcWorkAddress.Item(a) & vbCrLf
Next a

MsgBox strAddress, vbInformation + vbOKOnly, "Work Experience - Address"

End Sub

Private Sub tabFields_Click()
Select Case tabFields.SelectedItem.Key
        Case "Personal"
            'To Do: Hide other tabs
            picPersonal.Visible = True
            picEducational.Visible = False
            picTertiary.Visible = False
            picOther.Visible = False
            picVoluntary.Visible = False
            picWork.Visible = False
            picReference.Visible = False
        Case "Educational"
            'To Do: Hide other tabs
            picEducational.Visible = True
            picPersonal.Visible = False
            picTertiary.Visible = False
            picOther.Visible = False
            picVoluntary.Visible = False
            picWork.Visible = False
            picReference.Visible = False
        Case "Tertiary"
            'To Do: Hide other tabs
            picTertiary.Visible = True
            picEducational.Visible = False
            picPersonal.Visible = False
            picOther.Visible = False
            picVoluntary.Visible = False
            picWork.Visible = False
            picReference.Visible = False
        Case "Other"
            'To Do: Hide other tabs
            picOther.Visible = True
            picTertiary.Visible = False
            picEducational.Visible = False
            picPersonal.Visible = False
            picVoluntary.Visible = False
            picWork.Visible = False
            picReference.Visible = False
        Case "Voluntary"
            ' To Do: Hide other tabs
            picVoluntary.Visible = True
            picOther.Visible = False
            picTertiary.Visible = False
            picEducational.Visible = False
            picPersonal.Visible = False
            picWork.Visible = False
            picReference.Visible = False
        Case "Work"
            'To Do: Hide other tabs
            picWork.Visible = True
            picVoluntary.Visible = False
            picOther.Visible = False
            picTertiary.Visible = False
            picEducational.Visible = False
            picPersonal.Visible = False
            picReference.Visible = False
        Case "References"
            'To Do: Hide other tabs
            picWork.Visible = False
            picVoluntary.Visible = False
            picOther.Visible = False
            picTertiary.Visible = False
            picEducational.Visible = False
            picPersonal.Visible = False
            picReference.Visible = True
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
With Text1(Index)
'    .SetFocus
    .SelStart = 0
    .SelLength = Len(Text1(Index).Text)
End With
    'To Do: Set the default button
    frmDisplay.cmdAddPAddress.Default = True
End Sub

Private Sub Text2_GotFocus(Index As Integer)
With Text2(Index)
'    .SetFocus
    .SelStart = 0
    .SelLength = Len(Text2(Index).Text)
End With
    'To Do: Set the default button
    frmDisplay.cmdEduSub.Default = True
End Sub
