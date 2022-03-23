VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_view_ind_attandence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View attendance individually"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8490
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_back 
      Height          =   375
      Left            =   240
      Picture         =   "frm_view_ind_attandence.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Back"
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Profile"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   8055
      Begin VB.Label lbl_sec 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Section :-"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label label1 
         Caption         =   "Roll No :-"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbl_regno 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "17RNSB7047"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Name :-"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl_name 
         Caption         =   "Hashmat Wani"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Course :-"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lbl_course 
         Caption         =   "BCA"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Semester :-"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4995
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lbl_sem 
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Mnth/yr:-"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lbl_mnth_yr 
         Caption         =   "08/2019"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6360
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2175
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   8
      Cols            =   4
      FixedCols       =   0
      ForeColorFixed  =   255
      BackColorSel    =   12632256
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483633
      GridColorFixed  =   -2147483633
      FillStyle       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   8280
      X2              =   8280
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   2280
      X2              =   8280
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   240
      X2              =   360
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   240
      X2              =   240
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Label Label11 
      Caption         =   "Attendence Details"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "frm_view_ind_attandence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_back_Click()
frm_entry_att.Show
Unload Me

End Sub


Private Sub Form_Load()

End Sub
