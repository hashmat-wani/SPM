VERSION 5.00
Begin VB.Form frm_teacher_login 
   Caption         =   "Teacher"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Teacher"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5520
         Top             =   1920
      End
      Begin VB.CommandButton cmd_view 
         BackColor       =   &H80000010&
         Caption         =   "View Profile"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton cmd_attendance 
         BackColor       =   &H80000010&
         Caption         =   "Enter Student's Attendance"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3120
         Width           =   3495
      End
      Begin VB.Label lbl_logout 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbl_teacher_username 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   240
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frm_teacher_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_view_Click()
Me.Hide
frm_teacher_profile.Show
End Sub





Private Sub lbl_logout_Click()
Unload Me
frm_login.Show
End Sub


Private Sub Timer1_Timer()
If lbl_teacher_username.ForeColor = vbBlue Then
        lbl_teacher_username.ForeColor = vbRed

ElseIf lbl_teacher_username.ForeColor = vbRed Then
        lbl_teacher_username.ForeColor = vbGreen
Else
        lbl_teacher_username.ForeColor = vbBlue
End If

End Sub
