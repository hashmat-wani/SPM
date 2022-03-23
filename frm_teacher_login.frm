VERSION 5.00
Begin VB.Form frm_teacher_login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teacher"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
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
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000010&
         Caption         =   "Edit Student's Attendance"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   3615
      End
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   3615
      End
      Begin VB.CommandButton cmd_attendance 
         BackColor       =   &H80000010&
         Caption         =   "Enter Student's Attendance"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label lbl_teacher_username 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   3960
         TabIndex        =   3
         Top             =   240
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frm_teacher_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_attendance_Click()
Me.Hide
frm_insert_att.Show
End Sub

Private Sub cmd_view_Click()
Me.Hide
frm_teacher_profile.Show
End Sub





Private Sub Command1_Click()
Me.Hide

frm_edit_delete_attd.Show
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
