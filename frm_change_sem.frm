VERSION 5.00
Begin VB.Form frm_change_sem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Semester "
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Subject Details"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.Timer Timer1 
         Interval        =   800
         Left            =   5280
         Top             =   240
      End
      Begin VB.ComboBox cmb_batch 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   780
         Width           =   3255
      End
      Begin VB.ComboBox cmb_course_name 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1270
         Width           =   2655
      End
      Begin VB.ComboBox cmb_sem 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmd_update 
         BackColor       =   &H80000004&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmd_rst 
         BackColor       =   &H80000004&
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lbl_cur_sem 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   3960
         TabIndex        =   10
         Top             =   1300
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "change semester for students of below batch and course"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label8 
         Caption         =   "Batch"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "New sem"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_change_sem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c_id As Integer
Private Sub cmb_batch_Click()
cmb_sem.Clear
lbl_cur_sem.Caption = ""
Call connection
rs.Open "select name from course", conn, adOpenDynamic, adLockBatchOptimistic
cmb_course_name.Clear
           While (rs.EOF = False)
                cmb_course_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend

End Sub

Private Sub cmb_course_name_Click()
lbl_cur_sem.Caption = ""
cmb_sem.Clear
Call connection
rs.Open "select id from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
c_id = rs.Fields(0)
Call connection
rs.Open "select semester from student where batch = '" & cmb_batch & "' and courseid = " & c_id & "", conn, adOpenDynamic, adLockBatchOptimistic
If rs.EOF = False Then
    lbl_cur_sem.Caption = "current semester:" + rs.Fields(0)
   cmb_sem.Clear
Call connection
rs.Open "select no_of_sems from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
For i = 1 To rs.Fields(0)
    Select Case i
        Case 1: cmb_sem.AddItem "-"
        Case 2: cmb_sem.AddItem "I"
        Case 3: cmb_sem.AddItem "II"
        Case 4: cmb_sem.AddItem "III"
        Case 5: cmb_sem.AddItem "IV"
        Case 6: cmb_sem.AddItem "V"
        Case 7: cmb_sem.AddItem "VI"
        Case 8: cmb_sem.AddItem "VII"
        Case 9: cmb_sem.AddItem "VIII"
    End Select
Next
End If
End Sub

Private Sub cmd_rst_Click()
Call Form_Load
End Sub

Private Sub cmd_update_Click()
Call connection
rs.Open "update student set semester = '" & cmb_sem & "' where courseid = " & c_id & " and batch = '" & cmb_batch & "'", conn, adOpenDynamic, adLockBatchOptimistic
MsgBox "Record updated successfully", vbInformation
Call Form_Load
End Sub

Private Sub Form_Load()
cmb_batch.Clear
cmb_course_name.Clear
cmb_sem.Clear
lbl_cur_sem.Caption = ""
For i = 1 To 50
cmb_batch.AddItem 1999 + i
Next

End Sub

Private Sub Timer1_Timer()
If Label1.ForeColor = vbRed Then
        Label1.ForeColor = vbBlue

ElseIf Label1.ForeColor = vbBlue Then
        Label1.ForeColor = vbGreen
Else
        Label1.ForeColor = vbRed
End If

End Sub
