VERSION 5.00
Begin VB.Form frm_edit_cr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update CR Details"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Update Class Representative Details"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmd_back 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   615
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1920
         Width           =   1455
      End
      Begin VB.ComboBox cmb_dept_name 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   4095
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1440
         Width           =   4095
      End
      Begin VB.ComboBox cmb_sec 
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
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmd_update 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Update"
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton cmd_delete 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Delete"
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton cmd_students_list 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   3
         Top             =   2520
         Width           =   435
      End
      Begin VB.TextBox txt_regno 
         DataSource      =   "Adodc1"
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
         Left            =   2280
         TabIndex        =   2
         Top             =   2400
         Width           =   2310
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Semester"
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
         Left            =   1200
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Students Reg.No."
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Department"
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
         Left            =   840
         TabIndex        =   13
         Top             =   960
         Width           =   1215
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
         Left            =   1320
         TabIndex        =   12
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Section"
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
         Left            =   3960
         TabIndex        =   11
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Batch"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_edit_cr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_course_name_Click()
cmb_sem.Clear
cmb_sec.Clear
Call connection
rs.Open "select no_of_sems from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
For i = 1 To rs.Fields(0)
    Select Case i
        Case 1: cmb_sem.AddItem "I"
        Case 2: cmb_sem.AddItem "II"
        Case 3: cmb_sem.AddItem "III"
        Case 4: cmb_sem.AddItem "IV"
        Case 5: cmb_sem.AddItem "V"
        Case 6: cmb_sem.AddItem "VI"
        Case 7: cmb_sem.AddItem "VII"
        Case 8: cmb_sem.AddItem "VIII"
    End Select
Next

End Sub

Private Sub cmb_dept_name_Click()
cmb_sem.Clear
cmb_sec.Clear
Call connection
 rs.Open "select name from course where deptid =(select id from department where name = '" & cmb_dept_name.Text & "' )", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_course_name.Clear
           While (rs.EOF = False)
                cmb_course_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend


End Sub

Private Sub cmb_sem_Click()
cmb_sec.Clear
Call connection
rs.Open "select no_of_secs from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
For i = 1 To rs.Fields(0)
    Select Case i
        Case 1: cmb_sec.AddItem "A"
        Case 2: cmb_sec.AddItem "B"
        Case 3: cmb_sec.AddItem "C"
        Case 4: cmb_sec.AddItem "D"
        Case 5: cmb_sec.AddItem "E"
    End Select
Next

End Sub


Private Sub cmd_back_Click()
Unload Me
frm_view_cr.Show
End Sub

Private Sub cmd_delete_Click()
warning = MsgBox("Are you sure?", vbYesNo + vbQuestion, "warning")
If warning = vbYes Then
Call connection
rs.Open "select id from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
c_id = rs.Fields(0)
Call connection
rs.Open "select id from department where name = '" & cmb_dept_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
d_id = rs.Fields(0)

Call connection
        rs.Open "delete from CR where regno = '" & Trim(txt_regno.Text) & "' and deptid = " & d_id & " and courseid = " & c_id & " and semester = '" & cmb_sem & "' and section = '" & cmb_sec & "' and batch = '" & cmb_batch & "'", conn, adOpenDynamic, adLockBatchOptimistic
        MsgBox "Record deleted successfully", vbInformation
Unload Me
frm_view_cr.Show
End If

End Sub

Private Sub cmd_update_Click()
If txt_regno.Text = "" Then
    txt_regno.BackColor = vbGreen
    MsgBox "select the students registration number", vbInformation
    End If
Dim d_id, c_id As Integer
Call connection
rs.Open "select id from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
c_id = rs.Fields(0)
Call connection
rs.Open "select id from department where name = '" & cmb_dept_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
d_id = rs.Fields(0)
Call connection
rs.Open "select count(regno) from CR where regno = '" & Trim(txt_regno.Text) & "' and semester = '" & cmb_sem & "'", conn, adOpenDynamic, adLockBatchOptimistic
n = rs.Fields(0)
If n > 0 Then
    MsgBox "Given Registration Number already exists", vbCritical
    txt_regno.Text = ""
    txt_regno.BackColor = vbGreen
Else
        Call connection
        rs.Open "update CR set regno = '" & Trim(txt_regno.Text) & "' where deptid = " & d_id & " and courseid = " & c_id & " and semester = '" & cmb_sem & "' and section = '" & cmb_sec & "' and batch = '" & cmb_batch & "'", conn, adOpenDynamic, adLockBatchOptimistic
       MsgBox "Record updated successfully", vbInformation
        Unload Me
        frm_view_cr.Show
End If


End Sub

Private Sub cmd_students_list_Click()
frm_edit_get_regno.Show
End Sub

Private Sub Form_Load()
txt_regno.Text = ""
cmb_dept_name.Clear
cmb_course_name.Clear
cmb_sec.Clear
cmb_sem.Clear
cmb_batch.Clear

txt_regno.Enabled = False
cmb_dept_name.Enabled = False
cmb_course_name.Enabled = False
cmb_sec.Enabled = False
cmb_sem.Enabled = False
cmb_batch.Enabled = False
For i = 1 To 50
cmb_batch.AddItem 1999 + i
Next

Call connection
rs.Open "select * from department", conn, adOpenDynamic, adLockBatchOptimistic
cmb_dept_name.Clear
           While (rs.EOF = False)
                cmb_dept_name.AddItem (rs.Fields(1))
                rs.MoveNext
            Wend


End Sub



Private Sub txt_regno_Change()
txt_regno.BackColor = vbWhite
End Sub

