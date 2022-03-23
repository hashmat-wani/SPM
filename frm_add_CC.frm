VERSION 5.00
Begin VB.Form frm_add_CC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Class Co-ordinate"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Class Co-ordinator Details"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
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
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox cmb_teacher_name 
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
         Width           =   4095
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
         TabIndex        =   5
         Top             =   480
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
         TabIndex        =   4
         Top             =   960
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
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmd_save 
         Appearance      =   0  'Flat
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmd_rst 
         Appearance      =   0  'Flat
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2640
         Width           =   1335
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
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Class Co-ordinator"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   2055
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
         TabIndex        =   8
         Top             =   480
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
         TabIndex        =   7
         Top             =   960
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
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm_add_CC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_course_name_Click()
cmb_course_name.BackColor = vbWhite
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
cmb_dept_name.BackColor = vbWhite
Call connection
 rs.Open "select name from course where deptid =(select id from department where name = '" & cmb_dept_name.Text & "' )", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_course_name.Clear
           While (rs.EOF = False)
                cmb_course_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend
    Call connection
    rs.Open "select name from teacher where deptid =(select id from department where name = '" & cmb_dept_name & "')", conn, adOpenDynamic, adLockBatchOptimistic
    cmb_teacher_name.Clear
    While (rs.EOF = False)
        cmb_teacher_name.AddItem (rs.Fields(0))
        rs.MoveNext
    Wend


End Sub

Private Sub cmb_sec_click()
cmb_sec.BackColor = vbWhite
End Sub

Private Sub cmb_sem_Click()
cmb_sem.BackColor = vbWhite
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

Private Sub cmb_teacher_name_Click()
cmb_teacher_name.BackColor = vbWhite
End Sub

Private Sub cmd_rst_Click()
Call Form_Load
End Sub

Private Sub cmd_save_Click()
If validation = False Then
    Exit Sub
End If
Dim d_id, t_id, c_id, sem As Integer
Select Case cmb_sem.Text
Case "I": sem = 1
Case "II": sem = 2
Case "III": sem = 3
Case "IV": sem = 4
Case "V": sem = 5
Case "VI": sem = 6
Case "VII": sem = 7
Case "VIII": sem = 8
End Select

Call connection
rs.Open "select id from teacher where name = '" & cmb_teacher_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
t_id = rs.Fields(0)
Call connection
rs.Open "select id from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
c_id = rs.Fields(0)
Call connection
rs.Open "select id from department where name = '" & cmb_dept_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
d_id = rs.Fields(0)
Call connection
rs.Open "select count(deptid) from class_coordinator where deptid = " & d_id & " and courseid = " & c_id & " and semester = " & sem & " and section = '" & cmb_sec & "'", conn, adOpenDynamic, adLockBatchOptimistic
n = rs.Fields(0)
If n = 0 Then
    MsgBox "Record not found", vbCritical
    cmb_sem.Clear
    cmb_sem.SetFocus
Else
        Call connection
        rs.Open "update class_coordinator set teacherid= " & t_id & " where deptid = " & d_id & " and courseid = " & c_id & " and semester = " & sem & " and section = '" & cmb_sec & "'", conn, adOpenDynamic, adLockBatchOptimistic
       MsgBox "Record Saved successfully", vbInformation
    Call Form_Load
End If


End Sub

Private Sub Form_Load()
cmb_dept_name.Clear
cmb_course_name.Clear
cmb_sec.Clear
cmb_sem.Clear
cmb_teacher_name.Clear
For Each ctr In frm_add_CC.Controls
If TypeOf ctr Is ComboBox Or TypeOf ctr Is TextBox Then
ctr.BackColor = vbWhite
End If
Next
Call connection
rs.Open "select * from department", conn, adOpenDynamic, adLockBatchOptimistic
cmb_dept_name.Clear
           While (rs.EOF = False)
                cmb_dept_name.AddItem (rs.Fields(1))
                rs.MoveNext
            Wend


End Sub

Public Function validation()
          For Each ctr In frm_add_CC.Controls
           If TypeOf ctr Is ComboBox Or TypeOf ctr Is TextBox Then
            If Trim(ctr.Text) = Empty Then
                ctr.BackColor = vbGreen
            Else
                ctr.BackColor = vbWhite
            End If
           End If
          Next
       For Each ctr In frm_add_CC.Controls
            If TypeOf ctr Is TextBox Or TypeOf ctr Is ComboBox Then
                If ctr.BackColor = vbGreen Then
                    MsgBox "Fill all Fields", vbInformation
                    validation = False
                    Exit Function
                End If
            End If
        Next
validation = True
End Function

