VERSION 5.00
Begin VB.Form frm_add_sub 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subjects"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7215
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
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
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
         TabIndex        =   10
         Top             =   1920
         Width           =   1335
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
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
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
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   1335
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   2655
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox txt_sub_name 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Text            =   "txt_sub_name"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "sem"
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
         Left            =   4560
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Subject name"
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
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
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
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   735
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
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_add_sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_dept_name_Click()
cmb_dept_name.BackColor = vbWhite
Call connection
 rs.Open "select name from course where deptid =(select id from department where name = '" & cmb_dept_name.Text & "' )", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_course_name.Clear
           While (rs.EOF = False)
                cmb_course_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend

End Sub

Private Sub cmb_course_name_Click()
cmb_course_name.BackColor = vbWhite
cmb_sem.Clear
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
Public Function validation()
          For Each ctr In frm_add_sub.Controls
           If TypeOf ctr Is ComboBox Or TypeOf ctr Is TextBox Then
            If Trim(ctr.Text) = Empty Then
                ctr.BackColor = vbGreen
            Else
                ctr.BackColor = vbWhite
            End If
           End If
          Next
       For Each ctr In frm_add_sub.Controls
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

Private Sub cmd_reset_Click()
Call Form_Load
End Sub

Private Sub cmb_sem_Click()
cmb_sem.BackColor = vbWhite
End Sub

Private Sub cmd_rst_Click()
Call Form_Load
End Sub

Private Sub cmd_save_Click()
If validation = False Then
    Exit Sub
End If
Dim n, c_id, m_id As Integer
Call connection
rs.Open "select id from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
c_id = rs.Fields(0)
Call connection
rs.Open "select count(id) from subjects where name = '" & Trim(txt_sub_name.Text) & "' and courseid = " & c_id & " and semester ='" & cmb_sem & "'", conn, adOpenDynamic, adLockBatchOptimistic
n = rs.Fields(0)
If n > 0 Then
    MsgBox "Record already exists", vbCritical
    txt_sub_name.ForeColor = vbRed
Else
    Call connection
    rs.Open "select no_of_secs from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
    secs = rs.Fields(0)
    Call connection
    rs.Open "select max(id) from subjects", conn, adOpenDynamic, adLockBatchOptimistic
    If IsNull(rs.Fields(0)) Then
        For i = 1 To secs
        Call connection
        rs.Open "insert into subjects(id,name,courseid,semester,section) values(" & i & ",'" & Trim(txt_sub_name.Text) & "'," & c_id & ",'" & cmb_sem & "','" & Chr$(64 + i) & "')", conn, adOpenDynamic, adLockBatchOptimistic
        Next
    Else
        m_id = rs.Fields(0)
        For i = 1 To secs
        Call connection
        rs.Open "insert into subjects(id,name,courseid,semester,section) values(" & m_id + i & ",'" & Trim(txt_sub_name.Text) & "'," & c_id & ",'" & cmb_sem & "','" & Chr$(64 + i) & "')", conn, adOpenDynamic, adLockBatchOptimistic
        Next
    End If
    MsgBox "Record Saved successfully", vbInformation
    txt_sub_name.Text = ""
End If

End Sub

Private Sub Form_Load()
cmb_dept_name.Enabled = True
cmb_course_name.Enabled = True
cmb_sem.Enabled = True
cmb_dept_name.Clear
cmb_course_name.Clear
cmb_sem.Clear
txt_sub_name = ""
For Each ctr In frm_add_sub.Controls
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

Private Sub txt_sub_name_Change()
txt_sub_name.ForeColor = vbBlack
End Sub

Private Sub txt_sub_name_Click()
txt_sub_name.BackColor = vbWhite
End Sub
