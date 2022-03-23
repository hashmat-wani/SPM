VERSION 5.00
Begin VB.Form Frm_sub_assignment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subject Allotment"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7320
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
      TabIndex        =   5
      Top             =   120
      Width           =   6855
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
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
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
         TabIndex        =   9
         Top             =   960
         Width           =   4815
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
         TabIndex        =   8
         Top             =   480
         Width           =   4815
      End
      Begin VB.CommandButton cmd_teacher_list 
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
         Left            =   3240
         TabIndex        =   7
         Top             =   2040
         Width           =   435
      End
      Begin VB.TextBox txt_id 
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1920
         Width           =   1470
      End
      Begin VB.Label Label5 
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
         Left            =   3840
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Left            =   480
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "TeacherID"
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
         TabIndex        =   13
         Top             =   2040
         Width           =   1095
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Teacher ID"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   1440
         Width           =   15
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Action"
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
      Left            =   5520
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
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
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmd_reset 
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
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Subjects"
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
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   5175
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   360
         Width           =   4815
      End
   End
End
Attribute VB_Name = "Frm_sub_assignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_course_name_Click()
cmb_course_name.BackColor = vbWhite
cmb_sem.Clear
cmb_sec.Clear
List1.Clear
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
cmb_dept_name.BackColor = vbWhite

cmb_sem.Clear
cmb_sec.Clear
List1.Clear
Call connection
 rs.Open "select name from course where deptid =(select id from department where name = '" & cmb_dept_name.Text & "' )", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_course_name.Clear
           While (rs.EOF = False)
                cmb_course_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend
End Sub

Private Sub cmb_sec_click()
Call connection
rs.Open " select name from subjects where courseid=(select id from course where name = '" & cmb_course_name & "') and semester = '" & cmb_sem & "' and section = '" & cmb_sec & "'", conn, adOpenDynamic, adLockBatchOptimistic
List1.Clear
While (rs.EOF = False)
                List1.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend

End Sub

Private Sub cmb_sem_Click()
cmb_sem.BackColor = vbWhite
List1.Clear
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

Private Sub cmd_reset_Click()
Call Form_Load
End Sub

Private Sub cmd_save_Click()
If validation = False Then
    Exit Sub
End If
Dim cid As Integer
Call connection
rs.Open "select id from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
cid = rs.Fields(0)
If List1.SelCount = 0 Then
    MsgBox "Select minimum one subject", vbCritical
Else
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            Call connection
            rs.Open "update subjects set teacherid = '" & Trim(txt_id.Text) & "' where id = (select id from subjects where name = '" & List1.List(i) & "' and courseid = " & cid & " and semester = '" & cmb_sem & "' and section = '" & cmb_sec & "')", conn, adOpenDynamic, adLockBatchOptimistic
        End If
    Next
    MsgBox "Record Saved successfully", vbInformation
    Call Form_Load
End If
End Sub

Private Sub cmd_teacher_list_Click()
Frm_getteacher_id.Show
End Sub

Private Sub Form_Load()
txt_id.Enabled = False
cmb_dept_name.Clear
cmb_course_name.Clear
txt_id = ""
cmb_sem.Clear
cmb_sec.Clear
List1.Clear
For Each ctr In Frm_sub_assignment.Controls
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
          For Each ctr In Frm_sub_assignment.Controls
           If TypeOf ctr Is ComboBox Or TypeOf ctr Is TextBox Then
            If Trim(ctr.Text) = Empty Then
                ctr.BackColor = vbGreen
            Else
                ctr.BackColor = vbWhite
            End If
           End If
          Next
       For Each ctr In Frm_sub_assignment.Controls
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


