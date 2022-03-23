VERSION 5.00
Begin VB.Form Frm_add_course 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Courses"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Cmd_reset 
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
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox cmb_sems 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1390
      Width           =   1095
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
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txt_course_name 
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
      Left            =   2160
      TabIndex        =   5
      Text            =   "70"
      Top             =   960
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Course Details"
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
      Width           =   6615
      Begin VB.TextBox txt_shortage 
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
         Left            =   3000
         TabIndex        =   11
         Text            =   "70"
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox cmb_sections 
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
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1300
         Width           =   1095
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Shortage Criteria in %age"
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
         TabIndex        =   12
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "No. of Sections"
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
         Left            =   3480
         TabIndex        =   9
         Top             =   1350
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Department "
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
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Course Name"
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
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "No. of Semesters"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1350
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Frm_add_course"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n, m_id, d_id As Integer
Private Sub cmb_dept_name_Click()
cmb_dept_name.BackColor = vbWhite

End Sub


Private Sub cmb_sems_Click()
cmb_sems.BackColor = vbWhite
End Sub

Private Sub cmd_reset_Click()
Call Form_Load
End Sub

Private Sub cmd_save_Click()
If validation = False Then
    Exit Sub
End If

Call connection
rs.Open "select count(id) from course where name = '" & Trim(txt_course_name.Text) & "'", conn, adOpenDynamic, adLockBatchOptimistic
n = rs.Fields(0)
If n > 0 Then
    MsgBox "Record already exists", vbCritical
    txt_course_name.ForeColor = vbRed
    Exit Sub
Else
    Call connection
    rs.Open "select id from department where name = '" & cmb_dept_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
    d_id = rs.Fields(0)
    
    Call connection
    rs.Open "select max(id) from course", conn, adOpenDynamic, adLockBatchOptimistic
    If IsNull(rs.Fields(0)) Then
    m_id = 0
    Call connection
        rs.Open "insert into course values(1,'" & Trim(txt_course_name.Text) & "'," & d_id & ",'" & Val(cmb_sems) & "','" & Val(cmb_sections) & "', '" & txt_shortage & "')", conn, adOpenDynamic, adLockBatchOptimistic
    Else
        m_id = rs.Fields(0)
    Call connection
        rs.Open "insert into course values(" & m_id + 1 & " ,'" & Trim(txt_course_name.Text) & "'," & d_id & ",'" & Val(cmb_sems) & "','" & Val(cmb_sections) & "','" & txt_shortage & "')", conn, adOpenDynamic, adLockBatchOptimistic
    End If
        
End If
Dim sem As String
Dim i, j As Integer
For i = 1 To cmb_sems.Text
For j = 65 To 64 + cmb_sections.Text
    If m_id = 0 Then
    Call connection
        rs.Open "insert into class_coordinator(deptid,courseid,semester,section) values(" & d_id & " ,1, " & i & ",'" & Chr$(j) & "')", conn, adOpenDynamic, adLockBatchOptimistic
    Else
        Call connection
            rs.Open "insert into class_coordinator(deptid,courseid,semester,section) values(" & d_id & " ," & m_id + 1 & ", " & i & ",'" & Chr$(j) & "')", conn, adOpenDynamic, adLockBatchOptimistic
  End If
Next j
Next i
MsgBox "Record Saved successfully", vbInformation
Call Form_Load
End Sub

Private Sub Form_Load()
txt_shortage.Text = ""
cmb_dept_name.Clear
cmb_sems.Clear
With cmb_sems
    .AddItem 1
    .AddItem 2
    .AddItem 3
    .AddItem 4
    .AddItem 5
    .AddItem 6
    .AddItem 7
    .AddItem 8
End With
cmb_sections.Clear
With cmb_sections
    .AddItem 1
    .AddItem 2
    .AddItem 3
    .AddItem 4
    .AddItem 5
End With

txt_course_name = ""
For Each ctr In Frm_add_course.Controls
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
          For Each ctr In Frm_add_course.Controls
           If TypeOf ctr Is ComboBox Or TypeOf ctr Is TextBox Then
            If Trim(ctr.Text) = Empty Then
                ctr.BackColor = vbGreen
            Else
                ctr.BackColor = vbWhite
            End If
           End If
          Next
       For Each ctr In Frm_add_course.Controls
            If TypeOf ctr Is TextBox Or TypeOf ctr Is ComboBox Then
                If ctr.BackColor = vbGreen Then
                    MsgBox "Fill all fields", vbInformation
                    validation = False
                    Exit Function
                End If
            End If
        Next
validation = True
End Function


Private Sub txt_course_name_Change()
txt_course_name.ForeColor = vbBlack

End Sub

Private Sub txt_course_name_Click()
txt_course_name.BackColor = vbWhite
End Sub

Private Sub txt_shortage_Change()
txt_shortage.BackColor = vbWhite
End Sub

Private Sub txt_shortage_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If

End Sub
