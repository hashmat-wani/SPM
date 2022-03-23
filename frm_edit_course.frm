VERSION 5.00
Begin VB.Form frm_edit_course 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Course"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "UPDATE Details"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
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
         TabIndex        =   15
         Text            =   "70"
         Top             =   2400
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
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmd_edit 
         BackColor       =   &H80000004&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3000
         Width           =   975
      End
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
         TabIndex        =   11
         Top             =   360
         Width           =   615
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
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
         Height          =   405
         Left            =   1920
         TabIndex        =   9
         Text            =   "70"
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox txt_id 
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
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmd_update 
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
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmd_delete 
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
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3000
         Width           =   975
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label Label6 
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
         TabIndex        =   16
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label5 
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
         Left            =   3360
         TabIndex        =   14
         Top             =   1965
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "ID"
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
         Left            =   1440
         TabIndex        =   8
         Top             =   480
         Width           =   495
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
         TabIndex        =   4
         Top             =   1965
         Width           =   1695
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
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_edit_course"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n_value, o_section, n_section As Integer
Dim o_cname, o_dname, n_cname, n_dname As String



Private Sub cmb_sems_Click()
cmb_sems.BackColor = vbWhite

End Sub

Private Sub cmd_back_Click()

frm_view_courses.Show
Unload Me
End Sub

Private Sub cmd_delete_Click()
warning = MsgBox("Are you sure?", vbYesNo + vbQuestion, "warning")
If warning = vbYes Then
Call connection
rs.Open "delete from course where id = " & Val(Trim(txt_id.Text)) & " ", conn, adOpenDynamic, adLockBatchOptimistic
MsgBox "Record deleted successfully", vbInformation
Unload Me
frm_view_courses.Show
End If
End Sub

Private Sub cmd_edit_Click()
txt_shortage.Enabled = True
txt_course_name.Enabled = True
cmb_dept_name.Enabled = True
cmb_sems.Enabled = True
cmb_sections.Enabled = True
cmd_update.Enabled = True
o_cname = txt_course_name.Text
o_dname = cmb_dept_name.Text
o_section = Val(cmb_sections)

End Sub

Private Sub cmd_update_Click()
n_value = Val(cmb_sems.Text)
n_cname = txt_course_name.Text
n_dname = cmb_dept_name.Text
n_section = Val(cmb_sections)
If validation = False Then
    Exit Sub
End If
        Call connection
        rs.Open "select id from department where name = '" & cmb_dept_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
        d_id = rs.Fields(0)
    If o_cname = n_cname Then
        Call connection
        rs.Open "update course set deptid = " & d_id & ", no_of_sems = " & Val(cmb_sems) & ", no_of_secs = " & Val(cmb_sections) & ", shortage_criteria = '" & Trim(txt_shortage.Text) & "' where id = " & Val(Trim(txt_id.Text)) & " ", conn, adOpenDynamic, adLockBatchOptimistic
    Else
        Call connection
        rs.Open "select count(id) from course where name = '" & Trim(txt_course_name.Text) & "'", conn, adOpenDynamic, adLockBatchOptimistic
        n = rs.Fields(0)
        If n > 0 Then
            MsgBox "Record With this Course name already exists", vbCritical
            txt_course_name.ForeColor = vbRed
            Unload Me
            frm_view_courses.Show
            Exit Sub
        Else
            Call connection
            rs.Open "update course set name = '" & Trim(txt_course_name.Text) & "', deptid = " & d_id & ", no_of_sems = " & Val(cmb_sems) & ", no_of_secs = " & Val(cmb_sections) & ", shortage_criteria = '" & Trim(txt_shortage.Text) & "' where id = " & Val(Trim(txt_id.Text)) & " ", conn, adOpenDynamic, adLockBatchOptimistic
    
        End If
    End If
    
    If o_dname <> n_dname Then
    Call connection
     rs.Open "update class_coordinator set deptid = " & d_id & " where courseid = " & Val(Trim(txt_id.Text)) & "", conn, adOpenDynamic, adLockBatchOptimistic
    End If
    
        
        If (o_section > n_section) Then
            Call connection
            rs.Open "delete from class_coordinator where section > '" & Chr$(64 + n_section) & "' and courseid = " & Val(Trim(txt_id.Text)) & "", conn, adOpenDynamic, adLockBatchOptimistic
            
        Else
            For i = 1 To n_value
            For j = o_section + 1 To n_section

                Call connection
                rs.Open "insert into class_coordinator(deptid,courseid,semester,section) values(" & d_id & " ," & Val(Trim(txt_id.Text)) & ", " & i & ",'" & Chr$(64 + j) & "')", conn, adOpenDynamic, adLockBatchOptimistic
            Next j
            Next i
         End If
        
        
        
        
        If (frmEditCourse_oldsem > n_value) Then
            Call connection
            rs.Open "delete from class_coordinator where semester > " & n_value & " and courseid = " & Val(Trim(txt_id.Text)) & "", conn, adOpenDynamic, adLockBatchOptimistic
            
        Else
            For i = frmEditCourse_oldsem + 1 To n_value
            For j = 65 To 64 + cmb_sections
                Call connection
                rs.Open "insert into class_coordinator(deptid,courseid,semester,section) values(" & d_id & " ," & Val(Trim(txt_id.Text)) & ", " & i & ",'" & Chr$(j) & "')", conn, adOpenDynamic, adLockBatchOptimistic
            Next j
            Next i
         End If
        MsgBox "Record Updated Successfully", vbInformation
        Unload Me
        frm_view_courses.Show
End Sub
Private Sub Form_Load()
txt_shortage.Enabled = False
cmb_dept_name.Enabled = False
cmb_sems.Enabled = False
cmb_sections.Enabled = False
cmd_update.Enabled = False
txt_id.Enabled = False
txt_course_name.Enabled = False
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


    Call connection
rs.Open "select * from department", conn, adOpenDynamic, adLockBatchOptimistic
cmb_dept_name.Clear
           While (rs.EOF = False)
                cmb_dept_name.AddItem (rs.Fields(1))
                rs.MoveNext
            Wend
End Sub

Private Sub cmb_dept_name_Click()
cmb_dept_name.BackColor = vbWhite
End Sub



Private Sub txt_course_name_Change()
txt_course_name.ForeColor = vbBlack
End Sub

Private Sub txt_course_name_Click()
txt_course_name.BackColor = vbWhite
End Sub

Public Function validation()
          For Each ctr In frm_edit_course.Controls
           If TypeOf ctr Is ComboBox Or TypeOf ctr Is TextBox Then
            If Trim(ctr.Text) = Empty Then
                ctr.BackColor = vbGreen
            Else
                ctr.BackColor = vbWhite
            End If
           End If
          Next
       For Each ctr In frm_edit_course.Controls
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
