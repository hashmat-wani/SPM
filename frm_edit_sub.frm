VERSION 5.00
Begin VB.Form frm_edit_sub 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Subject Details"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6495
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
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2520
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2520
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
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2520
         Width           =   975
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   3375
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
         Left            =   1920
         TabIndex        =   6
         Text            =   "txt_sub_name"
         Top             =   960
         Width           =   2775
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
         TabIndex        =   2
         Top             =   360
         Width           =   1575
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
         TabIndex        =   1
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1575
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
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   1920
         Width           =   975
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
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frm_edit_sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim conn1 As New ADODB.connection
Dim cname, sname, sems As String
Private Sub cmb_course_name_Click()
cmb_course_name.BackColor = vbWhite
cmb_course_name.ForeColor = vbBlack
cmb_sems.Clear
Call connection
rs.Open "select no_of_sems from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
For i = 1 To rs.Fields(0)
    Select Case i
        Case 1: cmb_sems.AddItem "I"
        Case 2: cmb_sems.AddItem "II"
        Case 3: cmb_sems.AddItem "III"
        Case 4: cmb_sems.AddItem "IV"
        Case 5: cmb_sems.AddItem "V"
        Case 6: cmb_sems.AddItem "VI"
        Case 7: cmb_sems.AddItem "VII"
        Case 8: cmb_sems.AddItem "VIII"
    End Select
Next

End Sub

Private Sub cmb_sems_Click()
cmb_sems.BackColor = vbWhite
cmb_sems.ForeColor = vbBlack
End Sub


Private Sub cmd_back_Click()
frm_view_subject.Show
Unload Me
End Sub

Private Sub cmd_delete_Click()
warning = MsgBox("Are you sure?", vbYesNo + vbQuestion, "warning")
If warning = vbYes Then
Call connection
rs.Open "select no_of_secs from course where name = '" & cname & "'", conn, adOpenDynamic, adLockBatchOptimistic
    secs = rs.Fields(0)
    Call connection1
     rs2.Open "select id from subjects where name = '" & sname & "' and semester = '" & sems & "'", conn1, adOpenDynamic, adLockBatchOptimistic
   
For i = 1 To secs
Call connection
rs.Open "delete from subjects where id = " & Val(rs2.Fields(0)) & " ", conn, adOpenDynamic, adLockBatchOptimistic
rs2.MoveNext
Next
MsgBox "Record deleted successfully", vbInformation
Unload Me
frm_view_subject.Show
End If

End Sub

Private Sub cmd_edit_Click()
cmd_update.Enabled = True
cmd_delete.Enabled = True
txt_sub_name.Enabled = True
cmb_sems.Enabled = True
cname = cmb_course_name.Text
sname = txt_sub_name.Text
sems = cmb_sems.Text


End Sub



Private Sub cmd_update_Click()
If validation = False Then
    Exit Sub
End If
Call connection
rs.Open "select count(id) from subjects where name = '" & Trim(txt_sub_name.Text) & "' and courseid =(select id from course where name = '" & cmb_course_name & "') and semester = '" & cmb_sems & "'", conn, adOpenDynamic, adLockBatchOptimistic
n = rs.Fields(0)
If n > 0 Then
    MsgBox "Record already exists", vbCritical
        txt_sub_name.ForeColor = vbRed
        cmb_course_name.ForeColor = vbRed
        cmb_sems.ForeColor = vbRed
Else
Call connection
    Call connection
    rs.Open "select no_of_secs from course where name = '" & cname & "'", conn, adOpenDynamic, adLockBatchOptimistic
    secs = rs.Fields(0)
    Call connection1
    rs2.Open "select id from subjects where name = '" & sname & "' and semester = '" & sems & "' ", conn1, adOpenDynamic, adLockBatchOptimistic
     For i = 1 To secs
    Call connection
    rs.Open "update subjects set name = '" & Trim(txt_sub_name.Text) & "', semester = '" & cmb_sems & "' where id = " & Val(rs2.Fields(0)) & "", conn, adOpenDynamic, adLockBatchOptimistic
    rs2.MoveNext
    Next
MsgBox "Record Updated Successfully", vbInformation
Unload Me
frm_view_subject.Show
End If
End Sub

Private Sub Form_Load()
cmd_delete.Enabled = False
cmb_course_name.Enabled = False
cmd_update.Enabled = False
txt_sub_name.Enabled = False
cmb_course_name.Enabled = False
cmb_sems.Enabled = False
txt_id.Enabled = False
Call connection
 rs.Open "select name from course", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_course_name.Clear
           While (rs.EOF = False)
                cmb_course_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend
    cmb_sems.Clear
   With cmb_sems
            .AddItem "I"
            .AddItem "II"
            .AddItem "III"
            .AddItem "IV"
            .AddItem "V"
            .AddItem "VI"
            .AddItem "VII"
            .AddItem "VIII"
    End With

End Sub

Public Function validation()
          For Each ctr In frm_edit_sub.Controls
           If TypeOf ctr Is ComboBox Or TypeOf ctr Is TextBox Then
           If Trim(ctr.Text) = Empty Then
                ctr.BackColor = vbGreen
            Else
                ctr.BackColor = vbWhite
            End If
           End If
          Next
       For Each ctr In frm_edit_sub.Controls
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


Private Sub txt_sub_name_Change()
txt_sub_name.ForeColor = vbBlack
End Sub
Private Sub txt_sub_name_Click()
txt_sub_name.BackColor = vbWhite
End Sub
Public Function connection1()
If conn1.State = 1 Then conn1.Close
conn1.Open "DSN=Semester_Process_Management"
End Function

