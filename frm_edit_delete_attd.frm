VERSION 5.00
Begin VB.Form frm_edit_delete_attd 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "select"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton cmd_back 
         Height          =   375
         Left            =   120
         Picture         =   "frm_edit_delete_attd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Back"
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox cmb_year 
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
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmb_month 
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
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
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
         TabIndex        =   11
         Top             =   360
         Width           =   4695
      End
      Begin VB.TextBox txt_regno 
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
         Left            =   5040
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cmb_subject_name 
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
         Top             =   840
         Width           =   4695
      End
      Begin VB.Label Label7 
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
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Reg.No."
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
         Left            =   4200
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lbl_sub 
         Caption         =   "Subject"
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
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Month/year"
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
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Attendence profile of"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   3975
      Begin VB.TextBox txt_lect_deliverd 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txt_lect_attended 
         Alignment       =   2  'Center
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
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lbl_std_name 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label9 
         Caption         =   "Name :-"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Lect.  Delliverd"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Lect.  Attended"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1815
      Left            =   4200
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
      Begin VB.CommandButton cmd_update 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_edit_delete_attd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim teacher_rec As New ADODB.Recordset
Dim sec As String



Private Sub cmb_course_name_Click()

        Call connection
        rs.Open "select distinct(name) from subjects where courseid = (select id from course where name = '" & cmb_course_name & "') and teacherid = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
       
        cmb_subject_name.Clear
        While (rs.EOF = False)
        cmb_subject_name.AddItem rs.Fields(0)
    rs.MoveNext
    Wend


End Sub


Private Sub cmd_back_Click()
Unload Me
frm_teacher_login.Show
End Sub

Private Sub cmd_delete_Click()
Dim month_no As Integer
    Select Case cmb_month.Text
        Case MonthName(1)
            month_no = 1
        Case MonthName(2)
            month_no = 2
        Case MonthName(3)
            month_no = 3
        Case MonthName(4)
            month_no = 4
        Case MonthName(5)
            month_no = 5
        Case MonthName(6)
            month_no = 6
        Case MonthName(7)
            month_no = 7
        Case MonthName(8)
            month_no = 8
        Case MonthName(9)
            month_no = 9
        Case MonthName(10)
            month_no = 10
        Case MonthName(11)
            month_no = 11
        Case MonthName(12)
            month_no = 12
            
    End Select
warning = MsgBox("Are you sure?", vbYesNo + vbQuestion, "warning")
If warning = vbYes Then

Call connection
rs.Open "delete from attendance  where courseid = (select id from course where name = '" & cmb_course_name & "') and subjectid =(select id from subjects where name = '" & cmb_subject_name.Text & "' and teacherid = " & teacher_id & " and section = '" & sec & "') and regno= '" & Trim(txt_regno.Text) & "' and att_month='" & month_no & "' and att_year= '" & cmb_year.Text & "'", conn, adOpenDynamic, adLockBatchOptimistic
              
                 MsgBox "Record Deleted Successfully", vbInformation
txt_regno.Text = ""
txt_regno.SetFocus
txt_lect_deliverd.Text = ""
txt_lect_attended.Text = ""

End If
               
End Sub

Private Sub cmd_edit_Click()
cmd_update.Enabled = True
txt_lect_attended.Locked = False
End Sub

Private Sub cmd_update_Click()
cmd_update.Enabled = False
If txt_lect_attended.Text = Empty Then
    txt_lect_attended.BackColor = vbGreen
    MsgBox "Empty field not allowed", vbInformation
    Exit Sub
Else
     txt_lect_attended.BackColor = vbWhite
End If
If Val(txt_lect_attended.Text) > Val(txt_lect_deliverd.Text) Then
     txt_lect_attended.ForeColor = vbRed
     MsgBox "lect attended can't greater then lect deliverd", vbCritical
     Exit Sub
Else
    txt_lect_attended.ForeColor = vbBlack
End If
Dim month_no As Integer
    Select Case cmb_month.Text
        Case MonthName(1)
            month_no = 1
        Case MonthName(2)
            month_no = 2
        Case MonthName(3)
            month_no = 3
        Case MonthName(4)
            month_no = 4
        Case MonthName(5)
            month_no = 5
        Case MonthName(6)
            month_no = 6
        Case MonthName(7)
            month_no = 7
        Case MonthName(8)
            month_no = 8
        Case MonthName(9)
            month_no = 9
        Case MonthName(10)
            month_no = 10
        Case MonthName(11)
            month_no = 11
        Case MonthName(12)
            month_no = 12
            
    End Select
    Call connection
                                                                         
    rs.Open "update attendance set lec_attended = '" & Trim(txt_lect_attended.Text) & "'  where courseid = (select id from course where name = '" & cmb_course_name.Text & "') and subjectid =(select id from subjects where name = '" & cmb_subject_name.Text & "' and teacherid = " & teacher_id & " and section = '" & sec & "') and regno= '" & Trim(txt_regno.Text) & "' and att_month='" & month_no & "' and att_year= '" & cmb_year.Text & "'", conn, adOpenDynamic, adLockBatchOptimistic
    
    MsgBox "Record Updated Sucessfully", vbInformation
      
      
txt_lect_attended.Locked = True
End Sub

Private Sub cmd_view_Click()

        Call connection
        rs.Open "select name,section from student where courseid =(select id from course where name = '" & cmb_course_name.Text & "') and regno = '" & Trim(txt_regno.Text) & "'", conn, adOpenDynamic, adLockBatchOptimistic
          If rs.EOF = True Then
          MsgBox "No record exists of this search. check searching contents", vbInformation
          Exit Sub
          Else
            lbl_std_name.Caption = rs.Fields(0)
            sec = rs.Fields(1)
           End If
Dim month_no As Integer
    Select Case cmb_month.Text
        Case MonthName(1)
            month_no = 1
        Case MonthName(2)
            month_no = 2
        Case MonthName(3)
            month_no = 3
        Case MonthName(4)
            month_no = 4
        Case MonthName(5)
            month_no = 5
        Case MonthName(6)
            month_no = 6
        Case MonthName(7)
            month_no = 7
        Case MonthName(8)
            month_no = 8
        Case MonthName(9)
            month_no = 9
        Case MonthName(10)
            month_no = 10
        Case MonthName(11)
            month_no = 11
        Case MonthName(12)
            month_no = 12
            
    End Select
Call connection
rs.Open "select lec_delivered,lec_attended from attendance where courseid =(select id from course where name = '" & cmb_course_name.Text & "') and regno = '" & Trim(txt_regno.Text) & "' and subjectid = (select id from subjects where name = '" & cmb_subject_name.Text & "' and teacherid = " & teacher_id & " and section = '" & sec & "') and att_year=  '" & cmb_year.Text & "'  and att_month = '" & month_no & "'", conn, adOpenDynamic, adLockBatchOptimistic
  
          If rs.EOF = True Then
          MsgBox "No attendance exists for this month.", vbInformation
            txt_lect_deliverd.Text = ""
            txt_lect_attended.Text = ""
          
            Exit Sub
          End If
       txt_lect_deliverd.Text = rs.Fields(0)
txt_lect_attended.Text = rs.Fields(1)
cmd_edit.Enabled = True

cmd_delete.Enabled = True
            
End Sub



Private Sub Form_Load()
Call connection

rs.Open "select distinct(courseid) from subjects where teacherid = " & teacher_id & "", conn, adOpenDynamic, adLockOptimistic
cmb_course_name.Clear
    
    While (rs.EOF = False)
       If teacher_rec.State = 1 Then teacher_rec.Close
       teacher_rec.Open "select name from course where id = " & rs.Fields(0) & "", conn, adOpenDynamic, adLockBatchOptimistic
           
           While (teacher_rec.EOF = False)
                cmb_course_name.AddItem (teacher_rec.Fields(0))
                teacher_rec.MoveNext
            Wend
     rs.MoveNext
Wend

txt_regno.Locked = False
txt_lect_deliverd.Locked = True
txt_lect_attended.Locked = True
cmd_edit.Enabled = False
cmd_update.Enabled = False
cmd_delete.Enabled = False

For i = 1 To 12
    cmb_month.AddItem MonthName(i)
Next

For i = 2000 To 2050
    cmb_year.AddItem i
Next
End Sub


Private Sub subject_name_Click()
txt_rollno.Locked = False
End Sub
