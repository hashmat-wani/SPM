VERSION 5.00
Begin VB.Form frm_entry_att 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Attendance"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Attendence of"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   6615
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1920
         Width           =   1335
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   16
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   480
         Width           =   4335
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   4335
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
         Left            =   960
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label9 
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
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbl_month 
         Caption         =   "Month/Year"
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
         Left            =   600
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label5 
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
         Left            =   1080
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "View Mode"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   2895
      Begin VB.OptionButton opt_total 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton opt_mnth 
         Caption         =   "Month Wise"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
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
      Height          =   3975
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   6615
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000C0&
         Caption         =   "View Detained List"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   4695
      End
      Begin VB.CommandButton cmd_ind_att 
         Caption         =   "View  individual  Attendance "
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   3615
         Begin VB.CommandButton cmd_view_all_attendence 
            Caption         =   "View  all  Attendance "
            BeginProperty Font 
               Name            =   "Lucida Handwriting"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            MaskColor       =   &H00000000&
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   3015
         End
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
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Height          =   1215
         Left            =   1680
         TabIndex        =   14
         Top             =   1560
         Width           =   3615
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
            Left            =   720
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label3 
         Caption         =   "OR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   1200
         Width           =   495
      End
   End
   Begin VB.Label lbl_atm 
      BackStyle       =   0  'Transparent
      Caption         =   "ATTENDANCE"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frm_entry_att"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim attendance_rec As New ADODB.Recordset
Dim course_criteria As Integer

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

Private Sub cmb_dept_name_Click()
cmb_sem.Clear
cmb_dept_name.BackColor = vbWhite
Call connection
 rs.Open "select name from course where deptid =(select id from department where name = '" & cmb_dept_name.Text & "' )", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_course_name.Clear
           While (rs.EOF = False)
                cmb_course_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend
End Sub








Private Sub cmb_month_Click()
cmb_month.BackColor = vbWhite
End Sub

Private Sub cmb_sem_Click()
cmb_sem.BackColor = vbWhite
End Sub

Private Sub cmb_year_Click()
cmb_year.BackColor = vbWhite
End Sub



Private Sub cmd_ind_att_Click()
   X = "Subject Name " & Space(60) & "|^Lect deliverd " & Space(8) & "|^lect Attended " & Space(8) & "|^ %age" & Space(16)
   frm_view_ind_attandence.MSFlexGrid1.FormatString = X
    
    If opt_mnth.Value = True Then
        If cmb_month = Empty Or cmb_year = Empty Then
        MsgBox "Select Month/Year", vbInformation
        Exit Sub
        End If
    ElseIf cmb_year = Empty Then
        MsgBox "Select Year", vbInformation
        Exit Sub
    End If

    Call connection
    rs.Open "select name,courseid,semester,section from student where regno = '" & txt_regno.Text & "'", conn, adOpenDynamic, adLockBatchOptimistic
          If rs.EOF = True Then
          MsgBox "No record exists of this search", vbCritical
          Exit Sub
          Else
          c_id = rs.Fields(1)
          sem = rs.Fields(2)
          sec = rs.Fields(3)
          frm_view_ind_attandence.lbl_regno = txt_regno.Text
        frm_view_ind_attandence.lbl_name.Caption = rs.Fields(0)
         If attendance_rec.State = 1 Then attendance_rec.Close
         attendance_rec.Open "select name from course where id = " & c_id & "", conn, adOpenDynamic, adLockOptimistic
        frm_view_ind_attandence.lbl_course.Caption = attendance_rec.Fields(0)
        frm_view_ind_attandence.lbl_sem.Caption = rs.Fields(2)
        frm_view_ind_attandence.lbl_sec.Caption = rs.Fields(3)
        
          End If
Call connection
rs.Open "select shortage_criteria from course where id = " & c_id & "", conn, adOpenDynamic, adLockBatchOptimistic
course_criteria = rs.Fields(0)
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

If opt_mnth.Value = True Then
        frm_view_ind_attandence.lbl_mnth_yr.Caption = cmb_month.Text & "/" & cmb_year.Text
    
    Call connection
    rs.Open "select distinct(name) from subjects where courseid = " & c_id & " and semester = '" & sem & "'", conn, adOpenDynamic, adLockBatchOptimistic
   i = 1
   While (rs.EOF = False)
    frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 0) = rs.Fields(0)
                  If attendance_rec.State = 1 Then attendance_rec.Close
                                   
                  attendance_rec.Open "select lec_delivered,lec_attended from attendance where subjectid =(select id from subjects where name = '" & rs.Fields(0) & "' and section = '" & sec & "') and regno = '" & txt_regno.Text & "' and att_month= '" & month_no & "' and att_year = '" & cmb_year.Text & "'", conn, adOpenDynamic, adLockOptimistic
                    If attendance_rec.EOF = False Then
                        frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 1) = attendance_rec.Fields(0)
                        frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 2) = attendance_rec.Fields(1)
                        X = CInt(attendance_rec(0))
                        Y = CInt(attendance_rec(1))
                        z = (Y / X) * 100
                        
                        If Len(z) > 4 Then
                                frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 3) = Left(z, 4)
                        Else
                                 frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 3) = z
                        End If
                            If z < course_criteria Then
                       frm_view_ind_attandence.MSFlexGrid1.Row = i
                       For j = 0 To 3
                       frm_view_ind_attandence.MSFlexGrid1.Col = j
                       frm_view_ind_attandence.MSFlexGrid1.CellForeColor = vbRed
                       frm_view_ind_attandence.MSFlexGrid1.CellBackColor = &HE0E0E0
                        Next
                         End If
                       
                    Else
                        frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 1) = 0
                        frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 2) = 0
                        frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 3) = 0
                    End If
                        
    i = i + 1
   rs.MoveNext
    Wend
  
Else
    Call connection
        rs.Open "select max(att_month) from attendance where regno = '" & txt_regno.Text & "' and att_year = '" & cmb_year.Text & "'", conn, adOpenDynamic, adLockOptimistic
            If rs.EOF = False Then
                max_month = rs.Fields(0)
            End If
        
        frm_view_ind_attandence.lbl_mnth_yr.Caption = "upto " & max_month & "/" & cmb_year.Text
        
    Call connection
    rs.Open "select distinct(name) from subjects where courseid = " & c_id & " and semester = '" & sem & "'", conn, adOpenDynamic, adLockBatchOptimistic
   i = 1
   While (rs.EOF = False)
    frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 0) = rs.Fields(0)
                  If attendance_rec.State = 1 Then attendance_rec.Close
                  
                  attendance_rec.Open "select sum(lec_delivered),sum(lec_attended) from attendance where subjectid =(select id from subjects where name = '" & rs.Fields(0) & "' and section = '" & sec & "' ) and regno = '" & txt_regno.Text & "' and att_month <= '" & max_month & "' and att_year = '" & cmb_year.Text & "'", conn, adOpenDynamic, adLockOptimistic
                  
                    If Not (IsNull(attendance_rec.Fields(0))) Then
                        frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 1) = attendance_rec.Fields(0)
                        frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 2) = attendance_rec.Fields(1)
                        X = CInt(attendance_rec.Fields(0))
                        Y = CInt(attendance_rec.Fields(1))
                        z = (Y / X) * 100
                        
                        If Len(z) > 4 Then
                                frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 3) = Left(z, 4)
                        Else
                                 frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 3) = z
                        End If
                              If z < course_criteria Then
                       frm_view_ind_attandence.MSFlexGrid1.Row = i
                       For j = 0 To 3
                       frm_view_ind_attandence.MSFlexGrid1.Col = j
                       frm_view_ind_attandence.MSFlexGrid1.CellForeColor = vbRed
                       frm_view_ind_attandence.MSFlexGrid1.CellBackColor = &HE0E0E0
                        Next
                         End If
                       
                    Else
                        frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 1) = 0
                        frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 2) = 0
                        frm_view_ind_attandence.MSFlexGrid1.TextMatrix(i, 3) = 0
                    End If
                        
    i = i + 1
   rs.MoveNext
    Wend
    
 End If


frm_view_ind_attandence.Show
Unload Me
End Sub

Private Sub cmd_view_all_attendence_Click()
    If validation = False Then
        Exit Sub
    End If


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

    frm_view_all_attandence.cmb_sec.Clear
    Call connection
    rs.Open "select no_of_secs from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
    For i = 1 To rs.Fields(0)
    Select Case i
        Case 1: frm_view_all_attandence.cmb_sec.AddItem "A"
        Case 2: frm_view_all_attandence.cmb_sec.AddItem "B"
        Case 3: frm_view_all_attandence.cmb_sec.AddItem "C"
        Case 4: frm_view_all_attandence.cmb_sec.AddItem "D"
        Case 5: frm_view_all_attandence.cmb_sec.AddItem "E"
    End Select
Next


frm_view_all_attandence.Show
Hide
End Sub



Private Sub Command1_Click()
Unload Me
frm_detained_list.Show
End Sub

Private Sub Form_Load()
cmb_month.Enabled = True
opt_mnth.Value = True

    Call connection
        rs.Open "select name from department", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_dept_name.Clear
            While (rs.EOF = False)
                cmb_dept_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend
For i = 1 To 12
    cmb_month.AddItem MonthName(i)
Next

For i = 2000 To 2050
    cmb_year.AddItem i
Next
End Sub







Private Sub opt_mnth_Click()
cmb_month.Enabled = True
End Sub

Private Sub opt_total_Click()
cmb_month.Enabled = False
End Sub

Private Sub Timer1_Timer()
If lbl_atm.ForeColor = vbBlue Then
        lbl_atm.ForeColor = vbRed

ElseIf lbl_atm.ForeColor = vbRed Then
        lbl_atm.ForeColor = vbGreen
Else
        lbl_atm.ForeColor = vbBlue
End If
End Sub

Public Function validation()
          For Each ctr In frm_entry_att.Controls
           If TypeOf ctr Is ComboBox Then
           If ctr.Enabled = True Then
            If Trim(ctr.Text) = Empty Then
                 ctr.BackColor = vbGreen
            Else
                ctr.BackColor = vbWhite
            End If
           End If
           End If

          Next
       For Each ctr In frm_entry_att.Controls
            If TypeOf ctr Is ComboBox Then
                If ctr.BackColor = vbGreen Then
                    MsgBox "Fill all Fields", vbInformation
                    validation = False
                    Exit Function
                End If
            End If
        Next
validation = True
End Function

