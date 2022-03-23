VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_detained_list 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detain List"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Select"
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
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton cmd_back 
         Height          =   375
         Left            =   120
         Picture         =   "frm_detained_list.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Back"
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox cmb_sub 
         Height          =   375
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox cmb_sem 
         Height          =   375
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cmb_section 
         Height          =   375
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cmb_year 
         Height          =   375
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cmb_dept_name 
         Height          =   375
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   5055
      End
      Begin VB.ComboBox cmb_course_name 
         Height          =   375
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label Label3 
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
         Left            =   1080
         TabIndex        =   15
         Top             =   1800
         Width           =   855
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
         TabIndex        =   11
         Top             =   1320
         Width           =   975
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
         Left            =   4440
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label5 
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
         Left            =   5280
         TabIndex        =   9
         Top             =   720
         Width           =   495
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
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Year"
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
         Left            =   4680
         TabIndex        =   7
         Top             =   1800
         Width           =   495
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
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3375
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   30
      Cols            =   5
      FixedCols       =   0
      BackColorSel    =   14737632
      BackColorBkg    =   -2147483633
      GridColorFixed  =   12632256
      FillStyle       =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl_sub_course_name 
      Alignment       =   2  'Center
      Caption         =   "Datastructures attendanceBca 3rd sem) "
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   2760
      Width           =   9135
   End
   Begin VB.Label lbl_month_year 
      Alignment       =   2  'Center
      Caption         =   " upto 30th sep"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   3120
      Width           =   7455
   End
   Begin VB.Label Label11 
      Caption         =   "Detained Students"
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
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   120
      Y1              =   3720
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   240
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   2160
      X2              =   8160
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   8160
      X2              =   8160
      Y1              =   3720
      Y2              =   3480
   End
End
Attribute VB_Name = "frm_detained_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim attendance_rec As New ADODB.Recordset


Private Sub cmb_course_name_Click()
cmb_course_name.BackColor = vbWhite
MSFlexGrid1.Clear
cmb_sem.Clear
cmb_section.Clear
cmb_sub.Clear
cmb_year.Clear
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
MSFlexGrid1.Clear
cmb_sem.Clear
cmb_section.Clear
cmb_sub.Clear
cmb_year.Clear
cmb_dept_name.BackColor = vbWhite
Call connection
 rs.Open "select name from course where deptid =(select id from department where name = '" & cmb_dept_name.Text & "' )", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_course_name.Clear
           While (rs.EOF = False)
                cmb_course_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend

End Sub





Private Sub cmb_section_Click()
cmb_section.BackColor = vbWhite
cmb_year.Clear

Call connection
        rs.Open "select name from subjects where courseid=(select id from course where name = '" & cmb_course_name.Text & "') and semester = '" & cmb_sem & "' and section = '" & cmb_section & "'", conn, adOpenDynamic, adLockBatchOptimistic
          cmb_sub.Clear
          While (rs.EOF = False)
            cmb_sub.AddItem (rs.Fields(0))
            rs.MoveNext
          Wend

End Sub

Private Sub cmb_sem_Click()
cmb_sem.BackColor = vbWhite
cmb_sub.Clear
cmb_year.Clear
    MSFlexGrid1.Clear
    cmb_section.Clear
    Call connection
    rs.Open "select no_of_secs from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
    For i = 1 To rs.Fields(0)
    Select Case i
        Case 1: cmb_section.AddItem "A"
        Case 2: cmb_section.AddItem "B"
        Case 3: cmb_section.AddItem "C"
        Case 4: cmb_section.AddItem "D"
        Case 5: cmb_section.AddItem "E"
    End Select
Next


    
End Sub

Private Sub cmb_sub_Click()
cmb_sub.BackColor = vbWhite
cmb_year.Clear
For i = 2000 To 2050
    cmb_year.AddItem i
Next

End Sub

Private Sub cmb_year_Click()
If validation = False Then
    Exit Sub
End If

MSFlexGrid1.Clear

Dim X As String

 
    X = "^ Reg.No. " & Space(5) & " |<  Name " & Space(40) & "|^Lect  delivered " & Space(5) & "|^lect Attended " & Space(5) & "|^ %age" & Space(8)
    MSFlexGrid1.FormatString = X
    lbl_sub_course_name.Caption = Empty
    lbl_month_year.Caption = Empty

    lbl_sub_course_name.Caption = cmb_sub.Text & "  Attendance " & Space(1) & cmb_course_name.Text & Space(1) & cmb_sem.Text & "(" & cmb_section.Text & ")"
        
           
                

            Call connection
            rs.Open "select max(att_month) from attendance where subjectid =(select id from subjects where name = '" & cmb_sub.Text & "' and section = '" & cmb_section & "') and semester = '" & cmb_sem & "' and att_year = '" & cmb_year.Text & "' and courseid =(select id from course where name = '" & cmb_course_name.Text & "') and section = '" & cmb_section & "'", conn, adOpenDynamic, adLockOptimistic
            If rs.EOF = False Then
                max_month = rs.Fields(0)
            End If
        
            lbl_month_year.Caption = "upto " & max_month & "/" & cmb_year.Text
          
        

               Call connection
                rs.Open "select regno,name from student where courseid= (select id from course where name = '" & cmb_course_name.Text & "') and semester = '" & cmb_sem & "' and section = '" & cmb_section & "'", conn, adOpenDynamic, adLockBatchOptimistic
                    i = 1
                        While (rs.EOF = False)
                                 
                               
                                    If attendance_rec.State = 1 Then attendance_rec.Close
                                         attendance_rec.Open "select sum(lec_delivered),sum(lec_attended) from attendance where subjectid =(select id from subjects where name = '" & cmb_sub.Text & "' and section = '" & cmb_section & "') and semester = '" & cmb_sem & "' and regno = '" & rs.Fields(0) & "' and att_month <= '" & max_month & "' and att_year = '" & cmb_year.Text & "' and courseid =(select id from course where name = '" & cmb_course_name.Text & "') and section = '" & cmb_section & "'", conn, adOpenDynamic, adLockOptimistic
                                    
                                            If Not (IsNull(attendance_rec.Fields(0))) Then
                                                    
                                                    X = CInt(attendance_rec(0))
                                                    Y = CInt(attendance_rec(1))
                                                    z = (Y / X) * 100
                    If z < 40 Then
                        MSFlexGrid1.TextMatrix(i, 0) = rs.Fields(0)
                        MSFlexGrid1.TextMatrix(i, 1) = rs.Fields(1)
                        MSFlexGrid1.TextMatrix(i, 2) = attendance_rec.Fields(0)
                        MSFlexGrid1.TextMatrix(i, 3) = attendance_rec.Fields(1)
                        If Len(z) > 4 Then
                            MSFlexGrid1.TextMatrix(i, 4) = Left(z, 4)
                        Else
                            MSFlexGrid1.TextMatrix(i, 4) = z
                        End If
                         i = i + 1
                    End If
                    End If
                   
                rs.MoveNext
                Wend

End Sub

Private Sub cmd_back_Click()
frm_entry_att.Show
    Unload Me

End Sub

Private Sub Form_Load()

Frame4.ForeColor = vbRed
Call connection
        rs.Open "select name from department", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_dept_name.Clear
            While (rs.EOF = False)
                cmb_dept_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend
End Sub

Public Function validation()
          For Each ctr In frm_detained_list.Controls
           If TypeOf ctr Is ComboBox Then
            If Trim(ctr.Text) = Empty Then
                ctr.BackColor = vbGreen
            Else
                ctr.BackColor = vbWhite
            End If
           End If
          Next
       For Each ctr In frm_detained_list.Controls
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


