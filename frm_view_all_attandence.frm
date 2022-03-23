VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_view_all_attandence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View all Attendance"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Select "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   8775
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb_sub_name 
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton cmd_back 
      Height          =   375
      Left            =   720
      Picture         =   "frm_view_all_attandence.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click to go previous page"
      Top             =   480
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6135
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10821
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1800
      Width           =   7455
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
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   9135
   End
End
Attribute VB_Name = "frm_view_all_attandence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim attendance_rec As New ADODB.Recordset
Dim course_criteria As Integer




Private Sub cmb_sec_click()
    cmb_sec.BackColor = vbWhite
    MSFlexGrid1.Clear
    Call connection
        rs.Open "select name from subjects where courseid=(select id from course where name = '" & frm_entry_att.cmb_course_name.Text & "') and semester = '" & frm_entry_att.cmb_sem & "' and section = '" & cmb_sec & "'", conn, adOpenDynamic, adLockBatchOptimistic
          cmb_sub_name.Clear
          While (rs.EOF = False)
            cmb_sub_name.AddItem (rs.Fields(0))
            rs.MoveNext
          Wend

    End Sub



Private Sub cmb_sub_name_Click()
MSFlexGrid1.Clear
Call connection
rs.Open "select shortage_criteria from course where name = '" & frm_entry_att.cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
course_criteria = rs.Fields(0)

Dim X As String

 
    X = "^ Reg.No. " & Space(5) & " |^  Name " & Space(40) & "|^Lect  deliverd " & Space(5) & "|^lect Attended " & Space(5) & "|^ %age" & Space(8)
    frm_view_all_attandence.MSFlexGrid1.FormatString = X
    frm_view_all_attandence.lbl_sub_course_name.Caption = Empty
    frm_view_all_attandence.lbl_month_year.Caption = Empty

    lbl_sub_course_name.Caption = cmb_sub_name.Text & "  Attendance " & Space(1) & frm_entry_att.cmb_course_name.Text & Space(1) & frm_entry_att.cmb_sem.Text & "(" & cmb_sec.Text & ")"
        
If frm_entry_att.opt_mnth.Value = True Then
            lbl_month_year.Caption = "for " & frm_entry_att.cmb_month.Text & Space(1) & frm_entry_att.cmb_year.Text
                 
                 Call connection
                 rs.Open "select regno,name from student where courseid= (select id from course where name = '" & frm_entry_att.cmb_course_name.Text & "') and semester = '" & frm_entry_att.cmb_sem & "' and section = '" & cmb_sec & "'", conn, adOpenDynamic, adLockBatchOptimistic
                    i = 1
                    While (rs.EOF = False)
                                MSFlexGrid1.TextMatrix(i, 0) = rs.Fields(0)
                                MSFlexGrid1.TextMatrix(i, 1) = rs.Fields(1)
                                If attendance_rec.State = 1 Then attendance_rec.Close
                                    attendance_rec.Open "select lec_delivered,lec_attended from attendance where subjectid =(select id from subjects where name = '" & cmb_sub_name.Text & "' and section = '" & cmb_sec & "') and semester = '" & frm_entry_att.cmb_sem & "' and regno = '" & rs.Fields(0) & "' and att_month = '" & month_no & "' and att_year = '" & frm_entry_att.cmb_year.Text & "' and courseid =(select id from course where name = '" & frm_entry_att.cmb_course_name.Text & "') and section = '" & cmb_sec & "'", conn, adOpenDynamic, adLockOptimistic
                                        If attendance_rec.EOF = False Then
                   
                                             MSFlexGrid1.TextMatrix(i, 2) = attendance_rec.Fields(0)
                                                MSFlexGrid1.TextMatrix(i, 3) = attendance_rec.Fields(1)
                                                X = CInt(attendance_rec(0))
                                                Y = CInt(attendance_rec(1))
                                                z = (Y / X) * 100
                        
                                                If Len(z) > 4 Then
                                                    MSFlexGrid1.TextMatrix(i, 4) = Left(z, 4)
                                                Else
                                                    MSFlexGrid1.TextMatrix(i, 4) = z
                                                End If
                      If z < course_criteria Then
                       frm_view_all_attandence.MSFlexGrid1.Row = i
                       For j = 0 To 4
                       frm_view_all_attandence.MSFlexGrid1.Col = j
                       frm_view_all_attandence.MSFlexGrid1.CellForeColor = vbRed
                       frm_view_all_attandence.MSFlexGrid1.CellBackColor = &HE0E0E0
                        Next
                         End If
                         
                            
                                        Else
                                            MSFlexGrid1.TextMatrix(i, 2) = 0
                                            MSFlexGrid1.TextMatrix(i, 3) = 0
                                            MSFlexGrid1.TextMatrix(i, 4) = 0
                                        End If
                        
                        i = i + 1
                rs.MoveNext
                Wend
                    
Else
            Call connection
            rs.Open "select max(att_month) from attendance where subjectid =(select id from subjects where name = '" & cmb_sub_name.Text & "' and section = '" & cmb_sec & "') and semester = '" & frm_entry_att.cmb_sem & "' and att_year = '" & frm_entry_att.cmb_year.Text & "' and courseid =(select id from course where name = '" & frm_entry_att.cmb_course_name.Text & "') and section = '" & cmb_sec & "'", conn, adOpenDynamic, adLockOptimistic
            If rs.EOF = False Then
                max_month = rs.Fields(0)
            End If
        
            lbl_month_year.Caption = "upto " & max_month & "/" & frm_entry_att.cmb_year.Text
          
        

                Call connection
                rs.Open "select regno,name from student where courseid= (select id from course where name = '" & frm_entry_att.cmb_course_name.Text & "') and semester = '" & frm_entry_att.cmb_sem & "' and section = '" & cmb_sec & "'", conn, adOpenDynamic, adLockBatchOptimistic
                    i = 1
                        While (rs.EOF = False)
                                 MSFlexGrid1.TextMatrix(i, 0) = rs.Fields(0)
                                MSFlexGrid1.TextMatrix(i, 1) = rs.Fields(1)
                               
                                    If attendance_rec.State = 1 Then attendance_rec.Close
                                         
                                         attendance_rec.Open "select sum(lec_delivered),sum(lec_attended) from attendance where subjectid =(select id from subjects where name = '" & cmb_sub_name.Text & "' and section = '" & cmb_sec & "') and semester = '" & frm_entry_att.cmb_sem & "' and regno = '" & rs.Fields(0) & "' and att_month <= '" & max_month & "' and att_year = '" & frm_entry_att.cmb_year.Text & "' and courseid =(select id from course where name = '" & frm_entry_att.cmb_course_name.Text & "') and section = '" & cmb_sec & "'", conn, adOpenDynamic, adLockOptimistic
                                    
                                             If Not (IsNull(attendance_rec.Fields(0))) Then
                                                    MSFlexGrid1.TextMatrix(i, 2) = attendance_rec.Fields(0)
                                                    MSFlexGrid1.TextMatrix(i, 3) = attendance_rec.Fields(1)
                                                    X = CInt(attendance_rec(0))
                                                    Y = CInt(attendance_rec(1))
                                                    z = (Y / X) * 100
                        
                                                    If Len(z) > 4 Then
                                                        MSFlexGrid1.TextMatrix(i, 4) = Left(z, 4)
                                                    Else
                                                        MSFlexGrid1.TextMatrix(i, 4) = z
                                                    End If
                         If z < course_criteria Then
                       frm_view_all_attandence.MSFlexGrid1.Row = i
                       For j = 0 To 4
                       frm_view_all_attandence.MSFlexGrid1.Col = j
                       frm_view_all_attandence.MSFlexGrid1.CellForeColor = vbRed
                       frm_view_all_attandence.MSFlexGrid1.CellBackColor = &HE0E0E0
                        Next
                         End If
                                                                        
                                            Else
                                                MSFlexGrid1.TextMatrix(i, 2) = 0
                                                MSFlexGrid1.TextMatrix(i, 3) = 0
                                                MSFlexGrid1.TextMatrix(i, 4) = 0
                                            End If
                                i = i + 1
                rs.MoveNext
                Wend
       
        End If
End Sub

Private Sub cmd_back_Click()
frm_entry_att.Show
Unload Me
End Sub

Private Sub Form_Load()
lbl_sub_course_name.Caption = Empty
lbl_month_year.Caption = Empty
End Sub

