VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_view_students 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Students List"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
   Begin VB.ComboBox cmb_Sect 
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ComboBox cmb_batch 
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9763
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColorSel    =   -2147483628
      ForeColorSel    =   -2147483629
      FocusRect       =   2
      FillStyle       =   1
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
   Begin VB.Label Label8 
      Caption         =   "Semester"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Course"
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
      Left            =   390
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Section"
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
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Batch"
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
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Students List/Section"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   7935
   End
End
Attribute VB_Name = "frm_view_students"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmb_batch_Click()
MSFlexGrid1.Clear
End Sub

Private Sub cmb_course_name_Click()
MSFlexGrid1.Clear
cmb_sem.Clear
cmb_Sect.Clear
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

Private Sub cmb_Sect_Click()
MSFlexGrid1.Clear
MSFlexGrid1.RowHeight(0) = 350
MSFlexGrid1.Appearance = flex3D
MSFlexGrid1.BackColorBkg = vbWhite
MSFlexGrid1.FillStyle = flexFillRepeat
MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 0
MSFlexGrid1.RowSel = 0
MSFlexGrid1.ColSel = 4
MSFlexGrid1.BackColorSel = &H80000014
MSFlexGrid1.ForeColorSel = &H80000013
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellFontName = "Broadway"
MSFlexGrid1.CellFontSize = 8
MSFlexGrid1.CellFontUnderline = True
MSFlexGrid1.CellTextStyle = flexTextInsetLight
MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 3000
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 2500
MSFlexGrid1.ColWidth(4) = 8000

MSFlexGrid1.TextMatrix(0, 0) = "Reg.No."
MSFlexGrid1.TextMatrix(0, 1) = "Name"
MSFlexGrid1.TextMatrix(0, 2) = "Contact"
MSFlexGrid1.TextMatrix(0, 3) = "Email"
MSFlexGrid1.TextMatrix(0, 4) = "Address"

Call connection
rs.Open "select id from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
c_id = rs.Fields(0)
Call connection
rs.Open "select count(regno) from student where batch = '" & cmb_batch & "' and semester = '" & cmb_sem & "' and courseid = " & c_id & " and section = '" & cmb_Sect & "'", conn, adOpenDynamic, adLockBatchOptimistic
Dim n As Integer
n = rs.Fields(0)
Call connection
rs.Open "select regno,name,contact,email,address from student where batch = '" & cmb_batch & "' and semester = '" & cmb_sem & "' and courseid = " & c_id & " and section = '" & cmb_Sect & "'", conn, adOpenDynamic, adLockBatchOptimistic
 MSFlexGrid1.Rows = n + 1
 For i = 1 To n
 MSFlexGrid1.TextMatrix(i, 0) = (rs.Fields(0))
 MSFlexGrid1.TextMatrix(i, 1) = (rs.Fields(1))
 If rs.Fields(2) <> Empty Then
    MSFlexGrid1.TextMatrix(i, 2) = (rs.Fields(2))
 End If
 If rs.Fields(3) <> Empty Then
    MSFlexGrid1.TextMatrix(i, 3) = rs.Fields(3)
 End If
 If rs.Fields(4) <> Empty Then
    MSFlexGrid1.TextMatrix(i, 4) = rs.Fields(4)
 End If
 rs.MoveNext
 Next




End Sub

Private Sub cmb_sem_Click()
cmb_Sect.Clear
Call connection
rs.Open "select no_of_secs from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
For i = 1 To rs.Fields(0)
    Select Case i
        Case 1: cmb_Sect.AddItem "A"
        Case 2: cmb_Sect.AddItem "B"
        Case 3: cmb_Sect.AddItem "C"
        Case 4: cmb_Sect.AddItem "D"
        Case 5: cmb_Sect.AddItem "E"
    End Select
Next

End Sub

Private Sub Form_Load()
cmb_batch.Clear
For i = 1 To 50
cmb_batch.AddItem 1999 + i
Next
Call connection
rs.Open "select name from course", conn, adOpenDynamic, adLockBatchOptimistic
cmb_course_name.Clear
           While (rs.EOF = False)
                cmb_course_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend

End Sub
Private Sub MSFlexGrid1_EnterCell()
If MSFlexGrid1.Row <> 0 Then
        frm_edit_student.Show
        frm_edit_student.txt_reg_no.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
        frm_edit_student.txt_name.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
        frm_edit_student.txt_contact.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
        frm_edit_student.txt_email.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
        frm_edit_student.txt_address.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
        frm_edit_student.cmb_course_name = cmb_course_name.Text
        frm_edit_student.cmb_batch = cmb_batch.Text
        frm_edit_student.cmb_course_name = cmb_course_name.Text
        frm_edit_student.cmb_sem = cmb_sem.Text
        frm_edit_student.cmb_Sect = cmb_Sect.Text
        
    End If
    Unload Me
End Sub

