VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_view_cr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CR List"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7800
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
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
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
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   4095
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
      Top             =   1320
      Width           =   4095
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
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
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
      Left            =   600
      TabIndex        =   9
      Top             =   1800
      Width           =   975
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
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      TabIndex        =   7
      Top             =   1320
      Width           =   735
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
      Left            =   800
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Class Representative List"
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
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   7335
   End
End
Attribute VB_Name = "frm_view_cr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim conn1 As New ADODB.connection

Private Sub cmb_batch_Click()
MSFlexGrid1.Clear

End Sub

Private Sub cmb_course_name_Click()
MSFlexGrid1.Clear
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
MSFlexGrid1.Clear
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

Private Sub cmb_sem_Click()
MSFlexGrid1.Clear
MSFlexGrid1.RowHeight(0) = 350
MSFlexGrid1.Appearance = flex3D
MSFlexGrid1.BackColorBkg = vbWhite
MSFlexGrid1.FillStyle = flexFillRepeat
MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 0
MSFlexGrid1.RowSel = 0
MSFlexGrid1.ColSel = 2
MSFlexGrid1.BackColorSel = &H80000014
MSFlexGrid1.ForeColorSel = &H80000013
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellFontName = "Broadway"
MSFlexGrid1.CellFontSize = 10
MSFlexGrid1.CellFontUnderline = True
MSFlexGrid1.CellTextStyle = flexTextInsetLight

MSFlexGrid1.ColWidth(0) = 2000
MSFlexGrid1.ColWidth(1) = 3800
MSFlexGrid1.ColWidth(2) = 1900
MSFlexGrid1.TextMatrix(0, 0) = "Reg.No."
MSFlexGrid1.TextMatrix(0, 1) = "Name"
MSFlexGrid1.TextMatrix(0, 2) = "Section"
Call connection
rs.Open "select count(regno) from CR where deptid =(select id from department where name = '" & cmb_dept_name & "') and courseid =(select id from course where name = '" & cmb_course_name & "') and semester = '" & cmb_sem & "'", conn, adOpenDynamic, adLockBatchOptimistic
Dim n As Integer
n = rs.Fields(0)
Call connection
rs.Open "select regno,section from CR where deptid =(select id from department where name = '" & cmb_dept_name & "') and courseid =(select id from course where name = '" & cmb_course_name & "') and semester = '" & cmb_sem & "'", conn, adOpenDynamic, adLockBatchOptimistic
 MSFlexGrid1.Rows = n + 1
 For i = 1 To n
  MSFlexGrid1.TextMatrix(i, 0) = rs.Fields(0)
  MSFlexGrid1.TextMatrix(i, 2) = rs.Fields(1)
  Call connection1
        rs2.Open "select name from student where regno = '" & rs.Fields(0) & "'", conn1, adOpenDynamic, adLockBatchOptimistic
        MSFlexGrid1.TextMatrix(i, 1) = (rs2.Fields(0))
        rs2.MoveNext
 rs.MoveNext
 Next

End Sub

Private Sub Form_Load()
cmb_dept_name.Clear
cmb_course_name.Clear
cmb_sem.Clear
cmb_batch.Clear
For i = 1 To 50
cmb_batch.AddItem 1999 + i
Next

Call connection
rs.Open "select * from department", conn, adOpenDynamic, adLockBatchOptimistic
cmb_dept_name.Clear
           While (rs.EOF = False)
                cmb_dept_name.AddItem (rs.Fields(1))
                rs.MoveNext
            Wend
End Sub

Private Sub MSFlexGrid1_EnterCell()
If MSFlexGrid1.Row <> 0 Then
        frm_edit_cr.Show
        frm_edit_cr.cmb_batch = cmb_batch.Text
        frm_edit_cr.cmb_dept_name = cmb_dept_name.Text
        frm_edit_cr.cmb_course_name = cmb_course_name.Text
        frm_edit_cr.cmb_sem = cmb_sem.Text
        frm_edit_cr.cmb_sec = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
        frm_edit_cr.txt_regno.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    Unload Me
End If
End Sub


Public Function connection1()
If conn1.State = 1 Then conn1.Close
conn1.Open "DSN=Semester_Process_Management"
End Function

