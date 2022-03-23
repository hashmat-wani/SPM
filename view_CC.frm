VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_view_cc 
   Caption         =   "Class Co-ordinator List"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmb_dept_name 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9128
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Class Co-ordinator List"
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
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   8655
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frm_view_cc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim conn1 As New ADODB.connection
Private Sub cmb_dept_name_Click()
MSFlexGrid1.RowHeight(0) = 350
MSFlexGrid1.Appearance = flex3D
MSFlexGrid1.BackColorBkg = vbWhite
MSFlexGrid1.FillStyle = flexFillRepeat
MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 0
MSFlexGrid1.RowSel = 0
MSFlexGrid1.ColSel = 3
MSFlexGrid1.BackColorSel = &H80000014
MSFlexGrid1.ForeColorSel = &H80000013
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellFontName = "Broadway"
MSFlexGrid1.CellFontSize = 10
MSFlexGrid1.CellFontUnderline = True
MSFlexGrid1.CellTextStyle = flexTextInsetLight

MSFlexGrid1.ColWidth(0) = 4000
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1800
MSFlexGrid1.ColWidth(3) = 1500
MSFlexGrid1.TextMatrix(0, 0) = "Name"
MSFlexGrid1.TextMatrix(0, 1) = "Course"
MSFlexGrid1.TextMatrix(0, 2) = "Semester"
MSFlexGrid1.TextMatrix(0, 3) = "Section"
Call connection
rs.Open "select count(deptid) from class_coordinator where deptid =(select id from department where name = '" & cmb_dept_name & "') and teacherid <> null", conn, adOpenDynamic, adLockBatchOptimistic
Dim n As Integer
n = rs.Fields(0)
Call connection
rs.Open "select teacherid,courseid,semester,section from class_coordinator where deptid =(select id from department where name = '" & cmb_dept_name & "') and teacherid <> null", conn, adOpenDynamic, adLockBatchOptimistic
 MSFlexGrid1.Rows = n + 1
 For i = 1 To n
 Call connection1
        t_id = rs.Fields(0)
        rs2.Open "select name from teacher where id = " & t_id & "", conn1, adOpenDynamic, adLockBatchOptimistic
        MSFlexGrid1.TextMatrix(i, 0) = (rs2.Fields(0))
        rs2.MoveNext
 Call connection1
        c_id = rs.Fields(1)
        rs2.Open "select name from course where id = " & c_id & "", conn1, adOpenDynamic, adLockBatchOptimistic
        MSFlexGrid1.TextMatrix(i, 1) = (rs2.Fields(0))
        rs2.MoveNext
        Select Case rs.Fields(2)
        Case 1: sem = "I"
        Case 2: sem = "II"
        Case 3: sem = "III"
        Case 4: sem = "IV"
        Case 5: sem = "V"
        Case 6: sem = "VI"
        Case 7: sem = "VII"
        Case 8: sem = "VIII"
        End Select
 MSFlexGrid1.TextMatrix(i, 2) = sem
 MSFlexGrid1.TextMatrix(i, 3) = (rs.Fields(3))
 rs.MoveNext
 Next

End Sub

Private Sub Form_Load()
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
        frm_edit_cc.Show
        frm_edit_cc.cmb_dept_name = cmb_dept_name.Text
        frm_edit_cc.cmb_course_name = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
        frm_edit_cc.cmb_sem = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
        frm_edit_cc.cmb_sec = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
        frm_edit_cc.cmb_teacher_name = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    Unload Me
End If
End Sub
Public Function connection1()
If conn1.State = 1 Then conn1.Close
conn1.Open "DSN=Semester_Process_Management"
End Function

