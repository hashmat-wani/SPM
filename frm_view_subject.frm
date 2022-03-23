VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_view_subject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Subjects"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   8355
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   3375
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9763
      _Version        =   393216
      Rows            =   1
      Cols            =   4
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
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   975
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
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Subject List"
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
      Height          =   2055
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   7935
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
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frm_view_subject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim conn1 As New ADODB.connection
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
MSFlexGrid1.ColSel = 3
MSFlexGrid1.BackColorSel = &H80000014
MSFlexGrid1.ForeColorSel = &H80000013
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellFontName = "Broadway"
MSFlexGrid1.CellFontSize = 10
MSFlexGrid1.CellFontUnderline = True
MSFlexGrid1.CellTextStyle = flexTextInsetLight
MSFlexGrid1.ColWidth(0) = 1000
MSFlexGrid1.ColWidth(1) = 3000
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 2600
MSFlexGrid1.TextMatrix(0, 0) = "ID"
MSFlexGrid1.TextMatrix(0, 1) = "Name"
MSFlexGrid1.TextMatrix(0, 2) = "Section"
MSFlexGrid1.TextMatrix(0, 3) = "Teacher"
Call connection
rs.Open "select id from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
c_id = rs.Fields(0)
Call connection
rs.Open "select count(id) from subjects where semester = '" & cmb_sem & "' and courseid = " & c_id & "", conn, adOpenDynamic, adLockBatchOptimistic
Dim n As Integer
n = rs.Fields(0)
Call connection
rs.Open "select id,name,section,teacherid from subjects where semester = '" & cmb_sem & "' and courseid = " & c_id & "", conn, adOpenDynamic, adLockBatchOptimistic
 MSFlexGrid1.Rows = n + 1
 For i = 1 To n
 MSFlexGrid1.TextMatrix(i, 0) = (rs.Fields(0))
 MSFlexGrid1.TextMatrix(i, 1) = (rs.Fields(1))
 MSFlexGrid1.TextMatrix(i, 2) = (rs.Fields(2))
    If rs.Fields(3) <> Empty Then
        Call connection1
        t_id = rs.Fields(3)
        rs2.Open "select name from teacher where id = " & t_id & "", conn1, adOpenDynamic, adLockBatchOptimistic
        MSFlexGrid1.TextMatrix(i, 3) = (rs2.Fields(0))
        rs2.MoveNext
 End If
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

Public Function connection1()
If conn1.State = 1 Then conn1.Close
conn1.Open "DSN=Semester_Process_Management"
End Function

Private Sub MSFlexGrid1_EnterCell()
If MSFlexGrid1.Row <> 0 Then
    If MSFlexGrid1.Col = 3 Then
        frm_edit_sub_allotment.Show
        frm_edit_sub_allotment.txt_id.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
        If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) <> "" Then
            frm_edit_sub_allotment.cmb_teacher_name = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
        End If
    Else
        frm_edit_sub.Show
        frm_edit_sub.txt_id.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
        frm_edit_sub.txt_sub_name.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
        frm_edit_sub.cmb_course_name = cmb_course_name.Text
        frm_edit_sub.cmb_sems = cmb_sem.Text
        
    End If
    Unload Me
End If
End Sub
