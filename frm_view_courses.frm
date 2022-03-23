VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_view_courses 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "view courses"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9128
      _Version        =   393216
      Rows            =   1
      Cols            =   5
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
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Course List"
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
      TabIndex        =   1
      Top             =   720
      Width           =   6855
   End
End
Attribute VB_Name = "frm_view_courses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_dept_name_Click()
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
MSFlexGrid1.CellFontSize = 10
MSFlexGrid1.CellFontUnderline = True
MSFlexGrid1.CellTextStyle = flexTextInsetLight
MSFlexGrid1.ColWidth(0) = 1000
MSFlexGrid1.ColWidth(1) = 2100
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(3) = 1200
MSFlexGrid1.ColWidth(4) = 1500

MSFlexGrid1.TextMatrix(0, 0) = "ID"
MSFlexGrid1.TextMatrix(0, 1) = "Name"
MSFlexGrid1.TextMatrix(0, 2) = "Semsters"
MSFlexGrid1.TextMatrix(0, 3) = "Sections"
MSFlexGrid1.TextMatrix(0, 4) = "Shortage Cr."

Call connection
rs.Open "select count(id) from course where deptid =(select id from department where name = '" & cmb_dept_name & "')", conn, adOpenDynamic, adLockBatchOptimistic
Dim n As Integer
n = rs.Fields(0)
Call connection
rs.Open "select id,name,no_of_sems,no_of_secs,shortage_criteria from course where deptid =(select id from department where name = '" & cmb_dept_name & "')", conn, adOpenDynamic, adLockBatchOptimistic
 MSFlexGrid1.Rows = n + 1
 For i = 1 To n
 MSFlexGrid1.TextMatrix(i, 0) = (rs.Fields(0))
 MSFlexGrid1.TextMatrix(i, 1) = (rs.Fields(1))
 MSFlexGrid1.TextMatrix(i, 2) = (rs.Fields(2))
 MSFlexGrid1.TextMatrix(i, 3) = (rs.Fields(3))
 MSFlexGrid1.TextMatrix(i, 4) = (rs.Fields(4))
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
    frm_edit_course.Show
    frm_edit_course.txt_shortage.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
    frm_edit_course.txt_id.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    frm_edit_course.txt_course_name.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
    frm_edit_course.cmb_sems = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
    frm_edit_course.cmb_sections = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
    frm_edit_course.cmb_dept_name = cmb_dept_name.Text
    frmEditCourse_oldsem = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
    Unload Me
End If
End Sub
