VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_get_regno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get Reg.No"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   11033
      _Version        =   393216
      Rows            =   1
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
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frm_get_regno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
MSFlexGrid1.Clear
MSFlexGrid1.RowHeight(0) = 350
MSFlexGrid1.Appearance = flex3D
MSFlexGrid1.BackColorBkg = vbWhite
MSFlexGrid1.FillStyle = flexFillRepeat
MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 0
MSFlexGrid1.RowSel = 0
MSFlexGrid1.ColSel = 1
MSFlexGrid1.BackColorSel = &H80000014
MSFlexGrid1.ForeColorSel = &H80000013
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellFontName = "Broadway"
MSFlexGrid1.CellFontSize = 8
MSFlexGrid1.CellFontUnderline = True
MSFlexGrid1.CellTextStyle = flexTextInsetLight
MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 4000
MSFlexGrid1.TextMatrix(0, 0) = "Reg.No."
MSFlexGrid1.TextMatrix(0, 1) = "Name"
Call connection
rs.Open "select id from course where name = '" & frm_add_CR.cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
c_id = rs.Fields(0)
Call connection
rs.Open "select count(regno) from student where batch = '" & frm_add_CR.cmb_batch & "' and semester = '" & frm_add_CR.cmb_sem & "' and courseid = " & c_id & " and section = '" & frm_add_CR.cmb_sec & "'", conn, adOpenDynamic, adLockBatchOptimistic
Dim n As Integer
n = rs.Fields(0)
Call connection
rs.Open "select regno,name from student where batch = '" & frm_add_CR.cmb_batch & "' and semester = '" & frm_add_CR.cmb_sem & "' and courseid = " & c_id & " and section = '" & frm_add_CR.cmb_sec & "'", conn, adOpenDynamic, adLockBatchOptimistic
 MSFlexGrid1.Rows = n + 1
 For i = 1 To n
 MSFlexGrid1.TextMatrix(i, 0) = (rs.Fields(0))
 MSFlexGrid1.TextMatrix(i, 1) = (rs.Fields(1))
 rs.MoveNext
 Next
End Sub

Private Sub MSFlexGrid1_EnterCell()
If MSFlexGrid1.Row <> 0 Then
        frm_add_CR.Show
        frm_add_CR.txt_regno.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    End If
    Unload Me
End Sub

