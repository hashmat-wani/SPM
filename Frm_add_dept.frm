VERSION 5.00
Begin VB.Form Frm_add_dept 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Departments"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_save 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Department name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txt_dept_name 
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
         TabIndex        =   1
         Top             =   360
         Width           =   4695
      End
   End
End
Attribute VB_Name = "Frm_add_dept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_save_Click()
Dim n, m_id As Integer
If Trim(txt_dept_name.Text) = "" Then
    txt_dept_name.BackColor = vbGreen
    MsgBox "Empty Field not allowed", vbCritical
Else
    Call connection
    rs.Open "select count(id) from department where name = '" & Trim(txt_dept_name.Text) & "'", conn, adOpenDynamic, adLockBatchOptimistic
    n = rs.Fields(0)
    If n > 0 Then
        MsgBox "Record already exists", vbCritical
        txt_dept_name.ForeColor = vbRed
    Else
        Call connection
        rs.Open "select max(id) from department", conn, adOpenDynamic, adLockBatchOptimistic
        If IsNull(rs.Fields(0)) Then
        Call connection
            rs.Open "insert into department values(1,'" & Trim(txt_dept_name.Text) & "')", conn, adOpenDynamic, adLockBatchOptimistic
        Else
            m_id = rs.Fields(0)
            Call connection
            rs.Open "insert into department values(" & m_id + 1 & ",'" & Trim(txt_dept_name.Text) & "')", conn, adOpenDynamic, adLockBatchOptimistic
        End If
        MsgBox "Record Added successfully", vbInformation
        Call Form_Load
        txt_dept_name.SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
txt_dept_name.Text = ""
End Sub

Private Sub txt_dept_name_Click()
txt_dept_name.BackColor = vbWhite
End Sub

Private Sub txt_dept_name_Change()
txt_dept_name.ForeColor = vbBlack
End Sub
