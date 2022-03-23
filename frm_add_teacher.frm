VERSION 5.00
Begin VB.Form frm_add_teacher 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teachers"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Teacher Details"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.CommandButton cmd_save 
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmd_rst 
         BackColor       =   &H80000004&
         Caption         =   "Reset"
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txt_email 
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
         Left            =   1680
         TabIndex        =   6
         Text            =   "txt_email"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txt_name 
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
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   4215
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
         TabIndex        =   1
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Email"
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
         Left            =   720
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
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
         Left            =   720
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Department"
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
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_add_teacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_dept_name_Click()
cmb_dept_name.BackColor = vbWhite
End Sub

Private Sub cmd_rst_Click()
Call Form_Load
End Sub

Private Sub cmd_save_Click()
If validation = False Then
    Exit Sub
End If
Dim n, m_id, d_id As Integer
Call connection
rs.Open "select count(id) from teacher where email = '" & Trim(txt_email.Text) & "'", conn, adOpenDynamic, adLockBatchOptimistic
n = rs.Fields(0)
If n > 0 Then
    MsgBox "Email already exists", vbCritical
    txt_email.ForeColor = vbRed
Else
    Call connection
    rs.Open "select id from department where name = '" & cmb_dept_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
    d_id = rs.Fields(0)
    Call connection
    rs.Open "select max(id) from teacher", conn, adOpenDynamic, adLockBatchOptimistic
    If IsNull(rs.Fields(0)) Then
        Call connection
        rs.Open "insert into teacher(id,name,deptid,email,password,username) values(1,'" & Trim(txt_name.Text) & "'," & d_id & ",'" & Trim(txt_email.Text) & "','" & Trim(txt_name.Text) & "','" & Trim(txt_email.Text) & "')", conn, adOpenDynamic, adLockBatchOptimistic
    Else
        m_id = rs.Fields(0)
        Call connection
        rs.Open "insert into teacher(id,name,deptid,email,password,username) values(" & m_id + 1 & " ,'" & Trim(txt_name.Text) & "'," & d_id & ",'" & Trim(txt_email.Text) & "','" & Trim(txt_name.Text) & "','" & Trim(txt_email.Text) & "')", conn, adOpenDynamic, adLockBatchOptimistic
    End If
    MsgBox "Record Saved successfully", vbInformation
    Call Form_Load
End If
End Sub

Private Sub Form_Load()
cmb_dept_name.Clear
txt_name.Text = ""
txt_email.Text = ""
For Each ctr In frm_add_teacher.Controls
If TypeOf ctr Is ComboBox Or TypeOf ctr Is TextBox Then
ctr.BackColor = vbWhite
End If
Next
Call connection
rs.Open "select * from department", conn, adOpenDynamic, adLockBatchOptimistic
cmb_dept_name.Clear
           While (rs.EOF = False)
                cmb_dept_name.AddItem (rs.Fields(1))
                rs.MoveNext
            Wend
End Sub

Public Function validation()
          For Each ctr In frm_add_teacher.Controls
           If TypeOf ctr Is ComboBox Or TypeOf ctr Is TextBox Then
            If Trim(ctr.Text) = Empty Then
                ctr.BackColor = vbGreen
            Else
                ctr.BackColor = vbWhite
            End If
           End If
          Next
       For Each ctr In frm_add_teacher.Controls
            If TypeOf ctr Is TextBox Or TypeOf ctr Is ComboBox Then
                If ctr.BackColor = vbGreen Then
                    MsgBox "Fill all Fields", vbInformation
                    validation = False
                    Exit Function
                End If
            End If
        Next
validation = True
End Function

Private Sub txt_email_Click()
txt_email.BackColor = vbWhite
End Sub
Private Sub txt_email_Change()
txt_email.ForeColor = vbBlack
End Sub

Private Sub txt_name_Click()
txt_name.BackColor = vbWhite
End Sub
