VERSION 5.00
Begin VB.Form Frm_edit_teacher 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Teacher"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "UPDATE Details"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton cmd_back 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   615
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   4335
      End
      Begin VB.CommandButton cmd_delete 
         BackColor       =   &H80000004&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmd_update 
         BackColor       =   &H80000004&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txt_id 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txt_teacher_name 
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
         Left            =   1920
         TabIndex        =   1
         Text            =   "70"
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Department "
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Teacher Name"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "ID"
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
         Left            =   1320
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "Frm_edit_teacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_back_Click()
frm_view_teacher.Show
Unload Me
End Sub

Private Sub cmd_delete_Click()
warning = MsgBox("Are you sure?", vbYesNo + vbQuestion, "warning")
If warning = vbYes Then
Call connection
rs.Open "delete from teacher where id = " & Val(Trim(txt_id.Text)) & " ", conn, adOpenDynamic, adLockBatchOptimistic
MsgBox "Record deleted successfully", vbInformation
Unload Me
frm_view_teacher.Show
End If
End Sub

Private Sub cmd_update_Click()

Call connection
rs.Open "select id from department where name = '" & cmb_dept_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
d_id = rs.Fields(0)
Call connection
rs.Open "update teacher set deptid = " & d_id & " where id = " & Val(Trim(txt_id.Text)) & " ", conn, adOpenDynamic, adLockBatchOptimistic
MsgBox "Record Updated Successfully", vbInformation
Unload Me
frm_view_teacher.Show
End Sub

Private Sub Form_Load()
txt_id.Enabled = False
txt_teacher_name.Enabled = False
    Call connection
rs.Open "select * from department", conn, adOpenDynamic, adLockBatchOptimistic
cmb_dept_name.Clear
           While (rs.EOF = False)
                cmb_dept_name.AddItem (rs.Fields(1))
                rs.MoveNext
            Wend
End Sub

