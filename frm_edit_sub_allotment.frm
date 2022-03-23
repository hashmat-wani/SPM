VERSION 5.00
Begin VB.Form frm_edit_sub_allotment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teacher Allotment"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Teacher Allotment"
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
      Width           =   5415
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
         TabIndex        =   8
         Top             =   360
         Width           =   615
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
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   1575
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
         TabIndex        =   5
         Top             =   2040
         Width           =   975
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
         TabIndex        =   2
         Text            =   "cmb_dept_name"
         Top             =   960
         Width           =   3615
      End
      Begin VB.ComboBox cmb_teacher_name 
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
         TabIndex        =   1
         Top             =   1440
         Width           =   3615
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
         Left            =   1080
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label6 
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
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Teacher"
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
         Left            =   480
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm_edit_sub_allotment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_dept_name_Click()
Call connection
rs.Open "select name from teacher where deptid =(select id from department where name = '" & cmb_dept_name & "')", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_teacher_name.Clear
           While (rs.EOF = False)
                cmb_teacher_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend

End Sub

Private Sub cmb_teacher_name_Click()
Call connection
rs.Open "select deptid from teacher where name = '" & cmb_teacher_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
d_id = rs.Fields(0)
Call connection
rs.Open "select name from department where id = " & d_id & "", conn, adOpenDynamic, adLockBatchOptimistic
cmb_dept_name.Text = rs.Fields(0)

End Sub

Private Sub cmd_back_Click()
frm_view_subject.Show
Unload Me
End Sub

Private Sub cmd_update_Click()
Call connection
rs.Open "select id from teacher where name = '" & cmb_teacher_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
t_id = rs.Fields(0)
Call connection
rs.Open "update subjects set teacherid = " & t_id & " where id = " & Val(Trim(txt_id.Text)) & " ", conn, adOpenDynamic, adLockBatchOptimistic
MsgBox "Record Updated Successfully", vbInformation
Unload Me
frm_view_subject.Show
End Sub

Private Sub Form_Load()
Call connection
 rs.Open "select name from department", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_dept_name.Clear
           While (rs.EOF = False)
                cmb_dept_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend
Call connection
rs.Open "select name from teacher", conn, adOpenDynamic, adLockBatchOptimistic
        cmb_teacher_name.Clear
           While (rs.EOF = False)
                cmb_teacher_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend

End Sub
