VERSION 5.00
Begin VB.Form frm_edit_department 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Details"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6375
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
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
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
         TabIndex        =   7
         Top             =   360
         Width           =   615
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1560
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
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
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
         Left            =   2160
         TabIndex        =   1
         Text            =   "70"
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Department Name"
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
         TabIndex        =   6
         Top             =   960
         Width           =   1935
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
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frm_edit_department"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_back_Click()
frm_view_dept.Show
Unload Me
End Sub

Private Sub cmd_delete_Click()
warning = MsgBox("Are you sure?", vbYesNo + vbQuestion, "warning")
If warning = vbYes Then
Call connection
rs.Open "delete from department where id = " & Val(Trim(txt_id.Text)) & " ", conn, adOpenDynamic, adLockBatchOptimistic
MsgBox "Record deleted successfully", vbInformation
End If
Unload Me
frm_view_dept.Show
End Sub

Private Sub cmd_update_Click()
If txt_dept_name = "" Then
    MsgBox "Empty Field not allowed", vbInformation
    txt_dept_name.BackColor = vbGreen
Else
    Call connection
    rs.Open "select count(id) from department where name = '" & Trim(txt_dept_name.Text) & "'", conn, adOpenDynamic, adLockBatchOptimistic
    n = rs.Fields(0)
    If n > 0 Then
        MsgBox "Record already exists", vbCritical
        txt_dept_name.ForeColor = vbRed
    Else
        Call connection
        rs.Open "update department set name = '" & Trim(txt_dept_name.Text) & "'where id = " & Val(Trim(txt_id.Text)) & " ", conn, adOpenDynamic, adLockBatchOptimistic
        MsgBox "Record Updated Successfully", vbInformation
        Unload Me
        frm_view_dept.Show
    End If
End If
End Sub

Private Sub Form_Load()
txt_id.Enabled = False
End Sub

Private Sub txt_dept_name_Change()
txt_dept_name.ForeColor = vbBlack

End Sub

Private Sub txt_dept_name_Click()
txt_dept_name.BackColor = vbWhite
End Sub
