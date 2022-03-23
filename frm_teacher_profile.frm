VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_teacher_profile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teacher Profile"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Set Username"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6480
      TabIndex        =   28
      Top             =   7680
      Width           =   5895
      Begin VB.CommandButton cmd_set_username 
         BackColor       =   &H80000004&
         Caption         =   "Set Username"
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txt_username 
         Height          =   405
         Left            =   2640
         TabIndex        =   29
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label10 
         Caption         =   "Username"
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
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6480
      TabIndex        =   20
      Top             =   4680
      Width           =   5895
      Begin VB.TextBox txt_new_pw_again 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txt_cur_pw 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txt_new_pw 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton cmd_change_password 
         BackColor       =   &H80000004&
         Caption         =   "Change Password"
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "New Password again"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Current Password"
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
         Left            =   600
         TabIndex        =   26
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "New Password"
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
         Left            =   960
         TabIndex        =   25
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Action"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6480
      TabIndex        =   16
      Top             =   480
      Width           =   5895
      Begin VB.CommandButton cmd_username 
         BackColor       =   &H80000004&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   1635
      End
      Begin VB.CommandButton cmd_password 
         BackColor       =   &H80000004&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2110
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1635
      End
      Begin VB.CommandButton cmd_edit_profile 
         BackColor       =   &H80000004&
         Caption         =   "Edit Profile"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   1635
      End
   End
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
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Edit Profile"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6480
      TabIndex        =   5
      Top             =   1680
      Width           =   5895
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
         Height          =   450
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txt_email 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   9
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txt_name 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txt_contact 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1560
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
         Left            =   360
         TabIndex        =   13
         Top             =   1080
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
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Contact"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   1335
      End
   End
   Begin VB.Frame frame_details 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6255
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
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1575
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
         Top             =   1080
         Width           =   2895
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4095
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7223
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2295
         Left            =   120
         TabIndex        =   33
         Top             =   6480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   1
         Cols            =   3
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
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "class co-ordinator"
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
         Left            =   1680
         TabIndex        =   32
         Top             =   6120
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Subjects Alloted"
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
         Left            =   1920
         TabIndex        =   15
         Top             =   1560
         Width           =   2415
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
         Left            =   1200
         TabIndex        =   4
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_teacher_profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim conn1 As New ADODB.connection
Dim o_email, n_email As String



Private Sub cmd_back_Click()
frm_teacher_login.Show
Unload Me

End Sub

Private Sub cmd_change_password_Click()
If txt_cur_pw.Text = "" Or txt_new_pw.Text = "" Or txt_new_pw_again = "" Then
     MsgBox "Empty Fields not allowed", vbInformation
     Else
    Call connection
    rs.Open "select password from teacher where id = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
    If Trim(txt_cur_pw.Text) = rs.Fields(0) Then
        If Trim(txt_new_pw.Text) = Trim(txt_new_pw_again.Text) Then
            Call connection
            rs.Open "update teacher set password = '" & Trim(txt_new_pw.Text) & "' where id = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
            MsgBox "Password changed syccessfully", vbInformation
            Form_Load
        Else
            MsgBox "Passwords do not match", vbInformation
        End If
    Else
        MsgBox " Your old password was entered incorrectly. Please try it again", vbInformation
        txt_cur_pw.SetFocus
    End If
End If
End Sub

Private Sub cmd_edit_profile_Click()
Call Form_Load
Frame1.Enabled = True
Frame1.ForeColor = vbRed
txt_name.SetFocus
End Sub

Private Sub cmd_username_Click()
Call Form_Load
Frame4.Enabled = True
Frame4.ForeColor = vbRed
Call connection
 rs.Open "select username from teacher where id = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
If rs.Fields(0) <> Empty Then
txt_username.Text = rs.Fields(0)
txt_username.ForeColor = vbRed
Else
txt_username.SetFocus
End If
End Sub

Private Sub cmd_password_Click()
Call Form_Load
Frame2.Enabled = True
Frame2.ForeColor = vbRed
txt_cur_pw.SetFocus
    
End Sub

Private Sub cmd_set_username_Click()
If txt_username.Text = "" Then
    MsgBox "Username can't be empty", vbInformation
    txt_username.BackColor = vbGreen
Else
    Call connection
    rs.Open "select count(id) from teacher where username = '" & Trim(txt_username.Text) & "'", conn, adOpenDynamic, adLockBatchOptimistic
    n = rs.Fields(0)
    If n > 0 Then
        MsgBox "Username already exists", vbCritical
        txt_username.ForeColor = vbRed
    Else
        Call connection
        rs.Open "update teacher set username = '" & Trim(txt_username.Text) & "' where id = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
        MsgBox "Record Saved successfully", vbInformation
        Call Form_Load
    End If
End If
End Sub

Private Sub cmd_update_Click()
n_email = txt_email.Text
If txt_name.Text = "" And txt_email.Text = "" Then
     MsgBox "Empty Fields not allowed", vbInformation
    txt_name.BackColor = vbGreen
    txt_email.BackColor = vbGreen
ElseIf txt_name.Text = "" Then
    MsgBox "Name can't be blank", vbInformation
    txt_name.BackColor = vbGreen
ElseIf txt_email.Text = "" Then
MsgBox "Email can't be blank", vbInformation
txt_email.BackColor = vbGreen
ElseIf o_email = n_email Then
    Call connection
        rs.Open "update teacher set name = '" & Trim(txt_name.Text) & "', contact = '" & Trim(txt_contact.Text) & "' where id = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
        MsgBox "Record Saved successfully", vbInformation
        Call Form_Load
Else
    Call connection
    rs.Open "select count(id) from teacher where email = '" & Trim(txt_email.Text) & "'", conn, adOpenDynamic, adLockBatchOptimistic
    n = rs.Fields(0)
    If n > 0 Then
        MsgBox "Email already exists", vbCritical
        txt_email.ForeColor = vbRed
    Else
        Call connection
        rs.Open "update teacher set name = '" & Trim(txt_name.Text) & "', email = '" & Trim(txt_email.Text) & "',contact = '" & Trim(txt_contact.Text) & "' where id = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
        MsgBox "Record Saved successfully", vbInformation
        Call Form_Load
    End If
End If

End Sub


Private Sub Form_Load()
Frame3.ForeColor = vbRed
txt_id.Enabled = False
cmb_dept_name.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False
Frame4.Enabled = False
txt_id.Text = teacher_id
cmb_dept_name.Clear
txt_contact.Text = ""
txt_email.Text = ""
txt_name.Text = ""
txt_username.Text = ""
txt_cur_pw.Text = ""
txt_new_pw.Text = ""
txt_new_pw_again.Text = ""
Call connection
rs.Open "select * from department", conn, adOpenDynamic, adLockBatchOptimistic
cmb_dept_name.Clear
           While (rs.EOF = False)
                cmb_dept_name.AddItem (rs.Fields(1))
                rs.MoveNext
            Wend
            Call connection
            rs.Open "select name,email,contact,deptid from teacher where id = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
    txt_name.Text = rs.Fields(0)
    txt_email.Text = rs.Fields(1)
    If rs.Fields(2) <> Empty Then
    txt_contact.Text = rs.Fields(2)
    End If
    d_id = rs.Fields(3)
    Call connection
    rs.Open "select name from department where id = " & d_id & "", conn, adOpenDynamic, adLockBatchOptimistic
    cmb_dept_name = rs.Fields(0)
            
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
 MSFlexGrid1.CellFontSize = 8
 MSFlexGrid1.CellFontUnderline = True
 MSFlexGrid1.CellTextStyle = flexTextInsetLight
 MSFlexGrid1.ColWidth(0) = 2200
 MSFlexGrid1.ColWidth(1) = 1500
 MSFlexGrid1.ColWidth(2) = 1100
 MSFlexGrid1.ColWidth(3) = 1100
 MSFlexGrid1.TextMatrix(0, 0) = "Subject"
 MSFlexGrid1.TextMatrix(0, 1) = "Course"
 MSFlexGrid1.TextMatrix(0, 2) = "Semester"
 MSFlexGrid1.TextMatrix(0, 3) = "Section"
 
 MSFlexGrid2.RowHeight(0) = 350
 MSFlexGrid2.Appearance = flex3D
 MSFlexGrid2.BackColorBkg = vbWhite
 MSFlexGrid2.FillStyle = flexFillRepeat
 MSFlexGrid2.Row = 0
 MSFlexGrid2.Col = 0
 MSFlexGrid2.RowSel = 0
 MSFlexGrid2.ColSel = 2
 MSFlexGrid2.BackColorSel = &H80000014
 MSFlexGrid2.ForeColorSel = &H80000013
 MSFlexGrid2.CellFontBold = True
 MSFlexGrid2.CellFontName = "Broadway"
 MSFlexGrid2.CellFontSize = 8
 MSFlexGrid2.CellFontUnderline = True
 MSFlexGrid2.CellTextStyle = flexTextInsetLight
 MSFlexGrid2.ColWidth(0) = 3000
 MSFlexGrid2.ColWidth(1) = 1450
 MSFlexGrid2.ColWidth(2) = 1450
 MSFlexGrid2.TextMatrix(0, 0) = "Course"
 MSFlexGrid2.TextMatrix(0, 1) = "Semester"
 MSFlexGrid2.TextMatrix(0, 2) = "Section"

 Call connection
 rs.Open "select count(id) from subjects where teacherid = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
 Dim n As Integer
 n = rs.Fields(0)
 Call connection
 rs.Open "select name,courseid,semester,section from subjects where teacherid = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
 MSFlexGrid1.Rows = n + 1
 For i = 1 To n
 MSFlexGrid1.TextMatrix(i, 0) = (rs.Fields(0))
 Call connection1
 rs2.Open "select name from course where id = " & (rs.Fields(1)) & "", conn1, adOpenDynamic, adLockBatchOptimistic
 MSFlexGrid1.TextMatrix(i, 1) = (rs2.Fields(0))
 rs2.MoveNext
 MSFlexGrid1.TextMatrix(i, 2) = (rs.Fields(2))
 MSFlexGrid1.TextMatrix(i, 3) = (rs.Fields(3))
 rs.MoveNext
 Next

     Call connection
 rs.Open "select count(teacherid) from class_coordinator where teacherid = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
 Dim n1 As Integer
 n1 = rs.Fields(0)
 Call connection
 rs.Open "select courseid,semester,section from class_coordinator where teacherid = " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
 MSFlexGrid2.Rows = n1 + 1
 For i = 1 To n1
 Call connection1
 rs2.Open "select name from course where id = " & (rs.Fields(0)) & "", conn1, adOpenDynamic, adLockBatchOptimistic
 MSFlexGrid2.TextMatrix(i, 0) = (rs2.Fields(0))
 rs2.MoveNext
 Select Case rs.Fields(1)
    Case 1: sem = "I"
    Case 2: sem = "II"
    Case 3: sem = "III"
    Case 4: sem = "IV"
    Case 5: sem = "V"
    Case 6: sem = "VI"
    Case 7: sem = "VII"
    Case 8: sem = "VIII"
 End Select
 MSFlexGrid2.TextMatrix(i, 1) = sem
 MSFlexGrid2.TextMatrix(i, 2) = (rs.Fields(2))
 rs.MoveNext
 Next


o_email = txt_email.Text
End Sub

Public Function connection1()
If conn1.State = 1 Then conn1.Close
conn1.Open "DSN=Semester_Process_Management"
End Function





Private Sub txt_contact_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If

End Sub

Private Sub txt_cur_pw_Click()
txt_cur_pw.BackColor = vbWhite
End Sub

Private Sub txt_email_Change()
txt_email.ForeColor = vbBlack

End Sub

Private Sub txt_email_Click()
txt_email.BackColor = vbWhite
End Sub

Private Sub txt_name_Click()
txt_name.BackColor = vbWhite
End Sub

Private Sub txt_new_pw_again_Click()
txt_new_pw_again.BackColor = vbWhite
End Sub

Private Sub txt_new_pw_Click()
txt_new_pw.BackColor = vbWhite
End Sub

Private Sub txt_username_Change()
txt_username.ForeColor = vbBlack
End Sub

Private Sub txt_username_Click()
txt_username.BackColor = vbWhite

End Sub
