VERSION 5.00
Begin VB.Form frm_login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LOGIN"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_login.frx":0000
   ScaleHeight     =   5880
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   4320
   End
   Begin VB.TextBox txt_pass 
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
      IMEMode         =   3  'DISABLE
      Left            =   6720
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txt_username 
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
      IMEMode         =   3  'DISABLE
      Left            =   6720
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ComboBox cmb_usertype 
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
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      Height          =   855
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label lbl_reset 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lbl_login 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lbl_atm 
      BackStyle       =   0  'Transparent
      Caption         =   "S.P.M"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmb_usertype_Click()
    If cmb_usertype.Text = "Student" Then
        txt_username.Text = Empty
        txt_username.Enabled = False
        txt_pass.Text = Empty
        txt_pass.Enabled = False
    Else
         txt_username.Enabled = True
         txt_pass.Enabled = True
    End If
End Sub







Private Sub lbl_login_Click()
user_type = Trim(cmb_usertype.Text)
    If Trim(cmb_usertype.Text) = "Student" Then
                    frm_entry_att.Show
        Unload Me
        
    ElseIf Trim(cmb_usertype.Text) = "Teacher" Then
          If Trim(txt_username.Text) = Empty Or Trim(txt_pass.Text) = Empty Then
            MsgBox "Empty field not Allowed", vbInformation
            Exit Sub
          End If
          Call connection
          rs.Open "select id from teacher where (email = '" & Trim(txt_username.Text) & "' or username = '" & Trim(txt_username.Text) & "')  and password = '" & Trim(txt_pass.Text) & "'", conn, adOpenDynamic, adLockBatchOptimistic
           If rs.EOF = True Then
           MsgBox "invalid Username/Passsword", vbInformation
           Exit Sub
           Else
            teacher_id = rs.Fields(0)
            frm_teacher_login.Show
            frm_teacher_login.lbl_teacher_username = RTrim(txt_username.Text)
            Unload Me
            
            Exit Sub
           End If
            
    ElseIf Trim(cmb_usertype.Text) = "Admin" Then
            If Trim(txt_username.Text) = Empty Or Trim(txt_pass.Text) = Empty Then
                MsgBox "Empty field not allowed", vbInformation
                Exit Sub
            End If
                Call connection
                rs.Open "select username from admin where username = '" & Trim(txt_username.Text) & "' and password = '" & Trim(txt_pass.Text) & "'", conn, adOpenDynamic, adLockBatchOptimistic
                If rs.EOF = True Then
                     MsgBox "invalid Email/Passsword", vbInformation
                Exit Sub
                End If
            MDIForm1.Department.Visible = True
            MDIForm1.course.Visible = True
            MDIForm1.subject.Visible = True
            MDIForm1.faculty.Visible = True
            MDIForm1.student.Visible = True
            MDIForm1.settings.Visible = True
            MDIForm1.Show

     End If
 Unload Me
End Sub

Private Sub lbl_reset_Click()
cmb_usertype.Text = "Admin"
txt_username.Text = Empty
txt_pass.Text = Empty
End Sub

Private Sub Form_Load()
cmb_usertype.AddItem "Admin"
cmb_usertype.AddItem "Teacher"
cmb_usertype.AddItem "Student"
cmb_usertype.Text = "Admin"
End Sub

Private Sub Timer1_Timer()
If lbl_atm.ForeColor = vbBlue Then
        lbl_atm.ForeColor = vbRed

ElseIf lbl_atm.ForeColor = vbRed Then
        lbl_atm.ForeColor = vbGreen
Else
        lbl_atm.ForeColor = vbBlue
End If
End Sub

