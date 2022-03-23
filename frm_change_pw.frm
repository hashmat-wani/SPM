VERSION 5.00
Begin VB.Form frm_change_pw 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
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
         TabIndex        =   4
         Top             =   2280
         Width           =   1575
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
         TabIndex        =   3
         Top             =   960
         Width           =   2895
      End
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
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   360
         Width           =   2895
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
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
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
         TabIndex        =   6
         Top             =   480
         Width           =   1815
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
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frm_change_pw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_change_password_Click()
If txt_cur_pw.Text = "" Or txt_new_pw.Text = "" Or txt_new_pw_again = "" Then
     MsgBox "Empty Fields not allowed", vbInformation
Else
    Call connection
    rs.Open "select password from admin where username = 'admin'", conn, adOpenDynamic, adLockBatchOptimistic
    If Trim(txt_cur_pw.Text) = rs.Fields(0) Then
        If Trim(txt_new_pw.Text) = Trim(txt_new_pw_again.Text) Then
            Call connection
            rs.Open "update admin set password = '" & Trim(txt_new_pw.Text) & "' where username = 'admin'", conn, adOpenDynamic, adLockBatchOptimistic
            MsgBox "Password changed successfully", vbInformation
            Unload Me
            
            Exit Sub
        Else
            MsgBox "Passwords do not match", vbInformation
        End If
    Else
        MsgBox " Your old password was entered incorrectly. Please try it again", vbInformation
        txt_cur_pw.SetFocus
    End If
End If

End Sub

