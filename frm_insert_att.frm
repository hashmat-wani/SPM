VERSION 5.00
Begin VB.Form frm_insert_att 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Attendance"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Attendence for"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   8535
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmd_back 
         Height          =   375
         Left            =   120
         Picture         =   "frm_insert_att.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Back"
         Top             =   360
         Width           =   255
      End
      Begin VB.ComboBox cmb_course_name 
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
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   840
         Width           =   5055
      End
      Begin VB.ComboBox cmb_sub_name 
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
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1800
         Width           =   5055
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
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   360
         Width           =   5055
      End
      Begin VB.Frame Frame1 
         Caption         =   "Reg. No."
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   5535
         Left            =   240
         TabIndex        =   38
         Top             =   2880
         Width           =   1695
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   120
            TabIndex        =   62
            Top             =   5040
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   120
            TabIndex        =   61
            Top             =   4680
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   60
            Top             =   4320
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   59
            Top             =   3960
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   58
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   57
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   56
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   55
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   54
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   53
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   52
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txt_reg_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Lectures Deliverd"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   5535
         Left            =   2640
         TabIndex        =   23
         Top             =   2880
         Width           =   2055
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   32
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   31
            Top             =   2520
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   30
            Top             =   2880
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   29
            Top             =   3240
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   28
            Top             =   3600
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   27
            Top             =   3960
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   26
            Top             =   4320
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   120
            TabIndex        =   25
            Top             =   4680
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_deliverd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   120
            TabIndex        =   24
            Top             =   5040
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Lectures Attended"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   5535
         Left            =   5400
         TabIndex        =   8
         Top             =   2880
         Width           =   2055
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   17
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   16
            Top             =   2520
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   15
            Top             =   2880
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   14
            Top             =   3240
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   13
            Top             =   3600
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   12
            Top             =   3960
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   11
            Top             =   4320
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   120
            TabIndex        =   10
            Top             =   4680
            Width           =   1695
         End
         Begin VB.TextBox txt_lect_attnd 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   120
            TabIndex        =   9
            Top             =   5040
            Width           =   1695
         End
      End
      Begin VB.ComboBox cmb_month 
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
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox cmb_year 
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ComboBox cmb_section 
         Height          =   315
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cmb_sem 
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
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   48
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lbl_sub 
         Caption         =   "Subject"
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
         Left            =   1080
         TabIndex        =   47
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Month/year"
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
         Left            =   600
         TabIndex        =   46
         Top             =   2280
         Width           =   1215
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
         Left            =   600
         TabIndex        =   45
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   44
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Section"
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
         Left            =   4320
         TabIndex        =   43
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Semester"
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
         Left            =   960
         TabIndex        =   42
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Action"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   0
      Top             =   8640
      Width           =   2895
      Begin VB.CommandButton cmd_save 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_rst 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_insert_att"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim teacher_rec As New ADODB.Recordset
'Dim arr1(0 To 20) As String
'Dim arr2(0 To 20) As String
'Dim count_arr1 As Byte
'Dim count_arr2 As Byte










Private Sub cmb_course_name_Click()
cmb_course_name.BackColor = vbWhite

cmb_sem.Clear
cmb_section.Clear
Call connection
rs.Open "select no_of_sems from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
For i = 1 To rs.Fields(0)
    Select Case i
        Case 1: cmb_sem.AddItem "I"
        Case 2: cmb_sem.AddItem "II"
        Case 3: cmb_sem.AddItem "III"
        Case 4: cmb_sem.AddItem "IV"
        Case 5: cmb_sem.AddItem "V"
        Case 6: cmb_sem.AddItem "VI"
        Case 7: cmb_sem.AddItem "VII"
        Case 8: cmb_sem.AddItem "VIII"
    End Select
Next
End Sub

Private Sub cmb_dept_name_Click()
cmb_dept_name.BackColor = vbWhite
cmb_sem.Clear
cmb_section.Clear
'If user_type = "Teacher" Then
'If rec.State = 1 Then rec.Close
'If teacher_rec.State = 1 Then teacher_rec.Close
        'Call connection
        Call connection
        rs.Open "select name from course where deptid =(select id from department where name = '" & cmb_dept_name.Text & "' )", conn, adOpenDynamic, adLockBatchOptimistic
            cmb_course_name.Clear
           While (rs.EOF = False)
                cmb_course_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend
 
         'teacher_rec.Open "select name from subjects where teacherid= '" & teacher_id & "' ", conn, adOpenDynamic, adLockBatchOptimistic
  ' although course name is according to the subject's course name he/she teaches but should be on the basis of department selected

 'While (rs.EOF = False)
        'arr1(count_arr1) = rs.Fields(0)
       'count_arr1 = count_arr1 + 1
       ' rs.MoveNext
'Wend
'While (teacher_rec.EOF = False)
 '       arr2(count_arr2) = teacher_rec.Fields(0)
  '      teacher_rec.MoveNext
   '     count_arr2 = count_arr2 + 1
'Wend

'For i = 0 To count_arr1 - 1
 '  For j = 0 To count_arr2 - 1
   '     If arr1(i) = arr2(j) Then
  '          cmb_course_name.AddItem arr1(i)
    '    End If
    'Next
'Next

'Else
'Call connection
 '       rs.Open "select name from course where deptid = (select id from department where name = '" & cmb_dept_name & "') ", conn, adOpenDynamic, adLockBatchOptimistic
  '          cmb_course_name.Clear
   '        While (rs.EOF = False)
    '            cmb_course_name.AddItem (rec.Fields(0))
     '           rs.MoveNext
      '      Wend
'End If
End Sub



Private Sub cmb_month_Click()
cmb_month.BackColor = vbWhite

End Sub

Private Sub cmb_section_Click()
cmb_section.BackColor = vbWhite

Call connection
        rs.Open "select name from subjects where courseid=(select id from course where name = '" & cmb_course_name.Text & "') and semester = '" & cmb_sem & "' and section = '" & cmb_section & "' and teacherid= " & teacher_id & "", conn, adOpenDynamic, adLockBatchOptimistic
          cmb_sub_name.Clear
          While (rs.EOF = False)
            cmb_sub_name.AddItem (rs.Fields(0))
            rs.MoveNext
          Wend
        
End Sub

Private Sub cmb_sem_Click()
cmb_sem.BackColor = vbWhite

cmb_section.Clear
Call connection
rs.Open "select no_of_secs from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
For i = 1 To rs.Fields(0)
    Select Case i
        Case 1: cmb_section.AddItem "A"
        Case 2: cmb_section.AddItem "B"
        Case 3: cmb_section.AddItem "C"
        Case 4: cmb_section.AddItem "D"
        Case 5: cmb_section.AddItem "E"
    End Select
Next

End Sub


Private Sub cmb_sub_name_Click()
cmb_sub_name.BackColor = vbWhite

End Sub

Private Sub cmb_year_Click()
cmb_year.BackColor = vbWhite

End Sub

Private Sub cmd_back_Click()
Unload Me
frm_teacher_login.Show
End Sub

Private Sub cmd_rst_Click()
Call Form_Load
End Sub

Private Sub cmd_save_Click()
If validation = False Then
    Exit Sub
End If

Dim i As Byte
Dim count As Byte
count = 0
    For i = 0 To 13
        If txt_reg_no(i).Text <> Empty Then
            count = count + 1
        End If
    Next
    Dim month_no As Byte
    Select Case cmb_month.Text
        Case MonthName(1)
            month_no = 1
        Case MonthName(2)
            month_no = 2
        Case MonthName(3)
            month_no = 3
        Case MonthName(4)
            month_no = 4
        Case MonthName(5)
            month_no = 5
        Case MonthName(6)
            month_no = 6
        Case MonthName(7)
            month_no = 7
        Case MonthName(8)
            month_no = 8
        Case MonthName(9)
            month_no = 9
        Case MonthName(10)
            month_no = 10
        Case MonthName(11)
            month_no = 11
        Case MonthName(12)
            month_no = 12
            
    End Select
    
     Call connection
     rs.Open " select id from course where name = '" & cmb_course_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
     c_id = rs.Fields(0)
     Call connection
     rs.Open " select id from subjects where name = '" & cmb_sub_name & "'", conn, adOpenDynamic, adLockBatchOptimistic
     s_id = rs.Fields(0)
     
        For i = 0 To count - 1
        If Trim(txt_reg_no(i).Text) <> Empty Then
            Call connection
    rs.Open "insert into attendance(courseid,semester,section,subjectid,regno,att_month,att_year,lec_delivered,lec_attended) values(" & c_id & ",'" & cmb_sem.Text & "','" & cmb_section.Text & "'," & s_id & ",'" & Trim(txt_reg_no(i).Text) & "','" & month_no & "','" & cmb_year.Text & "','" & Trim(Val(txt_lect_deliverd(i).Text)) & "','" & Trim(Val(txt_lect_attnd(i).Text)) & "')", conn, adOpenDynamic, adLockBatchOptimistic
    End If
    Next
    
            MsgBox "Record saved sucessfully", vbInformation
End Sub


Private Sub Form_Load()
For i = 0 To 13
txt_reg_no(i).Text = ""
txt_lect_deliverd(i).Text = ""
txt_lect_attnd(i).Text = ""
Next
cmb_dept_name.Clear
cmb_course_name.Clear
cmb_sem.Clear
cmb_section.Clear
cmb_sub_name.Clear
cmb_month.Clear
cmb_year.Clear
Call connection
rs.Open "select name from department", conn, adOpenDynamic, adLockBatchOptimistic
cmb_dept_name.Clear
           While (rs.EOF = False)
                cmb_dept_name.AddItem (rs.Fields(0))
                rs.MoveNext
            Wend


For i = 1 To 12
    cmb_month.AddItem MonthName(i)
Next

For i = 2000 To 2050
    cmb_year.AddItem i
Next
End Sub

Private Sub txt_lect_attnd_LostFocus(Index As Integer)
 ''If Val(txt_lect_attnd(Index).Text) = Empty Then
   ''  txt_lect_attnd(Index).BackColor = &HC0FFC0
     ''    MsgBox ("empty field not allowed")
       '' txt_lect_attnd(Index).SetFocus
       '' Else
     ''txt_lect_attnd(Index).BackColor = &H80000005
    '' End If
     
      
    ''If Val(txt_lect_deliverd(Index).Text) < Val(txt_lect_attnd(Index).Text) Then
           ' txt_lect_attnd(Index).ForeColor = green
           ' MsgBox ("Lecture Attended is more then Lecture Deliverd")
            'txt_lect_attnd(Index).SetFocus
   '' Else
             txt_lect_attnd(Index).ForeColor = black
   '' End If
   If Not (IsEmpty(txt_reg_no(Index).Text)) Then
        If Val(txt_lect_attnd(Index).Text) > Val(txt_lect_deliverd(Index).Text) Then
            txt_lect_attnd(Index).ForeColor = vbRed
            txt_lect_deliverd(Index).ForeColor = vbRed
            MsgBox "Lecture Attended is more then Lecture Deliverd", vbInformation
            
            
        Else
            txt_lect_attnd(Index).ForeColor = vbBlack
            txt_lect_deliverd(Index).ForeColor = vbBlack
        
         End If
    End If
End Sub



Private Sub txt_lect_deliverd_LostFocus(Index As Integer)
If txt_reg_no(Index).Text <> Empty Then

 If Val(txt_lect_deliverd(Index).Text) = Empty Then
     txt_lect_deliverd(Index).BackColor = &HC0FFC0
        MsgBox ("Empty field not allowed")
        txt_lect_deliverd(Index).SetFocus
       Else
     
        txt_lect_deliverd(Index).BackColor = &H80000005
    End If

        If Val(txt_lect_attnd(Index).Text) > Val(txt_lect_deliverd(Index).Text) Then
            txt_lect_attnd(Index).ForeColor = vbRed
            txt_lect_deliverd(Index).ForeColor = vbRed
            MsgBox "Lecture deliverd is less then Lecture attendend", vbInformation
            
            
        Else
            txt_lect_attnd(Index).ForeColor = vbBlack
            txt_lect_deliverd(Index).ForeColor = vbBlack
        
         End If
    End If
End Sub


        
        


Public Function validation()
          For Each ctr In frm_insert_att.Controls
           If TypeOf ctr Is ComboBox Then
            If Trim(ctr.Text) = Empty Then
                ctr.BackColor = vbGreen
            Else
                ctr.BackColor = vbWhite
            End If
           End If
          Next
       For Each ctr In frm_insert_att.Controls
            If TypeOf ctr Is ComboBox Then
                If ctr.BackColor = vbGreen Then
                    MsgBox "Fill all Fields", vbInformation
                    validation = False
                    Exit Function
                End If
            End If
        Next
validation = True
End Function

