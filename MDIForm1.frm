VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Semester Process Management"
   ClientHeight    =   7095
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13755
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Department 
      Caption         =   "Department"
      Visible         =   0   'False
      Begin VB.Menu department_add 
         Caption         =   "Add"
      End
      Begin VB.Menu Department_View_Edit_Delete 
         Caption         =   "View/Edit/Delete"
      End
   End
   Begin VB.Menu course 
      Caption         =   "Course"
      Visible         =   0   'False
      Begin VB.Menu Course_add 
         Caption         =   "Add"
      End
      Begin VB.Menu course_View_Edit_Delete 
         Caption         =   "View/Edit/Delete"
      End
   End
   Begin VB.Menu subject 
      Caption         =   "Subject"
      Visible         =   0   'False
      Begin VB.Menu add 
         Caption         =   "Add"
      End
      Begin VB.Menu subject_View_Edit_Delete 
         Caption         =   "View/Edit/Delete"
      End
      Begin VB.Menu subject_allot_teacher 
         Caption         =   "Allot Teacher"
      End
   End
   Begin VB.Menu faculty 
      Caption         =   "Faculty"
      Visible         =   0   'False
      Begin VB.Menu teacher_add 
         Caption         =   "Add"
      End
      Begin VB.Menu teacher_View_Edit_Delete 
         Caption         =   "View/Edit/Delete"
      End
      Begin VB.Menu class_coordinator 
         Caption         =   "Class Coordinator"
         Begin VB.Menu classcoordinator_add 
            Caption         =   "Add"
         End
         Begin VB.Menu classcoordinator_View_Edit_Delete 
            Caption         =   "View/Edit/Delete"
         End
      End
   End
   Begin VB.Menu student 
      Caption         =   "Student"
      Visible         =   0   'False
      Begin VB.Menu atudent_add 
         Caption         =   "Add"
      End
      Begin VB.Menu student_View_Edit_Delete 
         Caption         =   "View/Edit/Delete"
      End
      Begin VB.Menu student_change_sem 
         Caption         =   "Change sem batch wise"
      End
      Begin VB.Menu student_CR 
         Caption         =   "CR"
         Begin VB.Menu CR_add 
            Caption         =   "Add"
         End
         Begin VB.Menu CR_View_Edit_Delete 
            Caption         =   "View/Edit/Delete"
         End
      End
   End
   Begin VB.Menu settings 
      Caption         =   "Settings"
      Visible         =   0   'False
      Begin VB.Menu admin_change_password 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu logout 
      Caption         =   "Log Out"
      Index           =   0
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
frm_add_sub.Show
End Sub

Private Sub admin_change_password_Click()
frm_change_pw.Show
End Sub


Private Sub atudent_add_Click()
frm_add_student.Show
End Sub

Private Sub classcoordinator_add_Click()
frm_add_CC.Show
End Sub

Private Sub classcoordinator_View_Edit_Delete_Click()
frm_view_cc.Show
End Sub

Private Sub Course_add_Click()
Frm_add_course.Show
End Sub

Private Sub course_View_Edit_Delete_Click()
frm_view_courses.Show
End Sub

Private Sub CR_add_Click()
frm_add_CR.Show
End Sub

Private Sub CR_View_Edit_Delete_Click()
frm_view_cr.Show
End Sub

Private Sub department_add_Click()
Frm_add_dept.Show
End Sub

Private Sub Department_View_Edit_Delete_Click()
frm_view_dept.Show
End Sub


Private Sub logout_Click(Index As Integer)
Unload Me
frm_login.Show
End Sub

Private Sub student_change_sem_Click()
frm_change_sem.Show
End Sub

Private Sub student_View_Edit_Delete_Click()
frm_view_students.Show
End Sub

Private Sub subject_allot_teacher_Click()
Frm_sub_assignment.Show
End Sub

Private Sub subject_View_Edit_Delete_Click()
frm_view_subject.Show
End Sub

Private Sub teacher_add_Click()
frm_add_teacher.Show
End Sub

Private Sub teacher_View_Edit_Delete_Click()
frm_view_teacher.Show
End Sub
