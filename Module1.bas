Attribute VB_Name = "Module1"
Public conn As New ADODB.connection
Public rs As New ADODB.Recordset
Public teacher_id As Integer
Public frmEditCourse_oldsem As Integer
Public month_no As Integer
Public user_type As String




Public Function connection()
If conn.State = 1 Then conn.Close
conn.Open "DSN=Semester_Process_Management"
End Function
