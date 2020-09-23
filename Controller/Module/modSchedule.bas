Attribute VB_Name = "modSchedule"
Dim msql As String
Public Function FacultyInUse(cTimeIn As Date, cTimeOut As Date, cDay As String, cFaculty As String) As Boolean
Dim vRS As New ADODB.Recordset
If vRS.State = adStateOpen Then vRS.Close


msql = "SELECT tblSubject.SubjectTitle, tblSubjectOffering.Days & ' ' & tblSubjectOffering.TimeIn & '- ' & tblSubjectOffering.TimeOut AS Schedule, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName] AS TeacherFullName, tblSection.SectionTitle, tblSubjectOffering.TimeIn, tblSubjectOffering.TimeOut, tblSubjectOffering.Days, tblRoom.Building & ' - ' & tblRoom.Room AS Room " & _
        "FROM tblSubject INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblRoom INNER JOIN (tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) ON tblRoom.RoomID = tblSubjectOffering.RoomID) ON tblSection.SectionID = tblSubjectOffering.SectionID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
        "Where tblSubjectOffering.TimeIn >=#" & cTimeIn & "# and tblSubjectOffering.TimeOut <=#" & cTimeOut & "# and tblSubjectOffering.Days like'%" & cDay & "%' and tblSubjectOffering.TeacherID ='" & cFaculty & "'" & _
        "ORDER BY tblSubjectOffering.TimeIn;"

vRS.Open msql, con, 2, 3

'If ConnectRS(con, vRS, msql) = True Then
    If vRS.RecordCount >= 1 Then
        FacultyInUse = True
    Else
        FacultyInUse = False
    End If
'End If

Set vRS = Nothing
End Function
Public Function RoomInUse(cTimeIn As Date, cTimeOut As Date, cDay As String, cRoom As String) As Boolean
Dim vRS As New ADODB.Recordset
If vRS.State = adStateOpen Then vRS.Close

msql = "SELECT tblSubject.SubjectTitle, tblSubjectOffering.Days & ' ' & tblSubjectOffering.TimeIn & '- ' & tblSubjectOffering.TimeOut AS Schedule, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName] AS TeacherFullName, tblSection.SectionTitle, tblSubjectOffering.TimeIn, tblSubjectOffering.TimeOut, tblSubjectOffering.Days, tblRoom.Building & ' - ' & tblRoom.Room AS Room " & _
        "FROM tblSubject INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblRoom INNER JOIN (tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) ON tblRoom.RoomID = tblSubjectOffering.RoomID) ON tblSection.SectionID = tblSubjectOffering.SectionID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
        "Where tblSubjectOffering.TimeIn >=#" & cTimeIn & "# and tblSubjectOffering.TimeOut <=#" & cTimeOut & "# and tblSubjectOffering.Days like'%" & cDay & "%' and tblSubjectOffering.RoomID ='" & cRoom & "'" & _
        "ORDER BY tblSubjectOffering.TimeIn;"

vRS.Open msql, con, 2, 3

    'If ConnectRS(con, vRS, msql) = True Then
        If vRS.RecordCount >= 1 Then
        RoomInUse = True
        Else
        RoomInUse = False
        End If
    'End If
    
Set vRS = Nothing
End Function
Public Sub ShowTeacherSchedule(cFaculty As String, lstView As ListView)
Dim vRS As New ADODB.Recordset
Dim sSQL As String
Dim lv As ListItem
Dim time As Integer
Dim num1, num2 As Integer


On Error GoTo err

sSQL = "SELECT tblSubject.SubjectTitle, tblSubjectOffering.Days & ' ' & tblSubjectOffering.TimeIn & '- ' & tblSubjectOffering.TimeOut AS Schedule, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName] AS TeacherFullName, tblSection.SectionTitle, tblSubjectOffering.TimeIn, tblSubjectOffering.TimeOut, tblSubjectOffering.Days, tblRoom.Building & ' - ' & tblRoom.Room AS Room, Count(tblEnrolment.EnrollmentID) AS CountOfEnrollmentID, tblRoom.RoomID, tblTeacher.TeacherID,tblSection.SchoolYear, tblSection.Semester " & _
        "FROM tblSubject INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblRoom INNER JOIN ((tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblRoom.RoomID = tblSubjectOffering.RoomID) ON (tblSection.SectionID = tblSubjectOffering.SectionID) AND (tblSection.SectionID = tblGrade.SectionID)) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
        "WHERE tblTeacher.TeacherID ='" & cFaculty & "'and tblSection.SchoolYear='" & CurrentSchoolYear.SchoolYearTitle & "' and tblSection.Semester ='" & CurrentSemester.Semester & "'" & _
        "GROUP BY tblSubject.SubjectTitle, tblSubjectOffering.Days & ' ' & tblSubjectOffering.TimeIn & '- ' & tblSubjectOffering.TimeOut, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName], tblSection.SectionTitle, tblSubjectOffering.TimeIn, tblSubjectOffering.TimeOut, tblSubjectOffering.Days, tblRoom.Building & ' - ' & tblRoom.Room, tblRoom.RoomID, tblTeacher.TeacherID,tblSection.SchoolYear, tblSection.Semester;"

lstView.ListItems.Clear

   If ConnectRS(con, vRS, sSQL) = True Then
    Do Until vRS.EOF
            
            Set lv = lstView.ListItems.Add(, , vRS.Fields("SubjectTitle"))
            lv.SubItems(1) = vRS.Fields("SectionTitle")
            lv.SubItems(2) = vRS.Fields("CountOfEnrollmentID")
            lv.SubItems(3) = vRS.Fields("Schedule")
            lv.SubItems(4) = vRS.Fields("TeacherFullname")

        vRS.MoveNext
    Loop
    End If
Set vRS = Nothing
Exit Sub

err:
    Set vRS = Nothing
End Sub

Public Sub ShowRoomSchedule(cRoom As String, lstView As ListView)
Dim vRS As New ADODB.Recordset
Dim time As Integer
Dim sSQL As String
Dim lv As ListItem
Dim num1, num2 As Integer

On Error GoTo err


sSQL = "SELECT tblSubject.SubjectTitle, tblSubjectOffering.Days & ' ' & tblSubjectOffering.TimeIn & '- ' & tblSubjectOffering.TimeOut AS Schedule, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName] AS TeacherFullName, tblSection.SectionTitle, tblSubjectOffering.TimeIn, tblSubjectOffering.TimeOut, tblSubjectOffering.Days, tblRoom.Building & ' - ' & tblRoom.Room AS Room, Count(tblEnrolment.EnrollmentID) AS CountOfEnrollmentID, tblRoom.RoomID, tblTeacher.TeacherID,tblSection.SchoolYear, tblSection.Semester " & _
        "FROM tblSubject INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblRoom INNER JOIN ((tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblRoom.RoomID = tblSubjectOffering.RoomID) ON (tblSection.SectionID = tblSubjectOffering.SectionID) AND (tblSection.SectionID = tblGrade.SectionID)) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
        "WHERE tblRoom.RoomID ='" & cRoom & "' and tblSection.SchoolYear='" & CurrentSchoolYear.SchoolYearTitle & "' and tblSection.Semester ='" & CurrentSemester.Semester & "'" & _
        "GROUP BY tblSubject.SubjectTitle, tblSubjectOffering.Days & ' ' & tblSubjectOffering.TimeIn & '- ' & tblSubjectOffering.TimeOut, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName], tblSection.SectionTitle, tblSubjectOffering.TimeIn, tblSubjectOffering.TimeOut, tblSubjectOffering.Days, tblRoom.Building & ' - ' & tblRoom.Room, tblRoom.RoomID, tblTeacher.TeacherID,tblSection.SchoolYear, tblSection.Semester;"


    lstView.ListItems.Clear

    If ConnectRS(con, vRS, sSQL) = True Then
        Do Until vRS.EOF
            
            Set lv = lstView.ListItems.Add(, , vRS.Fields("SubjectTitle"))
            lv.SubItems(1) = vRS.Fields("SectionTitle")
            lv.SubItems(2) = vRS.Fields("CountOfEnrollmentID")
            lv.SubItems(3) = vRS.Fields("Schedule")
            lv.SubItems(4) = vRS.Fields("TeacherFullname")
            vRS.MoveNext
        Loop
    End If
Set vRS = Nothing
Exit Sub

err:
    Set vRS = Nothing
End Sub

Public Sub ShowConflictFaculty(cTimeIn As Date, cTimeOut As Date, cDay As String, cFaculty As String, lstView As ListView)
Dim vRS As New ADODB.Recordset
Dim sSQL As String
Dim lv As ListItem
Dim itmX As ListItem
Dim lvSI As ListSubItem
Dim intIndex As Integer

sSQL = "SELECT tblSubject.SubjectTitle, tblSubjectOffering.Days & ' ' & tblSubjectOffering.TimeIn & '- ' & tblSubjectOffering.TimeOut AS Schedule, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName] AS TeacherFullName, tblSection.SectionTitle, tblSubjectOffering.TimeIn, tblSubjectOffering.TimeOut, tblSubjectOffering.Days, tblRoom.Building & ' - ' & tblRoom.Room AS Room, Count(tblEnrolment.EnrollmentID) AS CountOfEnrollmentID " & _
        "FROM tblSubject INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblRoom INNER JOIN ((tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblRoom.RoomID = tblSubjectOffering.RoomID) ON (tblSection.SectionID = tblSubjectOffering.SectionID) AND (tblSection.SectionID = tblGrade.SectionID)) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
        "Where tblSubjectOffering.TimeIn >=#" & cTimeIn & "# and tblSubjectOffering.TimeOut <=#" & cTimeOut & "# and tblSubjectOffering.Days like'%" & cDay & "%' and tblSubjectOffering.TeacherID ='" & cFaculty & "'" & _
        "GROUP BY tblSubject.SubjectTitle, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName], tblSection.SectionTitle, tblSubjectOffering.TimeIn, tblSubjectOffering.TimeOut, tblSubjectOffering.Days, tblRoom.Building & ' - ' & tblRoom.Room " & _
        "ORDER BY tblSubjectOffering.TimeIn;"


vRS.Open sSQL, con, 2, 3

    'If ConnectRS(con, vRS, sSQL) = True Then
        lstView.ListItems.Clear
        Do Until vRS.EOF
            Set lv = lstView.ListItems.Add(, , vRS.Fields("SubjectTitle"))
                lv.SubItems(1) = vRS.Fields("SectionTitle")
                lv.SubItems(2) = vRS.Fields("CountOfEnrollmentID")
                lv.SubItems(3) = vRS.Fields("Schedule")
                lv.SubItems(4) = vRS.Fields("TeacherFullname")
                
            Set itmX = lstView.ListItems(lstView.ListItems.count)
            itmX.ForeColor = vbRed
            For intIndex = 1 To lstView.ColumnHeaders.count - 1
                Set lvSI = itmX.ListSubItems(intIndex)
                lvSI.ForeColor = itmX.ForeColor
                DoEvents
            Next

            vRS.MoveNext
        Loop
    'End If
Set vRS = Nothing
End Sub

Public Sub ShowConflictRoom(cTimeIn As Date, cTimeOut As Date, cDay As String, cFaculty As String, cRoom As String, lstView As ListView)
Dim vRS As New ADODB.Recordset
Dim sSQL As String
Dim lv As ListItem
Dim itmX As ListItem
Dim lvSI As ListSubItem
Dim intIndex As Integer

sSQL = "SELECT tblSubject.SubjectTitle, tblSubjectOffering.Days & ' ' & tblSubjectOffering.TimeIn & '- ' & tblSubjectOffering.TimeOut AS Schedule, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName] AS TeacherFullName, tblSection.SectionTitle, tblSubjectOffering.TimeIn, tblSubjectOffering.TimeOut, tblSubjectOffering.Days, tblRoom.Building & ' - ' & tblRoom.Room AS Room, Count(tblEnrolment.EnrollmentID) AS CountOfEnrollmentID " & _
        "FROM tblSubject INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblRoom INNER JOIN ((tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblRoom.RoomID = tblSubjectOffering.RoomID) ON (tblSection.SectionID = tblSubjectOffering.SectionID) AND (tblSection.SectionID = tblGrade.SectionID)) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
        "Where tblSubjectOffering.TimeIn >=#" & cTimeIn & "# and tblSubjectOffering.TimeOut <=#" & cTimeOut & "# and tblSubjectOffering.Days like'%" & cDay & "%' and tblSubjectOffering.RoomID ='" & cRoom & "'" & _
        "GROUP BY tblSubject.SubjectTitle, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName], tblSection.SectionTitle, tblSubjectOffering.TimeIn, tblSubjectOffering.TimeOut, tblSubjectOffering.Days, tblRoom.Building & ' - ' & tblRoom.Room " & _
        "ORDER BY tblSubjectOffering.TimeIn;"

vRS.Open sSQL, con, 2, 3

    'If ConnectRS(con, vRS, sSQL) = True Then
        lstView.ListItems.Clear
        Do Until vRS.EOF
        Set lv = lstView.ListItems.Add(, , vRS.Fields("SubjectTitle"))
            lv.SubItems(1) = vRS.Fields("SectionTitle")
            lv.SubItems(2) = vRS.Fields("CountOfEnrollmentID")
            lv.SubItems(3) = vRS.Fields("Schedule")
            lv.SubItems(4) = vRS.Fields("TeacherFullname")
            
            
            Set itmX = lstView.ListItems(lstView.ListItems.count)
            itmX.ForeColor = vbRed
            For intIndex = 1 To lstView.ColumnHeaders.count - 1
                Set lvSI = itmX.ListSubItems(intIndex)
                lvSI.ForeColor = itmX.ForeColor
                DoEvents
            Next
        
            vRS.MoveNext
        Loop
    'End If
Set vRS = Nothing
End Sub




