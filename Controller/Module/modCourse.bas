Attribute VB_Name = "modCourse"
Option Explicit

Public Const keyDepartment = "Col"

Public Type tCourse
    CourseID As String
    CourseTitle As String
    Major As String
    Years As Integer
    Curriculum As String
    DepartmentID As String
    CollegeID As String
    CurrentOffered As Integer
    Diploma As Integer
End Type

Public Function AddCourse(newCourse As tCourse) As TranDBResult
    
    Dim vRS As New ADODB.Recordset

    If CourseExistByID(newCourse.CourseID) = Success Then
        AddCourse = DuplicateID
        GoTo ReleaseAndExit
    End If

    If CourseExistByTitle(newCourse.CourseTitle) = Success Then
        AddCourse = DuplicateTitle
        GoTo ReleaseAndExit
    End If
    
    If CreateDefaultRSCourse(vRS) = Success Then
        vRS.AddNew
        
        vRS.Fields("CourseID").Value = newCourse.CourseID
        vRS.Fields("Course").Value = newCourse.CourseTitle
        vRS.Fields("CollegeID").Value = newCourse.CollegeID
        vRS.Fields("Major").Value = newCourse.Major
        vRS.Fields("Curriculum").Value = newCourse.Curriculum
        vRS.Fields("DepartmentID").Value = newCourse.DepartmentID
        vRS.Fields("CurrentOffered").Value = newCourse.CurrentOffered
        vRS.Fields("Diploma").Value = newCourse.Diploma
        vRS.Fields("YearLevel").Value = newCourse.Years
        vRS.Update
        AddCourse = Success
    Else
        AddCourse = NotConnected
    End If
    
    
    
ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function

Public Function EditCourse(newCourse As tCourse) As TranDBResult
    
    Dim OldCourse As tCourse
    Dim vRS As New ADODB.Recordset

    If GetCourseByID(newCourse.CourseID, OldCourse) Then
        If OldCourse.CourseTitle = newCourse.CourseTitle Then
            EditCourse = Success
            GoTo ReleaseAndExit
        Else
            If CourseExistByTitle(newCourse.CourseTitle) = Success Then
                EditCourse = DuplicateTitle
                GoTo ReleaseAndExit
            End If
        End If
    Else
        EditCourse = InvalidID
        GoTo ReleaseAndExit
    End If
    
    If ConnectRS(con, vRS, "SELECT * From tblCourse WHERE (((tblCourse.CourseID)='" & newCourse.CourseID & "'));") Then
        If vRS.RecordCount < 1 Then
            EditCourse = InvalidID
            GoTo ReleaseAndExit
        End If
    End If
    
        vRS.MoveFirst
        vRS.Fields("Course").Value = newCourse.CourseTitle
        vRS.Fields("CollegeID").Value = newCourse.CollegeID
        vRS.Fields("Major").Value = newCourse.Major
        vRS.Fields("Curriculum").Value = newCourse.Curriculum
        vRS.Fields("DepartmentID").Value = newCourse.DepartmentID
        vRS.Fields("CurrentOffered").Value = newCourse.CurrentOffered
        vRS.Fields("Diploma").Value = newCourse.Diploma
        vRS.Fields("YearLevel").Value = newCourse.Years
        vRS.Update
            
        EditCourse = Success
        

ReleaseAndExit:
    Set vRS = Nothing
End Function


Public Function DeleteCourse(sDepartmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    

    DeleteCourse = Failed
    
    If ConnectRS(con, vRS, "Delete * From tblCourse WHERE (((tblCourse.CourseID)='" & sDepartmentID & "'));") Then
        DeleteCourse = Success
    Else
        DeleteCourse = Failed
    End If

    Set vRS = Nothing
End Function




Public Function GetCourseMoveNext(ByRef vRS As ADODB.Recordset, ByRef vCourse As tCourse) As TranDBResult
    If Not vRS.EOF And Not vRS.BOF Then
        vCourse.CourseID = vRS.Fields("CourseID").Value
        vCourse.CollegeID = vRS.Fields("CollegeID").Value
        vCourse.CourseTitle = vRS.Fields("Course").Value
        vCourse.CurrentOffered = vRS.Fields("CurrentOffered").Value
        vCourse.Diploma = vRS.Fields("Diploma").Value
        vCourse.Curriculum = vRS.Fields("Curriculum").Value
        vCourse.DepartmentID = vRS.Fields("DepartmentID").Value
        vCourse.Major = vRS.Fields("Major").Value
        vCourse.Years = vRS.Fields("Years").Value
        vRS.MoveNext
        GetCourseMoveNext = Success
    Else
        GetCourseMoveNext = Failed
    End If
    
End Function



Public Function GetCourseByID(sCollegeID As String, ByRef vCourse As tCourse) As TranDBResult
    
    Dim vRS As New ADODB.Recordset

    If ConnectRS(con, vRS, "SELECT * From tblCourse WHERE (((tblCourse.CourseID)='" & sCollegeID & "'));") Then
        If AnyRecordExisted(vRS) Then
             vCourse.CourseID = vRS.Fields("CourseID").Value
            vCourse.CollegeID = vRS.Fields("CollegeID").Value
            vCourse.CourseTitle = vRS.Fields("Course").Value
            vCourse.CurrentOffered = vRS.Fields("CurrentOffered").Value
            vCourse.Diploma = vRS.Fields("Diploma").Value
            vCourse.Curriculum = vRS.Fields("Curriculum").Value
            vCourse.DepartmentID = vRS.Fields("DepartmentID").Value
            vCourse.Major = vRS.Fields("Major").Value
            vCourse.Years = vRS.Fields("YearLevel").Value
            GetCourseByID = Success
        Else
            GetCourseByID = Failed
        End If
    Else
        GetCourseByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function GetCourseByTitle(sCollegeTitle As String, ByRef vCourse As tCourse) As TranDBResult

    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT *  FROM tblCourse WHERE (((tblCourse.Course)='" & sCollegeTitle & "'));") Then
        If vRS.RecordCount > 0 Then
           vCourse.CourseID = vRS.Fields("CourseID").Value
            vCourse.CollegeID = vRS.Fields("CollegeID").Value
            vCourse.CourseTitle = vRS.Fields("Course").Value
            vCourse.CurrentOffered = CBool(vRS.Fields("CurrentOffered").Value)
            vCourse.Diploma = CBool(vRS.Fields("Diploma").Value)
            vCourse.Curriculum = vRS.Fields("Curriculum").Value
            vCourse.DepartmentID = vRS.Fields("DepartmentID").Value
            vCourse.Major = vRS.Fields("Major").Value
            vCourse.Years = vRS.Fields("YearLevel").Value
            GetCourseByTitle = Success
        Else
            GetCourseByTitle = Failed
        End If
    Else
        GetCourseByTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Private Sub ReadFromRecord(ByRef vRS As ADODB.Recordset, ByRef vCourse As tCourse)
    vCourse.CourseID = vRS.Fields("CourseID").Value
    vCourse.CollegeID = vRS.Fields("CollegeID").Value
    vCourse.CourseTitle = vRS.Fields("Course").Value
    vCourse.CurrentOffered = vRS.Fields("CurrentOffered").Value
    vCourse.Diploma = vRS.Fields("Diploma").Value
    vCourse.Curriculum = vRS.Fields("Curriculum").Value
    vCourse.DepartmentID = vRS.Fields("DepartmentID").Value
    vCourse.Major = vRS.Fields("Major").Value
    vCourse.Years = vRS.Fields("YearLevel").Value
End Sub

Public Function CourseExistByTitle(sCollegeTitle As String) As TranDBResult
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblCourse WHERE (((tblCourse.Course)='" & sCollegeTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            CourseExistByTitle = Success
        Else
            CourseExistByTitle = Failed
        End If
    Else
        CourseExistByTitle = Failed
    End If

    Set vRS = Nothing
End Function


Public Function CourseExistByID(sCollegeID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblCourse WHERE (((tblCourse.CourseID)='" & sCollegeID & "'));") Then
        If vRS.RecordCount > 0 Then
            CourseExistByID = Success
        Else
            CourseExistByID = Failed
        End If
    Else
        CourseExistByID = Failed
       
    End If

    Set vRS = Nothing
End Function


Public Function CreateDefaultRSCourse(ByRef vRS As ADODB.Recordset) As TranDBResult
    CreateDefaultRSCourse = Failed
    
    If ConnectRS(con, vRS, "SELECT * FROM tblCourse") Then
        CreateDefaultRSCourse = Success
    End If
End Function

Public Function CourseRecordExist() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSCourse(vRS) = Success Then
        
        If AnyRecordExisted(vRS) = True Then
            CourseRecordExist = Success
        Else
            CourseRecordExist = Failed
        End If
        
    Else
        CourseRecordExist = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function GetNewCourseID(ByRef sNewCourseID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim NewDNumber As Integer

    GetNewCourseID = Failed
    
    sSQL = "SELECT 'Course-' & String$(2-Len(Count(*)+1),'0') & Count(*)+1 AS NewID" & _
            " FROM tblCourse;"
            
    If ConnectRS(con, vRS, sSQL) = True Then
        
        sNewCourseID = vRS.Fields("NewID").Value
        
        While DepartmentExistByID(sNewCourseID) = Success
            NewDNumber = Val(Right(sNewCourseID, 2)) + 1
            sNewCourseID = "D-" & String(2 - Len(NewDNumber), "0") & NewDNumber
        Wend
        
       GetNewCourseID = Success
    
    Else
    
        GetNewCourseID = Failed
    End If

    Set vRS = Nothing

End Function






