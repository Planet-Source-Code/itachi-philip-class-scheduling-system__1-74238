Attribute VB_Name = "modRSSectionOffering"
Option Explicit

Public Const KeySectionOffering = "seof"

Public Type tSectionOffering
    SectionID As String
    SectionTitle As String
    SchoolYear As String
    DepartmentID As String
    Slots As Double
    Semester As String
    CreationDate As Date
    CreatedBy As String
    ModifiedDate As Date
    ModifiedBy As String
End Type


Public Function AddSectionOffering(vSectionID As String, vSectionOffering As tSectionOffering) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    If SectionOfferingExistByID(vSectionOffering.SectionID) = Success Then
        AddSectionOffering = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT * FROM tblSection"
    
    If ConnectRS(con, vRS, sSQL) = True Then
        vRS.AddNew
        vRS.Fields("SectionID").Value = vSectionOffering.SectionID
        vRS.Fields("SectionTitle").Value = vSectionOffering.SectionTitle
        vRS.Fields("SchoolYear").Value = vSectionOffering.SchoolYear
        vRS.Fields("DepartmentID").Value = vSectionOffering.DepartmentID
        vRS.Fields("Slots").Value = vSectionOffering.Slots
        vRS.Fields("Semester").Value = vSectionOffering.Semester
        vRS.Fields("CreationDate").Value = vSectionOffering.CreationDate
        vRS.Fields("CreatedBy").Value = vSectionOffering.CreatedBy
        
        vRS.Update
        
        AddSectionOffering = Success
    Else
        AddSectionOffering = Failed
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
    
End Function


Public Function EditSectionOffering(vSectionOffering As tSectionOffering) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    If SectionOfferingExistByID(vSectionOffering.SectionID) <> Success Then
        EditSectionOffering = InvalidID
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT * FROM tblSection WHERE SectionID='" & vSectionOffering.SectionID & "'"
    
    If ConnectRS(con, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            vRS.Fields("SectionID").Value = vSectionOffering.SectionID
            vRS.Fields("SchoolYear").Value = vSectionOffering.SchoolYear
            vRS.Fields("Slots").Value = vSectionOffering.Slots
            vRS.Fields("Semester").Value = vSectionOffering.Semester
            vRS.Fields("ModifiedDate").Value = vSectionOffering.ModifiedDate
            vRS.Fields("ModifiedBy").Value = vSectionOffering.ModifiedBy
        
            vRS.Update
            
            EditSectionOffering = Success
            
        Else
        
            EditSectionOffering = InvalidID
        End If
    Else
        EditSectionOffering = Failed
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
    
End Function



Public Function DeleteSectionOffering(sSectionOfferingID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "Delete * From tblSection WHERE (((tblSection.SectionID)='" & sSectionOfferingID & "'));") Then
        DeleteSectionOffering = Success
    Else
        DeleteSectionOffering = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function





Public Function GetSectionOfferingByID(sSectionOfferingID As String, ByRef vSectionOffering As tSectionOffering) As TranDBResult
'On Error Resume Next
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    
    'default
    GetSectionOfferingByID = Failed
    
    
    If Len(sSectionOfferingID) < 1 Then
        GetSectionOfferingByID = Failed
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT * FROM tblSection WHERE SectionID='" & sSectionOfferingID & "'"
    
    If ConnectRS(con, vRS, sSQL) = True Then
    'vRS.Open sSQL, con, 2, 3
        If AnyRecordExisted(vRS) = True Then
        
            vSectionOffering.SectionID = (vRS.Fields("SectionID"))
            vSectionOffering.SchoolYear = (vRS.Fields("SchoolYear"))
            vSectionOffering.Slots = (vRS.Fields("Slots"))
            vSectionOffering.Semester = (vRS.Fields("Semester"))
            vSectionOffering.CreationDate = (vRS.Fields("CreationDate"))
            vSectionOffering.CreatedBy = (vRS.Fields("CreatedBy"))
            GetSectionOfferingByID = Success
        
        Else
            GetSectionOfferingByID = Failed
        End If
    Else
        GetSectionOfferingByID = Failed
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function SectionOfferingExistByID(sSectionOfferingID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    SectionOfferingExistByID = Failed
    
    If Len(sSectionOfferingID) < 1 Then Exit Function
    
    sSQL = " SELECT tblSection.SectionID" & _
            " From tblSection " & _
            " WHERE (((tblSection.SectionID)='" & sSectionOfferingID & "'));"
            
    If ConnectRS(con, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            SectionOfferingExistByID = Success
        Else
            SectionOfferingExistByID = Failed
        End If
    Else
        SectionOfferingExistByID = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function TeacherAssignedBySchoolYear(sTeacherID As String, sSchoolYearTitle As String) As TranDBResult
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    TeacherAssignedBySchoolYear = Failed
    
    If Len(sSchoolYearTitle) < 1 Or Len(sTeacherID) < 1 Then Exit Function
    
    sSQL = "SELECT tblSection.TeacherID, tblSection.SchoolYear" & _
            " From tblSection" & _
            " WHERE (((tblSection.TeacherID)='" & sTeacherID & "') AND ((tblSection.SchoolYear)='" & sSchoolYearTitle & "'));"
            

    If ConnectRS(con, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            TeacherAssignedBySchoolYear = Success
        Else
            TeacherAssignedBySchoolYear = Failed
        End If
    Else
        TeacherAssignedBySchoolYear = Failed
    End If
    
    Set vRS = Nothing
End Function


Public Function GetAutoSectionOffering(sSchoolYear As String, sDepartmentID As String, iYearLevelID As Integer, dStudentPrevAveGrade As Double, ByRef sReturnSectionOfferingID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    sSQL = "SELECT tblSection.SectionID, Count(tblEnrolment.EnrolmentID) AS CountOfEnrolmentID, tblSection.MinGrade, tblSection.MaxStudentCount, ([tblSection]![MaxGrade]+[tblSection]![MinGrade]) AS GradeRank, tblSection.MaxGrade, tblSection.CreationDate, tblSection.SchoolYear, tblSection.DepartmentID, tblSection.YearLevelID" & _
            " FROM tblSection LEFT JOIN tblEnrolment ON tblSection.SectionID = tblEnrolment.SectionID" & _
            " GROUP BY tblSection.SectionID, tblSection.MinGrade, tblSection.MaxStudentCount, ([tblSection]![MaxGrade]+[tblSection]![MinGrade]), tblSection.MaxGrade, tblSection.CreationDate, tblSection.SchoolYear, tblSection.DepartmentID, tblSection.YearLevelID" & _
            " Having (((Count(tblEnrolment.EnrolmentID)) < [tblSection]![MaxStudentCount]) And ((tblSection.MinGrade) <= " & dStudentPrevAveGrade & ") And ((tblSection.MaxGrade) >= " & dStudentPrevAveGrade & ") And ((tblSection.SchoolYear) = '" & sSchoolYear & "') And ((tblSection.DepartmentID) = '" & sDepartmentID & "') And ((tblSection.YearLevelID) = " & iYearLevelID & "))" & _
            " ORDER BY ([tblSection]![MaxGrade]+[tblSection]![MinGrade]) DESC , tblSection.MaxGrade DESC , tblSection.CreationDate;"


    'Clipboard.SetText sSQL
    'defaults
    sReturnSectionOfferingID = ""
    GetAutoSectionOffering = Failed
    
    If ConnectRS(con, vRS, sSQL) = False Then
        'temp
        GetAutoSectionOffering = Failed
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetAutoSectionOffering = Failed
        GoTo ReleaseAndExit
    End If
    
    'success
    sReturnSectionOfferingID = (vRS.Fields("SectionID"))
    GetAutoSectionOffering = Success
    

ReleaseAndExit:
    Set vRS = Nothing
End Function
