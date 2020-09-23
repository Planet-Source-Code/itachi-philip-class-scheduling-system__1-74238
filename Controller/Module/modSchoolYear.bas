Attribute VB_Name = "modSchoolYear"
Option Explicit


Public Const keySchoolYear = "scho"

Public Type tSchoolYear
    SchoolYearTitle As String
    Locked As Boolean
End Type

Public Type tSemester
    SemesterID As String
    Semester As String
End Type

Public Function SchoolYearRecordExisted() As TranDBResult

    Dim vRS As New ADODB.Recordset
    

    If CreateDefaultrsYear(vRS) <> Success Then
        SchoolYearRecordExisted = Failed
        GoTo ReleaseAndExit
    End If
    

    If AnyRecordExisted(vRS) = True Then
        SchoolYearRecordExisted = Success
    Else
        SchoolYearRecordExisted = Failed
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function AddSchoolYear(newSchoolYear As tSchoolYear) As TranDBResult

    
    Dim vRS As New ADODB.Recordset
    
    'default
    AddSchoolYear = Failed
        
    If CreateDefaultrsYear(vRS) <> Success Then
        AddSchoolYear = NotConnected
        GoTo ReleaseAndExit
    End If
    
    If SchoolYearExistByTitle(newSchoolYear.SchoolYearTitle) = Success Then
        AddSchoolYear = DuplicateTitle
        GoTo ReleaseAndExit
    End If
    
    vRS.AddNew
    
    vRS.Fields("schoolyear").Value = newSchoolYear.SchoolYearTitle
    vRS.Fields("Locked").Value = newSchoolYear.Locked


    vRS.Update

    AddSchoolYear = Success

    
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function EditSchoolYear(OldSchoolYearTitle As String, newSchoolYear As tSchoolYear) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If OldSchoolYearTitle = newSchoolYear.SchoolYearTitle Then
        'nothing to process, hust return success
        EditSchoolYear = Success
    Else
        'find duplicate
        If SchoolYearExistByTitle(newSchoolYear.SchoolYearTitle) = Success Then
            EditSchoolYear = DuplicateTitle
        Else

            If ConnectRS(con, vRS, "SELECT  * From tblSchoolYear WHERE (((tblSchoolYear.SchoolYear)='" & OldSchoolYearTitle & "'));") Then
            
                'edit
                vRS.MoveFirst
                vRS.Fields("schoolyear").Value = newSchoolYear.SchoolYearTitle
                vRS.Fields("Locked").Value = newSchoolYear.Locked

                vRS.Update
        
                EditSchoolYear = Success
                'edited
            Else
                EditSchoolYear = Failed
            End If
        End If
    End If
        

    Set vRS = Nothing
End Function






Public Function ExecuteDeleteSchoolYear(sSchoolYearTitle As String) As TranDBResult
    
    Dim vSchoolYear As tSchoolYear
    Dim DeleteResult As Integer
    'default
    ExecuteDeleteSchoolYear = Failed
    
    'check if record exist and if it is edited by other user
    If MsgBox("WARNING:" & vbNewLine & _
    "Deleting School Year Record will affect all other record" & vbNewLine & vbNewLine & _
    "Delete this record anyway?", vbQuestion + vbYesNo) = vbYes Then
    
        If Len(sSchoolYearTitle) < 1 Then Exit Function
        
        
        'delete file
        DeleteResult = DeleteSchoolYear(sSchoolYearTitle)
        
        Select Case DeleteResult
            
            Case 1 'deleted
                MsgBox "School Year deleted.", vbInformation
            
            Case Else 'failed
                MsgBox "Deleting School Year went failed.", vbExclamation
                
        End Select
        
        
    End If
    
    ExecuteDeleteSchoolYear = DeleteResult
End Function





Public Function DeleteSchoolYear(sSchoolYearTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
        If ConnectRS(con, vRS, "DELETE tblSchoolYear.SchoolYear From tblSchoolYear WHERE (((tblSchoolYear.SchoolYear)='" & sSchoolYearTitle & "'));") Then
            DeleteSchoolYear = Success
        Else
            DeleteSchoolYear = Failed
        End If
        
    Set vRS = Nothing
End Function



Public Function GetSchoolYearMoveNext(ByRef vRS As ADODB.Recordset, ByRef vSchoolYear As tSchoolYear) As TranDBResult
    If Not vRS.EOF Then
        vSchoolYear.SchoolYearTitle = vRS.Fields("schoolyear").Value
        vSchoolYear.Locked = vRS.Fields("Lock").Value

        vRS.MoveNext
        GetSchoolYearMoveNext = Success
    Else
        GetSchoolYearMoveNext = Failed
    End If
    
End Function

Public Function SchoolYearExistByTitle(sSchoolYearTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    'default
    SchoolYearExistByTitle = Failed
        
    If CreateDefaultrsYear(vRS) <> 1 Then
        SchoolYearExistByTitle = Failed
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) Then
        vRS.MoveFirst
        vRS.Find "schoolyear= '" & sSchoolYearTitle & "'"
        
        If RecordNoMatch(vRS) Then
            SchoolYearExistByTitle = Failed
        Else
            SchoolYearExistByTitle = Success
        End If
    Else
        SchoolYearExistByTitle = Failed
    End If
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Function



Public Function CreateDefaultrsYear(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultrsYear = Failed
    
    If ConnectRS(con, vRS, "SELECT * FROM tblSchoolYear") Then
        CreateDefaultrsYear = Success
    End If
End Function

Public Function GetNextSchoolYear(sOldSchoolYear As String, ByRef newSchoolYear As String) As TranDBResult
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblSchoolYear.SchoolYear" & _
        " From tblSchoolYear" & _
        " Where (((Val(Left([tblSchoolYear]![SchoolYear], 4))) > " & Left(sOldSchoolYear, 4) & "))" & _
        " ORDER BY tblSchoolYear.SchoolYear;"
    
    If ConnectRS(con, vRS, sSQL) = True Then
        newSchoolYear = (vRS.Fields("SchoolYear"))
        GetNextSchoolYear = Success
    Else
        GetNextSchoolYear = Failed
    End If
    
    
    Set vRS = Nothing
End Function



Public Function GetSchoolYearByTitle(sSchoolYearTitle As String, ByRef vSchoolYear As tSchoolYear) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSchoolYearByTitle = Failed
        
    sSQL = "SELECT * FROM tblSchoolYear WHERE tblSchoolYear.SchoolYear='" & sSchoolYearTitle & "'"
      
    If ConnectRS(con, vRS, sSQL) = False Then
        GetSchoolYearByTitle = Failed
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) Then
    
        vSchoolYear.SchoolYearTitle = sSchoolYearTitle
        vSchoolYear.Locked = (vRS.Fields("Locked"))
    
        GetSchoolYearByTitle = Success

    Else
        GetSchoolYearByTitle = Failed
    End If
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Function
Public Function GetSemesterByTitle(sSemesterTitle As String, ByRef vSem As tSemester) As TranDBResult

    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT *  FROM tbSemester WHERE (((tblDepartment.Semester)='" & sSemesterTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            vSem.SemesterID = vRS.Fields("SemesterID").Value
            vSem.Semester = vRS.Fields("Semester").Value
            GetSemesterByTitle = Success
        Else
            GetSemesterByTitle = Failed
        End If
    Else
        GetSemesterByTitle = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function SaveActiveSchoolYear(sSYTitle As String)
    SaveSetting App.Title, "DataSetting", "activeschoolyear", sSYTitle
End Function

Public Function GetActiveSchoolYear() As String
    GetActiveSchoolYear = GetSetting(App.Title, "DataSetting", "activeschoolyear", "0000")
End Function

Public Function SaveActiveSemester(sSemester As String)
    SaveSetting App.Title, "DataSetting", "activesemester", sSemester
End Function
Public Function GetActiveSemester() As String
     GetActiveSemester = GetSetting(App.Title, "DataSetting", "activesemester", "0000")
End Function
