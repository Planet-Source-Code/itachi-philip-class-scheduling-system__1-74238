Attribute VB_Name = "modDBTeacher"
Option Explicit


Public Const KeyTeacher = "teac"

Public Type tTeacher
    TeacherID As String
    Gender As String
    FirstName As String
    MiddleName As String
    LastName As String
    OnService As String
    Department As String
    Username As String
    Password As String
    CreationDate As Date
End Type

Public Function TeacherRecordExisted() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSTeacher(vRS) = Success Then
        If AnyRecordExisted(vRS) = True Then
            TeacherRecordExisted = Success
        Else
            TeacherRecordExisted = Failed
        End If
    Else
        TeacherRecordExisted = Failed
    End If
    
    Set vRS = Nothing
End Function
Public Function AddTeacher(newTeacher As tTeacher) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim oldTeacher As tTeacher
    
    
    'check duplicate id
    If TeacherExistByID(newTeacher.TeacherID) = Success Then
        AddTeacher = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    
        'check each field
        If Len(Trim(newTeacher.TeacherID)) < 1 Then
            AddTeacher = InvalidID
            GoTo ReleaseAndExit
        End If
        
        If Len(Trim(newTeacher.FirstName)) < 1 Then
            AddTeacher = InvalidTeacherFirstName
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.MiddleName)) < 1 Then
            AddTeacher = InvalidTeacherMiddleName
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.LastName)) < 1 Then
            AddTeacher = InvalidTeacherLastName
            GoTo ReleaseAndExit
        End If


        
        If ConnectRS(con, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherID)='" & newTeacher.TeacherID & "'));") Then
            
            vRS.AddNew
            vRS.Fields("TeacherID").Value = newTeacher.TeacherID
            vRS.Fields("FirstName").Value = newTeacher.FirstName
            vRS.Fields("MiddleName").Value = newTeacher.MiddleName
            vRS.Fields("LastName").Value = newTeacher.LastName
            vRS.Fields("Gender").Value = newTeacher.Gender
            vRS.Fields("Status").Value = newTeacher.OnService
            vRS.Fields("DepartmentID").Value = newTeacher.Department
            vRS.Fields("Username").Value = newTeacher.Username
            vRS.Fields("Password").Value = newTeacher.Password
            vRS.Fields("CreationDate").Value = Now
            'ignore creation date
    
            vRS.Update
            
            AddTeacher = Success
        Else
            AddTeacher = Failed
        End If


ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function EditTeacher(newTeacher As tTeacher) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim oldTeacher As tTeacher
    
    If GetTeacherByID(newTeacher.TeacherID, oldTeacher) = Success Then

        If Len(Trim(newTeacher.FirstName)) < 1 Then
            EditTeacher = InvalidTeacherFirstName
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.MiddleName)) < 1 Then
            EditTeacher = InvalidTeacherMiddleName
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.LastName)) < 1 Then
            EditTeacher = InvalidTeacherLastName
            GoTo ReleaseAndExit
        End If
        
        
        If ConnectRS(con, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherID)='" & newTeacher.TeacherID & "'));") Then
            vRS.Fields("FirstName").Value = newTeacher.FirstName
            vRS.Fields("MiddleName").Value = newTeacher.MiddleName
            vRS.Fields("LastName").Value = newTeacher.LastName
            vRS.Fields("Status").Value = newTeacher.OnService
            vRS.Fields("Gender").Value = newTeacher.Gender
            vRS.Fields("DepartmentID").Value = newTeacher.Department
            vRS.Fields("Username").Value = newTeacher.Username
            vRS.Fields("Password").Value = newTeacher.Password
            vRS.Update
            
            EditTeacher = Success
        Else
            EditTeacher = Failed
        End If
    Else
        'teacher by id not found
        EditTeacher = InvalidID
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function ExecDeleteTeacher(sTeacherID As String) As TranDBResult
        
            If MsgBox("You are about to delete this Teacher with ID :" & vbNewLine & sTeacherID & vbNewLine & "Are you sure to DELETE this Teahcer Account?", vbQuestion + vbOKCancel) = vbOK Then
                
                If DeleteTeacher(sTeacherID) = Success Then
                    MsgBox "TEACHER entry successfully deleted.", vbInformation
                    ExecDeleteTeacher = Success
                Else
                    MsgBox "Unable to delete Teacher Account. The current was edited by another user", vbExclamation
                    ExecDeleteTeacher = Failed
                End If
            Else
                ExecDeleteTeacher = Failed
            End If
        
End Function
Public Function DeleteTeacher(sTeacherID As String) As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "DELETE tblTeacher.TeacherID From tblTeacher WHERE (((tblTeacher.TeacherID)='" & sTeacherID & "'));") Then
        DeleteTeacher = Success
    Else
        DeleteTeacher = Failed
    End If

    Set vRS = Nothing
End Function

Public Function GetTeacherMoveNext(ByRef vRS As ADODB.Recordset, ByRef vTeacher As tTeacher) As TranDBResult
    
    If Not vRS.EOF And Not vRS.BOF Then
        vTeacher.TeacherID = (vRS.Fields("teacherid"))
        vTeacher.FirstName = (vRS.Fields("FirstName"))
        vTeacher.MiddleName = (vRS.Fields("MiddleName"))
        vTeacher.LastName = (vRS.Fields("LastName"))
        vTeacher.Gender = (vRS.Fields("Gender"))
        vTeacher.OnService = (vRS.Fields("Status"))
        vTeacher.CreationDate = (vRS.Fields("CreationDate"))
        vTeacher.Department = vRS.Fields("DeparmentID")
        vTeacher.Username = vRS.Fields("Username")
        vTeacher.Password = vRS.Fields("Password")
        vRS.MoveNext
        
        GetTeacherMoveNext = Success
        
    Else
    
        GetTeacherMoveNext = Failed
        
    End If
    
End Function


Public Function GetTeacherByTitle(sTeacherTitle As String, ByRef vTeacher As tTeacher) As TranDBResult
        
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherTitle)='" & sTeacherTitle & "'));") Then
        If AnyRecordExisted(vRS) Then
            vTeacher.TeacherID = (vRS.Fields("teacherid"))
            vTeacher.FirstName = (vRS.Fields("FirstName"))
            vTeacher.MiddleName = (vRS.Fields("MiddleName"))
            vTeacher.LastName = (vRS.Fields("LastName"))
            vTeacher.Gender = (vRS.Fields("Gender"))
            vTeacher.OnService = (vRS.Fields("Status"))
            vTeacher.Department = vRS.Fields("DepartmentID")
            vTeacher.CreationDate = (vRS.Fields("CreationDate"))
            vTeacher.Username = vRS.Fields("Username")
            vTeacher.Password = vRS.Fields("Password")
            
            GetTeacherByTitle = Success
        Else
            GetTeacherByTitle = Failed
        End If
    Else
        
        GetTeacherByTitle = Failed
    End If
    
    Set vRS = Nothing
        
End Function

Public Function GetTeacherByFullName(sTeacherFullName As String, ByRef vTeacher As tTeacher) As TranDBResult
        
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    
    sSQL = "SELECT *" & _
            " From tblTeacher" & _
            " WHERE ((([LastName] & ', ' & [FirstName] & ' ' & [MiddleName])='" & sTeacherFullName & "'));"

    If ConnectRS(con, vRS, sSQL) Then
        If AnyRecordExisted(vRS) Then
            vTeacher.TeacherID = (vRS.Fields("teacherid"))
            vTeacher.FirstName = (vRS.Fields("FirstName"))
            vTeacher.MiddleName = (vRS.Fields("MiddleName"))
            vTeacher.LastName = (vRS.Fields("LastName"))
            vTeacher.Gender = (vRS.Fields("Gender"))
            vTeacher.Department = vRS.Fields("DepartmentID")
            vTeacher.OnService = (vRS.Fields("Status"))
            vTeacher.CreationDate = (vRS.Fields("CreationDate"))
            vTeacher.Username = vRS.Fields("Username")
            vTeacher.Password = vRS.Fields("Password")
            
            GetTeacherByFullName = Success
        Else
            GetTeacherByFullName = Failed
        End If
    Else
        
        GetTeacherByFullName = Failed
    End If
    
    Set vRS = Nothing
        
End Function

Public Function GetTeacherByID(sTeacherID As String, ByRef vTeacher As tTeacher) As TranDBResult
On Error Resume Next
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherID)='" & sTeacherID & "'));") Then
        If AnyRecordExisted(vRS) Then
            vTeacher.TeacherID = (vRS.Fields("teacherid"))
            vTeacher.FirstName = (vRS.Fields("FirstName"))
            vTeacher.MiddleName = (vRS.Fields("MiddleName"))
            vTeacher.LastName = (vRS.Fields("LastName"))
            vTeacher.Gender = (vRS.Fields("Gender"))
            vTeacher.OnService = (vRS.Fields("Status"))
            vTeacher.Department = vRS.Fields("DepartmentID")
            vTeacher.CreationDate = (vRS.Fields("CreationDate"))
            vTeacher.Username = vRS.Fields("Username")
            vTeacher.Password = vRS.Fields("Password")
            
            GetTeacherByID = Success
        Else
            GetTeacherByID = Failed
        End If
    Else
        
        GetTeacherByID = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function TeacherExistByTitle(sTeacherTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherTitle)='" & sTeacherTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            TeacherExistByTitle = Success
        Else
            TeacherExistByTitle = Failed
        End If
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function TeacherExistByFullName(sTeacherFullName As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'deafult
    TeacherExistByFullName = Failed
    
    
    If Len(sTeacherFullName) < 1 Then Exit Function
    
    sSQL = "SELECT [LastName] & ', ' & [FirstName] & ' ' & [MiddleName] AS TeacherFullName" & _
            " From tblTeacher" & _
            " WHERE ((([LastName] & ', ' & [FirstName] & ' ' & [MiddleName])='" & sTeacherFullName & "'));"

    If ConnectRS(con, vRS, sSQL) Then
        If vRS.RecordCount > 0 Then
            TeacherExistByFullName = Success
        Else
            TeacherExistByFullName = Failed
        End If
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function TeacherExistByID(sTeacherID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherID)='" & sTeacherID & "'));") Then
        If vRS.RecordCount > 0 Then
            TeacherExistByID = Success
        Else
            TeacherExistByID = Failed
        End If
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function CreateDefaultRSTeacher(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultRSTeacher = Failed
    
    If ConnectRS(con, vRS, "SELECT * FROM tblTeacher") Then
        CreateDefaultRSTeacher = Success
    End If
End Function
Public Function GetNewTeacherID(ByRef sNewTeacherID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim lNewNumber As Integer
    Dim sSQL As String


    'default
    sNewTeacherID = ""
    GetNewTeacherID = Failed
    sSQL = "SELECT 'TN-' & String$(7-Len(Max(Val(Right([tblTeacher].[TeacherID],7)))+1),'0') & Max(Val(Right([tblTeacher].[TeacherID],7)))+1 AS sNewID" & _
            " FROM tblTeacher;"


    If ConnectRS(con, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = False Then
            sNewTeacherID = "SN-0000001"
            GetNewTeacherID = Success
            GoTo ReleaseAndExit
        End If
    Else
        'fatal error
        GetNewTeacherID = Failed
        GoTo ReleaseAndExit
    End If
    
    sNewTeacherID = (vRS.Fields("snewid"))
    lNewNumber = 0
    While TeacherExistByID(sNewTeacherID) = Success
        If IsNumeric(Right(sNewTeacherID, 7)) = True Then
            lNewNumber = Val(Right(sNewTeacherID, 7)) + 1
        Else
            lNewNumber = 1
        End If
        sNewTeacherID = "TN-" & String$(7 - Len(lNewNumber), "0") & lNewNumber
    Wend
    
    GetNewTeacherID = Success


ReleaseAndExit:
    Set vRS = Nothing
End Function
