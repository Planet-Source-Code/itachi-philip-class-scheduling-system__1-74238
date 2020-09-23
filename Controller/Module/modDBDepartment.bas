Attribute VB_Name = "modDBDepartment"
Option Explicit

Public Const keyDepartment = "dept"

Public Type tDepartment
    DepartmentID As String
    DepartmentTitle As String
    CollegeID As String
End Type

Public Function AddDepartment(newDepartment As tDepartment) As TranDBResult
    Dim vRS As New ADODB.Recordset

    If DepartmentExistByID(newDepartment.DepartmentID) = Success Then
        AddDepartment = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    If DepartmentExistByTitle(newDepartment.DepartmentTitle) = Success Then
        AddDepartment = DuplicateTitle
        GoTo ReleaseAndExit
    End If
    
    If CreateDefaultRSDepartment(vRS) = Success Then
        'add new record
        vRS.AddNew
        vRS.Fields("DepartmentID").Value = newDepartment.DepartmentID
        vRS.Fields("DepartmentTitle").Value = newDepartment.DepartmentTitle
        vRS.Fields("CollegeID").Value = newDepartment.CollegeID
       vRS.Update
        AddDepartment = Success
    Else
        AddDepartment = NotConnected
    End If

ReleaseAndExit:

    Set vRS = Nothing
End Function

Public Function EditDepartment(newDepartment As tDepartment) As TranDBResult

    Dim OldDepartment As tDepartment

    Dim vRS As New ADODB.Recordset

    If GetDepartmentByID(newDepartment.DepartmentID, OldDepartment) Then
        If OldDepartment.DepartmentTitle = newDepartment.DepartmentTitle Then
            EditDepartment = Success
            GoTo ReleaseAndExit
        Else
            If DepartmentExistByTitle(newDepartment.DepartmentTitle) = Success Then
                EditDepartment = DuplicateTitle
                GoTo ReleaseAndExit
            End If
        End If
    Else
        EditDepartment = InvalidID
        GoTo ReleaseAndExit
    End If
    
    If ConnectRS(con, vRS, "SELECT * From tblDepartment WHERE (((tblDepartment.DepartmentID)='" & newDepartment.DepartmentID & "'));") Then
        If vRS.RecordCount < 1 Then
            EditDepartment = InvalidID
            GoTo ReleaseAndExit
        End If
    End If
    
        vRS.Fields("Departmenttitle").Value = newDepartment.DepartmentTitle
        vRS.Fields("CollegeID").Value = newDepartment.CollegeID
        vRS.Update
            
        EditDepartment = Success
        

ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function









Public Function DeleteDepartment(sDepartmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    
    'default
    DeleteDepartment = Failed
    
    If ConnectRS(con, vRS, "Delete * From tblDepartment WHERE (((tblDepartment.DepartmentID)='" & sDepartmentID & "'));") Then
        DeleteDepartment = Success
    Else
        DeleteDepartment = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function




Public Function GetDepartmentMoveNext(ByRef vRS As ADODB.Recordset, ByRef vDepartment As tDepartment) As TranDBResult

    'asuming that vRS is already connected
    If Not vRS.EOF And Not vRS.BOF Then
    
        'SUCCESS: Record exist
        'get values
        '----------------------------------------------------------------
        vDepartment.DepartmentID = (vRS.Fields("DepartmentID"))
        vDepartment.DepartmentTitle = (vRS.Fields("DepartmentTitle"))
        vDepartment.CollegeID = (vRS.Fields("CollegeID"))
        'move to the next record
        vRS.MoveNext
        'return true
        GetDepartmentMoveNext = Success
    Else
        GetDepartmentMoveNext = Failed
    End If
    
End Function



Public Function GetDepartmentByID(sDepartmentID As String, ByRef vDepartment As tDepartment) As TranDBResult
    
    Dim vRS As New ADODB.Recordset

    If ConnectRS(con, vRS, "SELECT * From tblDepartment WHERE (((tblDepartment.DepartmentID)='" & sDepartmentID & "'));") Then
        If AnyRecordExisted(vRS) Then
            'SUCCESS: Record exist
            'get values
            '----------------------------------------------------------------
            vDepartment.DepartmentID = (vRS.Fields("DepartmentID"))
            vDepartment.DepartmentTitle = (vRS.Fields("DepartmentTitle"))
            
            GetDepartmentByID = Success
        
        Else

            'FAILED: record does not exist
            GetDepartmentByID = Failed
        End If
    Else
        GetDepartmentByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function GetDepartmentByTitle(sDepartmentTitle As String, ByRef vDepartment As tDepartment) As TranDBResult

    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT *  FROM tblDepartment WHERE (((tblDepartment.DepartmentTitle)='" & sDepartmentTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            vDepartment.DepartmentID = (vRS.Fields("DepartmentID"))
            vDepartment.DepartmentTitle = (vRS.Fields("DepartmentTitle"))
            vDepartment.CollegeID = (vRS.Fields("CollegeID"))
            GetDepartmentByTitle = Success
        Else
            GetDepartmentByTitle = Failed
        End If
    Else
        GetDepartmentByTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Private Sub ReadFromRecord(ByRef vRS As ADODB.Recordset, ByRef vDepartment As tDepartment)
    
    vDepartment.DepartmentID = vRS.Fields("Departmentid").Value
    vDepartment.DepartmentTitle = vRS.Fields("Departmenttitle").Value
    vDepartment.CollegeID = (vRS.Fields("CollegeID"))
End Sub


Public Function DepartmentExistByTitle(sDepartmentTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblDepartment WHERE (((tblDepartment.DepartmentTitle)='" & sDepartmentTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            DepartmentExistByTitle = Success
        Else
            DepartmentExistByTitle = Failed
        End If
    Else
        DepartmentExistByTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function DepartmentExistByID(sDepartmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblDepartment WHERE (((tblDepartment.DepartmentID)='" & sDepartmentID & "'));") Then
        If vRS.RecordCount > 0 Then
            DepartmentExistByID = Success
        Else
            DepartmentExistByID = Failed
        End If
    Else
        DepartmentExistByID = Failed
       
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function CreateDefaultRSDepartment(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultRSDepartment = Failed
    
    If ConnectRS(con, vRS, "SELECT * FROM tblDepartment") Then
        CreateDefaultRSDepartment = Success
    End If
End Function

Public Function DepartmentRecordExist() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSDepartment(vRS) = Success Then
        
        If AnyRecordExisted(vRS) = True Then
            DepartmentRecordExist = Success
        Else
            DepartmentRecordExist = Failed
        End If
        
    Else
        DepartmentRecordExist = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function GetNewDepartmentID(ByRef sNewDepartmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim NewDNumber As Integer
    
    
    'default
    GetNewDepartmentID = Failed
    
    sSQL = "SELECT 'D-' & String$(2-Len(Count(*)+1),'0') & Count(*)+1 AS NewID" & _
            " FROM tblDepartment;"
            
    If ConnectRS(con, vRS, sSQL) = True Then
        
        sNewDepartmentID = vRS.Fields("NewID").Value
        
        While DepartmentExistByID(sNewDepartmentID) = Success
            NewDNumber = Val(Right(sNewDepartmentID, 2)) + 1
            sNewDepartmentID = "D-" & String(2 - Len(NewDNumber), "0") & NewDNumber
        Wend
        
        GetNewDepartmentID = Success
    
    Else
    
        GetNewDepartmentID = Failed
    End If
    
    
    
    Set vRS = Nothing

End Function


