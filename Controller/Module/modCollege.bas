Attribute VB_Name = "modCollege"
Option Explicit

Public Const keyDepartment = "Col"

Public Type tCollege
    CollegeID As String
    CollegeTitle As String
End Type

Public Function AddCollege(newCollege As tCollege) As TranDBResult

    Dim vRS As New ADODB.Recordset

    If CollegeExistByID(newCollege.CollegeID) = Success Then
        AddCollege = DuplicateID
        GoTo ReleaseAndExit
    End If

    If CollegeExistByTitle(newCollege.CollegeTitle) = Success Then
        AddCollege = DuplicateTitle
        GoTo ReleaseAndExit
    End If
    
    If CreateDefaultRSCollege(vRS) = Success Then
        vRS.AddNew
        vRS.Fields("CollegeID").Value = newCollege.CollegeID
        vRS.Fields("CollegeName").Value = newCollege.CollegeTitle
        vRS.Update
        AddCollege = Success
    Else
        AddCollege = NotConnected
    End If
    
    
    
ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function

Public Function EditCollege(newCollege As tCollege) As TranDBResult
    
    Dim OldCollege As tCollege
    Dim vRS As New ADODB.Recordset

    If GetCollegeByID(newCollege.CollegeID, OldCollege) Then
        If OldCollege.CollegeTitle = newCollege.CollegeTitle Then
            EditCollege = Success
            GoTo ReleaseAndExit
        Else
            If CollegeExistByTitle(newCollege.CollegeTitle) = Success Then
                EditCollege = DuplicateTitle
                GoTo ReleaseAndExit
            End If
        End If
    Else
        EditCollege = InvalidID
        GoTo ReleaseAndExit
    End If
    
    If ConnectRS(con, vRS, "SELECT * From tblCollege WHERE (((tblCollege.CollegeID)='" & newCollege.CollegeID & "'));") Then
        If vRS.RecordCount < 1 Then
            EditCollege = InvalidID
            GoTo ReleaseAndExit
        End If
    End If

        vRS.Fields("CollegeName").Value = newCollege.CollegeTitle
        vRS.Update
            
        EditCollege = Success
        

ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function DeleteCollege(sDepartmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    

    DeleteCollege = Failed
    
    If ConnectRS(con, vRS, "Delete * From tblCollege WHERE (((tblCollege.CollegeID)='" & sDepartmentID & "'));") Then
        DeleteCollege = Success
    Else
        DeleteCollege = Failed
    End If

    Set vRS = Nothing
End Function




Public Function GetCollegeMoveNext(ByRef vRS As ADODB.Recordset, ByRef vCollege As tCollege) As TranDBResult
    If Not vRS.EOF And Not vRS.BOF Then
        vCollege.CollegeID = (vRS.Fields("CollegeID"))
        vCollege.CollegeTitle = (vRS.Fields("CollegeName"))
        vRS.MoveNext
        GetCollegeMoveNext = Success
    Else
        GetCollegeMoveNext = Failed
    End If
    
End Function



Public Function GetCollegeByID(sCollegeID As String, ByRef vCollege As tCollege) As TranDBResult
    
    Dim vRS As New ADODB.Recordset

    If ConnectRS(con, vRS, "SELECT * From tblCollege WHERE (((tblCollege.CollegeID)='" & sCollegeID & "'));") Then
        If AnyRecordExisted(vRS) Then
             vCollege.CollegeID = (vRS.Fields("CollegeID"))
            vCollege.CollegeTitle = (vRS.Fields("CollegeName"))
            
            GetCollegeByID = Success
        Else
            GetCollegeByID = Failed
        End If
    Else
        GetCollegeByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function GetCollegeByTitle(sCollegeTitle As String, ByRef vCollege As tCollege) As TranDBResult

    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT *  FROM tblCollege WHERE (((tblCollege.CollegeName)='" & sCollegeTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            vCollege.CollegeID = (vRS.Fields("CollegeID"))
            vCollege.CollegeTitle = (vRS.Fields("CollegeName"))
            
            GetCollegeByTitle = Success
        Else
            GetCollegeByTitle = Failed
        End If
    Else
        GetCollegeByTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Private Sub ReadFromRecord(ByRef vRS As ADODB.Recordset, ByRef vCollege As tCollege)
    vCollege.CollegeID = vRS.Fields("CollegeID").Value
    vCollege.CollegeTitle = vRS.Fields("CollegeName").Value
End Sub


Public Function CollegeExistByTitle(sCollegeTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblCollege WHERE (((tblCollege.CollegeName)='" & sCollegeTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            CollegeExistByTitle = Success
        Else
            CollegeExistByTitle = Failed
        End If
    Else
        CollegeExistByTitle = Failed
    End If

    Set vRS = Nothing
End Function


Public Function CollegeExistByID(sCollegeID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblCollege WHERE (((tblCollege.CollegeID)='" & sCollegeID & "'));") Then
        If vRS.RecordCount > 0 Then
            CollegeExistByID = Success
        Else
            CollegeExistByID = Failed
        End If
    Else
        CollegeExistByID = Failed
       
    End If

    Set vRS = Nothing
End Function


Public Function CreateDefaultRSCollege(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultRSCollege = Failed
    
    If ConnectRS(con, vRS, "SELECT * FROM tblCollege") Then
        CreateDefaultRSCollege = Success
    End If
End Function

Public Function CollegeRecordExist() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSDepartment(vRS) = Success Then
        
        If AnyRecordExisted(vRS) = True Then
            CollegeRecordExist = Success
        Else
            CollegeRecordExist = Failed
        End If
        
    Else
        CollegeRecordExist = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function GetNewCollegeID(ByRef sNewCollegeID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim NewDNumber As Integer
    
 
    GetNewCollegeID = Failed
    
    sSQL = "SELECT 'Col-' & String$(2-Len(Count(*)+1),'0') & Count(*)+1 AS NewID" & _
            " FROM tblCollege;"
            
    If ConnectRS(con, vRS, sSQL) = True Then
        
        sNewCollegeID = vRS.Fields("NewID").Value
        
        While DepartmentExistByID(sNewCollegeID) = Success
            NewDNumber = Val(Right(sNewCollegeID, 2)) + 1
            sNewCollegeID = "D-" & String(2 - Len(NewDNumber), "0") & NewDNumber
        Wend
        
       GetNewCollegeID = Success
    
    Else
    
        GetNewCollegeID = Failed
    End If
    
    
    
    Set vRS = Nothing

End Function




