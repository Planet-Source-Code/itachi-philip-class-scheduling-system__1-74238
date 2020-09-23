Attribute VB_Name = "modDBYearLevel"
Option Explicit

Public Const KeyYearLevel = "year"

Public Type tYearLevel
    YearLevelID As Integer
    YearLevelTitle As String
End Type





Public Function YearLevelRecordExisted() As Boolean
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSYearLevel(vRS) = Success Then
        If AnyRecordExisted(vRS) Then
            YearLevelRecordExisted = True
        Else
            YearLevelRecordExisted = False
        End If
    Else
        YearLevelRecordExisted = False
    End If

    Set vRS = Nothing
End Function






Public Function AddYearLevel(newYearLevel As tYearLevel) As TranDBResult
    
    
    Dim vRS As New ADODB.Recordset
    
    If YearLevelExistByID(newYearLevel.YearLevelID) = Success Then
        AddYearLevel = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    If YearLevelExistByTitle(newYearLevel.YearLevelTitle) = Success Then
        AddYearLevel = DuplicateTitle
        GoTo ReleaseAndExit
    End If
    
    
    If CreateDefaultRSYearLevel(vRS) = Success Then
        vRS.AddNew
    
        vRS.Fields("yearlevelid").Value = newYearLevel.YearLevelID
        vRS.Fields("yearleveltitle").Value = newYearLevel.YearLevelTitle

        vRS.Update
        
        AddYearLevel = Success
    Else
        AddYearLevel = Failed
    End If
    
    
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function EditYearLevel(newYearLevel As tYearLevel) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim oldYearLevel As tYearLevel
    
    If GetYearLevelByID(newYearLevel.YearLevelID, oldYearLevel) Then
        If oldYearLevel.YearLevelTitle = newYearLevel.YearLevelTitle Then
            'nothing to proccess here just return success
            EditYearLevel = Success
        Else
            'find title duplicate
            If YearLevelExistByTitle(newYearLevel.YearLevelTitle) = Success Then
                EditYearLevel = DuplicateTitle
                GoTo ReleaseAndExit
            End If
        End If
        
        
        
        If ConnectRS(con, vRS, "SELECT * From tblYearLevel WHERE (((tblYearLevel.YearLevelID)=" & newYearLevel.YearLevelID & "));") Then
            'no duplicates
            'update
            vRS.Fields("yearlevelid").Value = newYearLevel.YearLevelID
            vRS.Fields("yearleveltitle").Value = newYearLevel.YearLevelTitle

            vRS.Update
            
            EditYearLevel = Success
        Else
            EditYearLevel = Failed
        End If
        
    Else
        EditYearLevel = Failed
    End If
ReleaseAndExit:
    Set vRS = Nothing
End Function




Public Function ExecuteDeleteYearLevel(lYearLevelID As Integer) As Boolean
    
      'check if record exist and if it is edited by other user
    If MsgBox("WARNING:" & vbNewLine & _
        "Deleting this YEAR LEVEL entry will affect all other record" & vbNewLine & vbNewLine & _
        "Delete this record anyway?", vbQuestion + vbYesNo) = vbYes Then
            
        If DeleteYearLevel(lYearLevelID) Then
            MsgBox "YEAR LEVEL entry and other related record succesfully deleted.", vbInformation
            ExecuteDeleteYearLevel = True
        Else
            
            MsgBox "Deleting YEAR LEVEL entry went failed.", vbExclamation
            ExecuteDeleteYearLevel = False
        End If
    Else
        ExecuteDeleteYearLevel = False
    End If
End Function

Public Function DeleteYearLevel(lYearLevelID As Integer) As TranDBResult
    
   Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "DELETE * From tblYearLevel WHERE (((tblYearLevel.YearLevelID)=" & lYearLevelID & "));") Then
        DeleteYearLevel = Success
    Else
        DeleteYearLevel = Failed
    End If
    
    Set vRS = Nothing
End Function



Public Function GetYearLevelMoveNext(ByRef vRS As ADODB.Recordset, ByRef vYearlevel As tYearLevel) As Boolean
    
   'assuming that recordset is already connected
    If Not vRS.EOF Then
        
        vYearlevel.YearLevelID = vRS.Fields("yearlevelid").Value
        vYearlevel.YearLevelTitle = vRS.Fields("yearleveltitle").Value
        
        vRS.MoveNext
        GetYearLevelMoveNext = True
    Else
        GetYearLevelMoveNext = False
    End If
    
End Function

Public Function YearLevelExistByID(lYearLevelID As Integer) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "SELECT * From tblYearLevel WHERE (((tblYearLevel.YearLevelID)=" & lYearLevelID & "));") Then
        If AnyRecordExisted(vRS) Then

            YearLevelExistByID = Success
        
        Else
            YearLevelExistByID = Failed
        End If
    Else
        YearLevelExistByID = Failed
    End If
    
    Set vRS = Nothing
   
End Function

Public Function YearLevelExistByTitle(lYearLevelTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "SELECT * From tblYearLevel WHERE (((tblYearLevel.YearLevelTitle)='" & lYearLevelTitle & "'));") Then
        If AnyRecordExisted(vRS) Then

            YearLevelExistByTitle = Success
        
        Else
            YearLevelExistByTitle = Failed
        End If
    Else
        YearLevelExistByTitle = Failed
    End If
    
    Set vRS = Nothing
   
End Function
Public Function GetYearLevelByID(lYearLevelID As Integer, ByRef vYearlevel As tYearLevel) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "SELECT * From tblYearLevel WHERE (((tblYearLevel.YearLevelID)=" & lYearLevelID & "));") Then
        If AnyRecordExisted(vRS) Then
            vRS.MoveFirst
            vYearlevel.YearLevelID = vRS.Fields("yearlevelid").Value
            vYearlevel.YearLevelTitle = vRS.Fields("yearleveltitle").Value

            GetYearLevelByID = Success
        
        Else
            GetYearLevelByID = Failed
        End If
    Else
        GetYearLevelByID = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function GetYearLevelbyTitle(sYearLevelTitle As String, ByRef vYearlevel As tYearLevel) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "SELECT * From tblYearLevel WHERE (((tblYearLevel.YearLevelTitle)='" & sYearLevelTitle & "'));") Then
        If AnyRecordExisted(vRS) Then
            vRS.MoveFirst
            vYearlevel.YearLevelID = vRS.Fields("yearlevelid").Value
            vYearlevel.YearLevelTitle = vRS.Fields("yearleveltitle").Value

            GetYearLevelbyTitle = Success
        
        Else
            GetYearLevelbyTitle = Failed
        End If
    Else
        GetYearLevelbyTitle = Failed
    End If
    
    Set vRS = Nothing
End Function




Public Function CreateDefaultRSYearLevel(ByRef vRS As ADODB.Recordset) As TranDBResult
    If ConnectRS(con, vRS, "SELECT * FROM tblYearLevel") Then

        CreateDefaultRSYearLevel = Success
    Else
        CreateDefaultRSYearLevel = Failed
    End If
End Function

Public Function YearLevelRecordExist() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSYearLevel(vRS) = Success Then
        If AnyRecordExisted(vRS) Then
            YearLevelRecordExist = Success
        Else
            YearLevelRecordExist = Failed
        End If
    Else
        YearLevelRecordExist = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function YLTitleToID(sTitle As String) As Integer
    Select Case UCase(sTitle)
        Case "I"
            YLTitleToID = 1
        Case "II"
            YLTitleToID = 2
        Case "III"
            YLTitleToID = 3
        Case "IV"
            YLTitleToID = 4
        Case Else
            YLTitleToID = 0
    End Select
End Function

Public Function YLIDtoTitle(iLYID As Integer) As String
    Select Case iLYID
        Case 1
            YLIDtoTitle = "I"
        Case 2
            YLIDtoTitle = "II"
        Case 3
            YLIDtoTitle = "III"
        Case 4
            YLIDtoTitle = "IV"
        Case Else
            YLIDtoTitle = "0"
    End Select
End Function
