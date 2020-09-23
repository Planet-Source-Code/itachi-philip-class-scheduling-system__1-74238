Attribute VB_Name = "modRSSection"
Option Explicit

Public Const keySection = "sect"

Public Type tSection

    SectionID As String
    SectionTitle As String
    
    DepartmentID As String
    YearLevelID As Integer
    
    CreationDate As Date
    CreatedBy As String
    ModifiedDate As Date
    ModifiedBy As String
End Type


Public Function CreateDefaultRSSection(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultRSSection = Failed
    
    If ConnectRS(con, vRS, "SELECT * FROM tblSection") Then
        CreateDefaultRSSection = Success
    End If
End Function




Public Function AddSection(newSection As tSection) As TranDBResult
   
   Dim vRS As New ADODB.Recordset
    

    
    
    
    'check each fields
    If Len(Trim(newSection.SectionID)) < 1 Then
        AddSection = InvalidSectionSectionID
        GoTo ReleaseAndExit
    End If
    
    If Len(Trim(newSection.SectionTitle)) < 1 Then
        AddSection = InvalidSectionSectionTitle
        GoTo ReleaseAndExit
    End If
    
    If DepartmentExistByID(newSection.DepartmentID) <> Success Then
        AddSection = InvalidSectionDepartmentID
        GoTo ReleaseAndExit
    End If

    If YearLevelExistByID(newSection.YearLevelID) <> Success Then
        AddSection = InvalidSectionYearLevelID
        GoTo ReleaseAndExit
    End If
    
    'find duplicate TITLE
    If SectionExistByID(newSection.SectionID) = Success Then
        AddSection = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    'find duplicate TITLE
    If SectionExistByTitle(newSection.SectionTitle) = Success Then
        AddSection = DuplicateTitle
        GoTo ReleaseAndExit
    End If

    
    If CreateDefaultRSSection(vRS) = Success Then
    
        'add new record
        vRS.AddNew
    
        vRS.Fields("sectionid").Value = Trim(newSection.SectionID)
        vRS.Fields("Sectiontitle").Value = Trim(newSection.SectionTitle)
        vRS.Fields("departmentid").Value = Trim(newSection.DepartmentID)
        vRS.Fields("yearlevelid").Value = newSection.YearLevelID
        
        vRS.Fields("CreationDate").Value = newSection.CreationDate
        vRS.Fields("CreatedBy").Value = newSection.CreatedBy
        
        vRS.Update
        
        AddSection = Success
    Else
        AddSection = Failed
    End If
    
    
    
ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function



Public Function EditSection(newSection As tSection) As TranDBResult
    
    Dim oldSection As tSection

    Dim vRS As New ADODB.Recordset
    


    'get old section
    If GetSectionByID(newSection.SectionID, oldSection) = Success Then
                
        If oldSection.SectionTitle <> newSection.SectionTitle Then
            'find duplicate title
            If SectionExistByTitle(newSection.SectionTitle) = Success Then
                EditSection = DuplicateTitle
                'exit function
                GoTo ReleaseAndExit
            End If
 
        End If
    Else
        'department not found
        'exit function
        EditSection = InvalidID
        GoTo ReleaseAndExit
    End If
    

    'find record to edit

    If ConnectRS(con, vRS, "SELECT * From tblSection WHERE (((tblSection.SectionID)='" & newSection.SectionID & "'));") Then
        If vRS.RecordCount < 1 Then
            EditSection = InvalidID
            GoTo ReleaseAndExit
        End If
    End If
        
      
        'vrs'editing
        vRS.MoveFirst
        
        vRS.Fields("Sectiontitle").Value = Trim(newSection.SectionTitle)
        vRS.Fields("departmentid").Value = Trim(newSection.DepartmentID)
        vRS.Fields("yearlevelid").Value = Trim(newSection.YearLevelID)
        
        vRS.Fields("ModifiedDate").Value = newSection.ModifiedDate
        vRS.Fields("ModifiedBy").Value = newSection.ModifiedBy

        vRS.Update
            
        EditSection = Success
        

ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function


Public Function DeleteSection(sSectionID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "Delete * From tblSection WHERE (((tblSection.SectionID)='" & sSectionID & "'));") Then
        DeleteSection = Success
    Else
        DeleteSection = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function



Public Function GetSectionByID(sSectionID As String, ByRef vSection As tSection) As TranDBResult
On Error Resume Next
    Dim vRS As New ADODB.Recordset

    If ConnectRS(con, vRS, "SELECT * From tblSection WHERE (((tblSection.SectionID)='" & sSectionID & "'));") Then
        If AnyRecordExisted(vRS) Then
            'SUCCESS: Record exist
            'get values
            '----------------------------------------------------------------

            vSection.SectionID = (vRS.Fields("sectionid"))
            vSection.SectionTitle = (vRS.Fields("Sectiontitle"))
            vSection.DepartmentID = (vRS.Fields("departmentid"))
            vSection.YearLevelID = (vRS.Fields("yearlevelid"))
            
            vSection.CreationDate = (vRS.Fields("CreationDate"))
            vSection.CreatedBy = (vRS.Fields("CreatedBy"))
            vSection.ModifiedDate = (vRS.Fields("ModifiedDate"))
            vSection.ModifiedBy = (vRS.Fields("ModifiedBy"))

            GetSectionByID = Success
        
        Else

            'FAILED: record does not exist
            GetSectionByID = Failed
        End If
    Else
        GetSectionByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function GetSectionBySectionOfferingID(sSectionOfferingID As String, ByRef vSection As tSection) As TranDBResult
 
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblSection.SectionID, tblSection.SectionTitle, tblSection.DepartmentID, tblSection.YearLevelID, tblSection.CreationDate, tblSection.CreatedBy, tblSection.ModifiedDate, tblSection.ModifiedBy" & _
            " FROM tblSection" & _
            " Where(((tblSection.SectionID) = '" & sSectionOfferingID & "'))" & _
            " GROUP BY tblSection.SectionID, tblSection.SectionTitle, tblSection.DepartmentID, tblSection.YearLevelID, tblSection.CreationDate, tblSection.CreatedBy, tblSection.ModifiedDate, tblSection.ModifiedBy;"

    
    If ConnectRS(con, vRS, sSQL) Then
        If AnyRecordExisted(vRS) Then
            'SUCCESS: Record exist
            'get values
            '----------------------------------------------------------------

            vSection.SectionID = (vRS.Fields("sectionid"))
            vSection.SectionTitle = (vRS.Fields("Sectiontitle"))
            vSection.DepartmentID = (vRS.Fields("departmentid"))
            vSection.YearLevelID = (vRS.Fields("yearlevelid"))
            
            vSection.CreationDate = (vRS.Fields("CreationDate"))
            vSection.CreatedBy = (vRS.Fields("CreatedBy"))
            'vSection.ModifiedDate = (vRS.Fields("ModifiedDate"))
            'vSection.ModifiedBy = (vRS.Fields("ModifiedBy"))

            GetSectionBySectionOfferingID = Success
        
        Else

            'FAILED: record does not exist
            GetSectionBySectionOfferingID = Failed
        End If
    Else
        GetSectionBySectionOfferingID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function



Public Function GetSectionByTitle(sSectionTitle As String, ByRef vSection As tSection) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT *  FROM tblSection WHERE (((tblSection.SectionTitle)='" & sSectionTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            
            vSection.SectionID = (vRS.Fields("sectionid"))
            vSection.SectionTitle = (vRS.Fields("Sectiontitle"))
            vSection.DepartmentID = (vRS.Fields("departmentid"))
            vSection.YearLevelID = (vRS.Fields("yearlevelid"))
            
            vSection.CreationDate = (vRS.Fields("CreationDate"))
            vSection.CreatedBy = (vRS.Fields("CreatedBy"))
            'vSection.ModifiedDate = (vRS.Fields("ModifiedDate"))
            'vSection.ModifiedBy = (vRS.Fields("ModifiedBy"))
            'return success
            GetSectionByTitle = Success
            
        Else
            'return failed
            GetSectionByTitle = Failed
        End If
    Else
        'return failed
        GetSectionByTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function GetSectionByFullTitle(sSectionFullTitle As String, ByRef vSection As tSection) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim splitsSectionFullTitle() As String
    Dim sSectionTitle As String
    
    
    On Error GoTo ErrFound
    
    
    splitsSectionFullTitle = Split(sSectionFullTitle, " - ")
    sSectionTitle = splitsSectionFullTitle(1)
    
    If ConnectRS(con, vRS, "SELECT *  FROM tblSection WHERE (((tblSection.SectionTitle)='" & sSectionTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            
            vSection.SectionID = (vRS.Fields("sectionid"))
            vSection.SectionTitle = (vRS.Fields("Sectiontitle"))
            vSection.DepartmentID = (vRS.Fields("departmentid"))
            vSection.YearLevelID = (vRS.Fields("yearlevelid"))
            
            vSection.CreationDate = (vRS.Fields("CreationDate"))
            vSection.CreatedBy = (vRS.Fields("CreatedBy"))
            'vSection.ModifiedDate = (vRS.Fields("ModifiedDate"))
            'vSection.ModifiedBy = (vRS.Fields("ModifiedBy"))
            'return success
            GetSectionByFullTitle = Success
            
        Else
            'return failed
            GetSectionByFullTitle = Failed
        End If
    Else
        'return failed
        GetSectionByFullTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
    Exit Function
    
ErrFound:
    Set vRS = Nothing
    GetSectionByFullTitle = Failed
End Function


Public Function SectionExistByID(sSectionID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblSection WHERE (((tblSection.SectionID)='" & sSectionID & "'));") Then
        If vRS.RecordCount > 0 Then
            SectionExistByID = Success
        Else
            SectionExistByID = Failed
        End If
        
    Else
        
        SectionExistByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function SectionExistByTitle(sSectionTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "SELECT * From tblSection WHERE (((tblSection.SectionTitle)='" & sSectionTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            SectionExistByTitle = Success
        Else
            SectionExistByTitle = Failed
        End If
    
    Else
    
        SectionExistByTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function
Public Function SectionExistByFullTitle(sSectionFullTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim splitsSectionFullTitle() As String
    Dim sSectionTitle As String
    
    
    On Error GoTo ErrFound
    
    
    splitsSectionFullTitle = Split(sSectionFullTitle, " - ")
    sSectionTitle = splitsSectionFullTitle(1)
    
    If ConnectRS(con, vRS, "SELECT * From tblSection WHERE (((tblSection.SectionTitle)='" & sSectionTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            SectionExistByFullTitle = Success
        Else
            SectionExistByFullTitle = Failed
        End If
    
    Else
    
        SectionExistByFullTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
    Exit Function
    
ErrFound:
    Set vRS = Nothing
    SectionExistByFullTitle = Failed
End Function











Public Function SectionRecordExist() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSSection(vRS) = Success Then
        If AnyRecordExisted(vRS) = True Then
            SectionRecordExist = Success
        Else
            SectionRecordExist = Failed
        End If
    Else
        SectionRecordExist = Failed
    End If
    
    Set vRS = Nothing
End Function


Public Function ExecuteDeleteSection(sSectionID As String) As TranDBResult
    
      'check if record exist and if it is edited by other user
    If MsgBox("WARNING:" & vbNewLine & _
        "Deleting this SECTION entry will affect all other record" & vbNewLine & vbNewLine & _
        "Delete this record anyway?", vbQuestion + vbYesNo) = vbYes Then
            
        If DeleteSection(sSectionID) = Success Then
            MsgBox "SECTION entry and other related record succesfully deleted.", vbInformation
            ExecuteDeleteSection = Success
        Else
            MsgBox "Deleting SECTION entry went failed.", vbExclamation
            ExecuteDeleteSection = Failed
        End If
    Else
        ExecuteDeleteSection = Failed
    End If
End Function



Public Function GetNewSectionID(ByRef sNewSectionID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim NewDNumber  As Long
    
    
    'default
    GetNewSectionID = Failed
    
    sSQL = "SELECT 'SEC-' & String$(6-Len(Count(*)+1),'0') & Count(*)+1 AS NewID" & _
          " FROM tblSection;"

    If ConnectRS(con, vRS, sSQL) = True Then
        
        
        sNewSectionID = vRS.Fields("NewID").Value
        
        While SectionExistByID(sNewSectionID) = Success
            NewDNumber = Val(Right(sNewSectionID, 6)) + 1
            sNewSectionID = "SEC-" & String(6 - Len(NewDNumber), "0") & NewDNumber
        Wend
        
        GetNewSectionID = Success
    
    Else
    
        GetNewSectionID = Failed
    End If
    
    
    
    Set vRS = Nothing

End Function




Public Function GetSectionFullTitle(sSectionID As String, ByRef sFullTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS Title" & _
            " FROM tblYearLevel INNER JOIN tblSection ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " Where(((tblSection.SectionID) = '" & sSectionID & "'))"

    If ConnectRS(con, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            sFullTitle = vRS.Fields("Title").Value
            GetSectionFullTitle = Success
        Else
            GetSectionFullTitle = Failed
        End If
    Else
        GetSectionFullTitle = Failed
    End If
    
    Set vRS = Nothing
End Function

