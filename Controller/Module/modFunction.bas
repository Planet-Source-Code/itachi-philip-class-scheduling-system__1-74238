Attribute VB_Name = "modFunction"
Option Explicit

Public Function ConvertTimeIN(ByVal cTimeIn As String) As Integer
    Dim vTime As Date
        vTime = Format$(cTimeIn, "hh:mm AM/PM")
    Select Case vTime
        Case "6:00 AM"
            ConvertTimeIN = 1
        Case "6:30 AM"
            ConvertTimeIN = 2
        Case "7:00 AM"
            ConvertTimeIN = 3
        Case "7:30 AM"
            ConvertTimeIN = 4
        Case "8:00 AM"
            ConvertTimeIN = 5
        Case "8:30 AM"
            ConvertTimeIN = 6
        Case "9:00 AM"
            ConvertTimeIN = 7
        Case "9:30 AM"
            ConvertTimeIN = 8
        Case "10:00 AM"
            ConvertTimeIN = 9
        Case "10:30 AM"
            ConvertTimeIN = 10
        Case "11:00 AM"
            ConvertTimeIN = 11
        Case "11:30 AM"
            ConvertTimeIN = 12
        Case "12:00 PM"
            ConvertTimeIN = 13
        Case "12:30 PM"
            ConvertTimeIN = 14
        Case "1:00 PM"
            ConvertTimeIN = 15
        Case "1:30 PM"
            ConvertTimeIN = 16
        Case "2:00 PM"
            ConvertTimeIN = 17
        Case "2:30 PM"
            ConvertTimeIN = 18
        Case "3:00 PM"
            ConvertTimeIN = 19
        Case "3:30 PM"
            ConvertTimeIN = 20
        Case "4:00 PM"
            ConvertTimeIN = 21
        Case "4:30 PM"
            ConvertTimeIN = 22
        Case "5:00 PM"
            ConvertTimeIN = 23
        Case "5:30 PM"
            ConvertTimeIN = 24
        Case "6:00 PM"
            ConvertTimeIN = 25
        Case "6:30 PM"
            ConvertTimeIN = 26
        Case "7:00 PM"
            ConvertTimeIN = 27
        Case "7:30 PM"
            ConvertTimeIN = 28
        Case "8:00 PM"
            ConvertTimeIN = 29
        Case "8:30 PM"
            ConvertTimeIN = 30
    End Select
End Function

Public Function ConvertTimeOut(ByVal cTimeOut As String) As Integer
    Dim vTime As Date
        vTime = Format$(cTimeOut, "hh:mm AM/PM")
    Select Case vTime
        Case "6:00 AM"
            ConvertTimeOut = 1
        Case "6:30 AM"
            ConvertTimeOut = 2
        Case "7:00 AM"
            ConvertTimeOut = 3
        Case "7:30 AM"
            ConvertTimeOut = 4
        Case "8:00 AM"
            ConvertTimeOut = 5
        Case "8:30 AM"
            ConvertTimeOut = 6
        Case "9:00 AM"
            ConvertTimeOut = 7
        Case "9:30 AM"
            ConvertTimeOut = 8
        Case "10:00 AM"
            ConvertTimeOut = 9
        Case "10:30 AM"
            ConvertTimeOut = 10
        Case "11:00 AM"
             ConvertTimeOut = 11
        Case "11:30 AM"
            ConvertTimeOut = 12
        Case "12:00 PM"
            ConvertTimeOut = 13
        Case "12:30 PM"
            ConvertTimeOut = 14
        Case "1:00 PM"
            ConvertTimeOut = 15
        Case "1:30 PM"
            ConvertTimeOut = 16
        Case "2:00 PM"
            ConvertTimeOut = 17
        Case "2:30 PM"
            ConvertTimeOut = 18
        Case "3:00 PM"
            ConvertTimeOut = 19
        Case "3:30 PM"
            ConvertTimeOut = 20
        Case "4:00 PM"
            ConvertTimeOut = 21
        Case "4:30 PM"
            ConvertTimeOut = 22
        Case "5:00 PM"
            ConvertTimeOut = 23
        Case "5:30 PM"
            ConvertTimeOut = 24
        Case "6:00 PM"
            ConvertTimeOut = 25
        Case "6:30 PM"
            ConvertTimeOut = 26
        Case "7:00 PM"
            ConvertTimeOut = 27
        Case "7:30 PM"
            ConvertTimeOut = 28
        Case "8:00 PM"
            ConvertTimeOut = 29
        Case "8:30 PM"
            ConvertTimeOut = 30
    End Select
End Function

Public Function AnyRecordExisted(ByRef vRS As ADODB.Recordset) As Boolean
    If vRS.State = adStateClosed Then
        AnyRecordExisted = False
        Exit Function
    End If
    
    
    vRS.Requery
    
    If (vRS.BOF = True) And (vRS.EOF = True) Then
        AnyRecordExisted = False
    Else
        On Error GoTo errh
        vRS.MoveFirst
        AnyRecordExisted = True
    End If

    Exit Function
    '--------------------------
    
errh:
    AnyRecordExisted = False
End Function
                
Public Function CatchError(sModuleName As String, sRoutineName As String, sDetail As String)
    MsgBox sModuleName & " - " & sRoutineName & " - " & sDetail
End Function
Public Function RecordNoMatch(ByRef vRS As ADODB.Recordset) As Boolean
On Error GoTo errh:

    RecordNoMatch = (vRS.BOF = True Or vRS.EOF = True)

    Exit Function
    
errh:
    RecordNoMatch = False
    
End Function
Public Function getRecordCount(ByRef vRS As ADODB.Recordset) As Long
    If AnyRecordExisted(vRS) Then
        vRS.Requery
        vRS.MoveLast
        getRecordCount = vRS.RecordCount
    Else
        getRecordCount = 0
    End If
End Function

Public Function RSMoveFirst(ByRef vRS As ADODB.Recordset) As Boolean
    If AnyRecordExisted(vRS) Then
        vRS.MoveFirst
        RSMoveFirst = True
    Else
        RSMoveFirst = False
    End If
End Function
Public Function HLTxt(ByRef txt As Object)
On Error Resume Next
    txt.SelStart = 0
    txt.SelLength = Len(txt)
    txt.SetFocus
End Function

Function GetINI(strMain As String, strSub As String) As String
  Dim strBuffer As String
  Dim lngLen As Long
  Dim lngRet As Long
  
  strBuffer = Space(100)
  lngLen = Len(strBuffer)
  lngRet = GetPrivateProfileString(strMain, strSub, vbNullString, strBuffer, lngLen, App.Path & "\config.txt")
  GetINI = Left(strBuffer, lngRet)
End Function

Public Sub SetINI(strMain As String, strSub As String, strvalue As String)
  WritePrivateProfileString strMain, strSub, strvalue, App.Path & "\config.txt"
End Sub

Public Function SortLV(ByRef lv As ListView, Optional HeaderIndex As Integer = 0, Optional newSortOrder As ListSortOrderConstants = lvwAscending, Optional AutoOrder As Boolean = True)
    
    Dim lvHeader As ColumnHeader
    
    If AutoOrder = True Then
        If lv.SortOrder = lvwAscending Then
           lv.SortOrder = lvwDescending
        Else
           lv.SortOrder = lvwAscending
        End If
    Else
        lv.SortOrder = newSortOrder
    End If
    
    If HeaderIndex > lv.ColumnHeaders.count - 1 Then
        HeaderIndex = 0
    End If
    
    lv.SortKey = HeaderIndex
    lv.Sorted = True
    lv.Refresh
    
    For Each lvHeader In lv.ColumnHeaders
        lvHeader.Icon = 0
    Next
    
    On Error Resume Next
    lv.ColumnHeaders(HeaderIndex + 1).Icon = lv.SortOrder + 1
End Function

Public Function UnSortLV(ByRef lv As ListView)
    
    Dim lvHeader As ColumnHeader
    
    lv.Sorted = False
    
    For Each lvHeader In lv.ColumnHeaders
        lvHeader.Icon = 0
    Next
End Function
Public Function CenterForm(ByRef Frm As Form)
    Frm.Move (Screen.Width - Frm.Width) / 2, (Screen.Height - Frm.Height) / 2
End Function
Public Function FillRecordToList(ByRef vRS As ADODB.Recordset, ByRef lv As ListView, sTableKey As String, Optional RecStartPos As Long = 0, Optional LimitCount As Long = 100, Optional WithID As Boolean = True, Optional WithIcon As Boolean = False)

    Dim i As Long
    Dim newColumnWidth As Integer
    Dim LimitCounter As Long
    Dim sCell As String
    Dim oldScaleMode As ScaleModeConstants
    
On Error Resume Next
    
    'minum fields must be 2
    If vRS.Fields.count < 2 Then Exit Function
    
    'get old scale mode
    'oldScaleMode = lv.Container.ScaleMode

    'lv.Container.ScaleMode = vbTwips
    lv.ListItems.Clear
    
    
    If AnyRecordExisted(vRS) Then
        
        'create column headers
        For i = lv.ColumnHeaders.count To vRS.Fields.count - 2
            lv.ColumnHeaders.Add
        Next
        
        'set items
        vRS.Requery
        vRS.Move RecStartPos
        
        LimitCounter = 0
        While Not vRS.EOF
            If WithIcon = True Then
                If WithID = True Then
                    lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), vRS.Fields(1).Value, 1, 1
                Else
                    lv.ListItems.Add , , vRS.Fields(0).Value, 1, 1
                End If
            Else
                If WithID = True Then
                    lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), vRS.Fields(1).Value
                Else
                    lv.ListItems.Add , , vRS.Fields(0).Value
                End If
            End If
            'add sub items
            For i = 2 To vRS.Fields.count - 1
           
                lv.ListItems(lv.ListItems.count).SubItems(i - 1) = (vRS.Fields(i))
            Next
            
            vRS.MoveNext
            
            LimitCounter = LimitCounter + 1
            
            If LimitCounter >= LimitCount Then
                GoTo tagExitSub
            End If
        Wend
    End If
    
tagExitSub:
    'lv.Container.ScaleMode = oldScaleMode
End Function

Public Function GetLVKey(lvListItem As ListItem) As String
On Error GoTo errh:
    GetLVKey = Right(lvListItem.Key, Len(lvListItem.Key) - 4)
    Exit Function
errh:
    GetLVKey = ""
End Function
Public Function SetLVKey(sID As String, sTableKey As String) As String
    SetLVKey = Left(sTableKey, 4) & sID
End Function
Public Function CheckTextBox(ByRef txt As Object, Optional sMSG As String = "TextBox", Optional ShowMSG As Boolean = True, Optional MinimumChar As Integer = 1) As Boolean
On Error Resume Next
    If Len(Trim(txt.Text)) < MinimumChar Then
        
        If ShowMSG Then
            MsgBox sMSG, vbExclamation
        End If
        
        txt.Text = ""
        txt.SetFocus
        
        CheckTextBox = False
    Else
        CheckTextBox = True
    End If
End Function


Public Function GetLVSelectedCount(ByRef lv As ListView) As Integer
    Dim i As Integer
    Dim iSelectedCount As Integer
    
    'default
    GetLVSelectedCount = 0
    
    'check if there is a record in the list
    If lv.ListItems.count < 1 Then Exit Function
    
    
    iSelectedCount = 0
    For i = 1 To lv.ListItems.count
        If lv.ListItems(i).Selected = True And Len(GetLVKey(lv.ListItems(i))) > 0 Then
            iSelectedCount = iSelectedCount + 1
        End If
    Next
    
    'return
    GetLVSelectedCount = iSelectedCount
End Function
