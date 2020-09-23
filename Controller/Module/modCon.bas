Attribute VB_Name = "modCon"
Option Explicit
Public CurrentSchoolYear As tSchoolYear
Public CurrentUser As CurrentUser
Public CurrentSemester As tSemester

Sub Main()
    InitCommonControls
    
     DBPath = GetINI("Configuration", "Path")      'get path from file
    If Trim(DBPath) = "" Or IsNull(DBPath) Then
JumpHere:
      frmLocate.Show 1                            'browse database
    End If
        
        If OpenDB = vbRetry Then GoTo JumpHere
        
        GetDataSettings
        
        mdiController.Show
        
        'If CheckSetup < 0 Then
        '    Exit Sub
        'End If
        
End Sub


Private Function CheckSetup() As Variant
    If CheckInstallation = False Then
        frmStart.SetTrial
        CheckSetup = -1
    Else
        CheckSetup = 1
        frmStart.Show
    End If
End Function

Public Function OpenDB() As Integer
  Dim isOpen      As Boolean
  Dim ANS         As VbMsgBoxResult
  
  isOpen = False
  On Error GoTo errhandler
    
    
    
  Do Until isOpen = True
    con.CursorLocation = adUseServer
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
    & "Data Source=" & DBPath & ";" _
    & "Persist Security Info=False;" _
    & "Jet OLEDB:Database Password=bagares"
    
    isOpen = True
  Loop
  OpenDB = isOpen
    
  Exit Function
errhandler:
  ANS = MsgBox("Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, _
  vbCritical + vbRetryCancel)
  If ANS = vbCancel Then
    OpenDB = vbCancel
    End
  ElseIf ANS = vbRetry Then
    OpenDB = vbRetry
  End If
End Function
Public Sub CloseDB()
    con.Close
    Set con = Nothing
End Sub

Public Function ConnectRS(ByRef vDB As ADODB.Connection, ByRef vRS As ADODB.Recordset, sSQL As String, Optional ShowMSG As Boolean = True) As Boolean
    
On Error GoTo errh

    Set vRS = Nothing
    Set vRS = New ADODB.Recordset
  
    vRS.Open sSQL, vDB, adOpenStatic, adLockOptimistic
    ConnectRS = True
    Exit Function
errh:
    If ShowMSG = True Then
        
        Clipboard.SetText sSQL

        MsgBox "FATAL ERROR" & vbNewLine & "Connection String: " & sSQL & vbNewLine & "Error: " & err.Description
    End If
    ConnectRS = False
End Function

