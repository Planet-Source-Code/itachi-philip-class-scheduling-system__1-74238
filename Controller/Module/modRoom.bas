Attribute VB_Name = "modRoom"
Option Explicit

Public Const KeyRoom = "room"

Public Type vRoom
    Roomname As String
    Building As String
    RoomID As String
    Capacity As Integer
    Department As String
End Type


Public Function AddRoom(newRoom As vRoom) As TranDBResult

    Dim vRS As New ADODB.Recordset
    If RoomExistByID(newRoom.RoomID) = Success Then
        AddRoom = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    If RoomExistByName(newRoom.Roomname) = Success Then
        AddRoom = DuplicateTitle
        GoTo ReleaseAndExit
    End If
            
    
    If CreateDefaultRSRoom(vRS) = Success Then
        vRS.AddNew
        vRS.Fields("RoomID").Value = newRoom.RoomID
        vRS.Fields("Room").Value = newRoom.Roomname
        vRS.Fields("Building").Value = newRoom.Building
        vRS.Fields("Capacity").Value = newRoom.Capacity
        vRS.Update
        AddRoom = Success
    Else
        AddRoom = Failed
    End If
  
ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function

Public Function CreateDefaultRSRoom(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultRSRoom = Failed
    
    If ConnectRS(con, vRS, "SELECT * FROM tblRoom") Then
        CreateDefaultRSRoom = Success
    End If
End Function
Public Function GetRoomByID(sRoomID As String, ByRef tRoom As vRoom) As TranDBResult
On Error Resume Next
    Dim vRS As New ADODB.Recordset

    If ConnectRS(con, vRS, "SELECT * From tblRoom WHERE (((tblRoom.RoomID)='" & sRoomID & "'));") Then
        If AnyRecordExisted(vRS) Then
            tRoom.RoomID = (vRS.Fields("RoomID"))
            tRoom.Roomname = (vRS.Fields("Room"))
            tRoom.Building = (vRS.Fields("Building"))
            tRoom.Capacity = (vRS.Fields("Capacity"))
            GetRoomByID = Success
        
        Else
            GetRoomByID = Failed
        End If
    Else
        GetRoomByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function
Public Function GetRoomDepartmentByID(sRoomID As String, ByRef tRoom As vRoom) As TranDBResult
On Error Resume Next
    Dim vRS As New ADODB.Recordset

    If ConnectRS(con, vRS, "SELECT * From tblRoomDepartment WHERE (((tblRoomDepartment.RoomID)='" & sRoomID & "'));") Then
        If AnyRecordExisted(vRS) Then
            tRoom.Department = (vRS.Fields("DepartmentID"))
            GetRoomDepartmentByID = Success
        Else
            GetRoomDepartmentByID = Failed
        End If
    Else
        GetRoomDepartmentByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function RoomExistByID(sRoomID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblRoom WHERE (((tblRoom.RoomID)='" & sRoomID & "'));") Then
        If vRS.RecordCount > 0 Then
            RoomExistByID = Success
        Else
            RoomExistByID = Failed
        End If
        
    Else
        
        RoomExistByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function
Public Function EditRoom(newRoom As vRoom) As TranDBResult
    Dim oldRoom As vRoom

    Dim vRS As New ADODB.Recordset
    
    If GetRoomByID(newRoom.RoomID, oldRoom) = Success Then
                
        If oldRoom.Roomname <> newRoom.Roomname Then
            EditRoom = Success
            GoTo ReleaseAndExit
        End If
    Else
        EditRoom = InvalidID
        GoTo ReleaseAndExit
    End If
    
    If ConnectRS(con, vRS, "SELECT * From tblRoom WHERE (((tblroom.roomID)='" & newRoom.RoomID & "'));") Then
        If vRS.RecordCount < 1 Then
            EditRoom = InvalidID
            GoTo ReleaseAndExit
        End If
    End If
        
        vRS.Fields("Room").Value = newRoom.Roomname
        vRS.Fields("Building").Value = newRoom.Building
        vRS.Fields("Capacity").Value = newRoom.Capacity
        vRS.Update
            
        EditRoom = Success
        

ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function
Public Function RoomExistByName(sRoomname As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblroom WHERE (((tblroom.room)='" & sRoomname & "'));") Then
        If vRS.RecordCount > 0 Then
            RoomExistByName = Success
        Else
            RoomExistByName = Failed
        End If
    Else
        RoomExistByName = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function GetNewRoomID(ByRef sNewID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim NewDNumber As Long
    
    sSQL = "SELECT 'ROOM-' & String$(2-Len(Count(*)+1),'0') & Count(*)+1 AS NewID" & _
            " FROM tblRoom;"

    If ConnectRS(con, vRS, sSQL) = True Then
        sNewID = vRS.Fields("NewID").Value
        While SubjectExistByID(sNewID) = Success
            NewDNumber = Val(Right(sNewID, 2)) + 1
            sNewID = "ROOM-" & String(2 - Len(NewDNumber), "0") & NewDNumber
        Wend
        GetNewRoomID = Success
    Else
        GetNewRoomID = Failed
    End If

    Set vRS = Nothing
End Function

