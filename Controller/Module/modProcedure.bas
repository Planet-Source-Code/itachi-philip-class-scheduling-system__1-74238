Attribute VB_Name = "modProcedure"
Option Explicit

Public Sub ConvertsDaysToInt(cDays As String)
        Select Case cDays
                Case "M"
                    sDays = 1
                Case "T"
                    sDays = 2
                Case "W"
                    sDays = 3
                Case "H"
                    sDays = 4
                Case "F"
                    sDays = 5
                Case "S"
                    sDays = 6
                Case "A"
                    sDays = 7
                Case "MH"
                    sDays = 1
                    sDays1 = 4
                Case "TF"
                    sDays = 2
                    sDays1 = 5
                Case "WS"
                    sDays = 3
                    sDays1 = 6
                Case "MWF"
                    sDays = 1
                    sDays1 = 3
                    sDays2 = 5
                Case "THS"
                    sDays = 2
                    sDays1 = 4
                    sDays2 = 6
        End Select
End Sub
Public Function GetDataSettings()
    CurrentSchoolYear.SchoolYearTitle = GetActiveSchoolYear
    CurrentSemester.Semester = GetActiveSemester
End Function

Public Sub InitAppFailed(sMSG As String)
    MsgBox "Writing to Log: " & sMSG, vbCritical
    End
End Sub

Public Function RandomRGBColor() As Long
    RandomRGBColor = RGB( _
        Int(Rnd() * 256), _
        Int(Rnd() * 256), _
        Int(Rnd() * 256))
End Function
