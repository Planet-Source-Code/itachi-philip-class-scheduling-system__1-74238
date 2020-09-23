Attribute VB_Name = "modVariables"
Option Explicit

Public sDays, sDays1, sDays2 As Integer
Public con As New ADODB.Connection

Public Const KeyStudent = "stud"

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public DBPath As String

Public Type CurrentUser
    USERNAME As String
    Fullname As String
End Type

Public Enum TranDBResult
    Success = 1
    NoResult = 0
    Failed = -99
    
    NotConnected = -1
    NoRecordExist = -2
    
    InvalidID = -11
    InvalidTitle = -12
    
    DuplicateID = -21
    DuplicateTitle = -22
    
    'teacher invalid result
    InvalidTeacherTitle = -201
    InvalidTeacherPassword = -202
    InvalidTeacherFirstName = -203
    InvalidTeacherMiddleName = -204
    InvalidTeacherLastName = -205
    InvalidTeacherContactNumber = -206
    InvalidTeacherAddress = -207
    
    
    'section invalid result
    DuplicateTeacherID = -301
    
    'student
    DuplicateFullName = -402
    
    'section
    InvalidSectionSectionID = -501
    InvalidSectionDepartmentID = -502
    InvalidSectionTeacherID = -503
    InvalidSectionSectionTitle = -504
    InvalidSectionYearLevelID = -505
    InvalidSectionRoomNumber = -506
    InvalidSectionMinAveGrade = -507
    InvalidSectionMaxAveGrade = -508
    InvalidSectionMaxStudentCount = -509
    'enrolment
    EnrolmentDuplicateEntryWithInYear = -591
    EnrolmentSchoolYearNotFound = -592
    EnrolmentStudentIDNotFound = -593
    EnrolmentSectionIDNotFound = -594
    EnrolmentInvalidAveGrade = 595
    
    'subject
    InvalidSubjectSubjectID = -701
    InvalidSubjectSubjectTitle = -702
    InvalidSubjectDepartmentID = -703
    InvalidSubjectYearLevelID = -704
    InvalidSubjectDescription = -705
    
    'grade
    InvalidGradeID = -801
    InvalidGradeEnrolmentID = -802
    InvalidGradeSubjectID = -803
    InvalidGradeGradeValue = -804
    
    'user
    UserNotExist = -901
    UserDuplicate = -902
    
    'log
    AlreadyLogIn = -1001
    SuccessIn = 1001
    
    DuplicateLoginName = -1101
    
End Enum


Public Function CheckInstallation() As Boolean
    Dim vRS As New ADODB.Recordset
    If ConnectRS(con, vRS, "Select * FROM tblExpiration") = True Then
         If AnyRecordExisted(vRS) = True Then
            
            CheckInstallation = True
            
        Else
            CheckInstallation = False
        End If
    Else
        CheckInstallation = False
    End If
    Set vRS = Nothing
End Function
