Attribute VB_Name = "modDebugPriveleges"
'Injection into a different process is a stable, but not very safe thing to be doing
'so i take no responsibility what you choose to do with this program.

'This was ported from a C++ application.

'Created by Marcin Kleczynski
'marcin@malwarebytes.org

Option Explicit

Private Const SE_DEBUG_NAME As String = "SeDebugPrivilege"
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_QUERY As Long = &H8
Private Const SE_PRIVILEGE_ENABLED As Long = &H2

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

Public Sub GetSeDebugPrivelege()
    LoadPrivilege SE_DEBUG_NAME
End Sub

Public Function LoadPrivilege(ByVal Privilege As String) As Boolean
    On Error GoTo ErrHandler

    Dim hToken&, SEDebugNameValue As LUID, tkp As TOKEN_PRIVILEGES, hProcessHandle&, tkpNewButIgnored As TOKEN_PRIVILEGES, lBuffer&

        hProcessHandle = GetCurrentProcess()
        OpenProcessToken hProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hToken
        LookupPrivilegeValue "", Privilege, SEDebugNameValue

            With tkp
                .PrivilegeCount = 1
                .TheLuid = SEDebugNameValue
                .Attributes = SE_PRIVILEGE_ENABLED
            End With

        AdjustTokenPrivileges hToken, False, tkp, Len(tkp), tkpNewButIgnored, lBuffer
        LoadPrivilege = True
        
    Exit Function
ErrHandler:
    MsgBox "An error occurred retrieving SE_DEBUG_NAME prileges in the LoadPrivelege() function. Note: This program is running without debug priveleges, that may interfere with removing the infection.", vbCritical + vbOKOnly
        Resume Next
End Function

