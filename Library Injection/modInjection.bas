Attribute VB_Name = "modInjection"
'Injection into a different process is a stable, but not very safe thing to be doing
'so i take no responsibility what you choose to do with this program.

'This was ported from a C++ application.

'Created by Marcin Kleczynski
'marcin@malwarebytes.org

Option Explicit

'Full access to a given process
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

'Just what it sais, infinite time period
Private Const INFINITE = &HFFFFFFFF

'Memory allocation
Private Const MEM_COMMIT = &H1000
Private Const MEM_RELEASE = &H8000
Private Const PAGE_READWRITE = &H4

'Returns current process ID
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'Allows allocation of memory
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long

'Frees the allocated memory
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

'Opens a handle to the process
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'Writes bytes to process memory
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

'Gets the handle a module such as kernel32.dll
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

'Gets a function address
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'Creates the actual remote thread in process
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long

'Waits for thread to finish
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'Closes handles to a process, thread, etc..
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Loads a library into the current process
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'This is the function that injects the library into the given process. If the process
'is our process, this function will call InjectIntoMe()
Public Function InjectLibrary(lPID&, sLibrary$) As Long
    Dim hProcess&, hThread&, lLinkToLibrary&, lSize&, hKernel&

        'If the file does not exist, just exit.
        If Not FileExists(sLibrary) Then
            MsgBox "File does not exist."
            Exit Function
        End If
        
        'If its our process, use different method
        If lPID = GetCurrentProcessId() Then
            'Use alternate method to inject into me
            InjectLibrary = InjectIntoMe(sLibrary)
            
            'Exit the function
            Exit Function
        End If
    
        'Obtain handle to the process
        hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, lPID)
        
        If hProcess = 0 Then
            MsgBox "hProcess returned NULL"
            Exit Function
        End If
        
        'Create the address size and allocate that much memory
        lSize = LenB(StrConv(sLibrary, vbFromUnicode)) + 1
        lLinkToLibrary = VirtualAllocEx(hProcess, 0&, lSize, MEM_COMMIT, PAGE_READWRITE)

        If lLinkToLibrary = 0 Then
            CloseHandle hProcess
        
            MsgBox "lLinkToLibrary failed"
            Exit Function
        End If
        
        'Write the library name to the address space
        If (WriteProcessMemory(hProcess, lLinkToLibrary, ByVal sLibrary, lSize, ByVal 0&) = 0) Then
            CloseHandle hProcess
            If lLinkToLibrary <> 0 Then VirtualFreeEx hProcess, lLinkToLibrary, 0, MEM_RELEASE
            
            MsgBox "WriteProcessMemory failed"
            Exit Function
        End If
        
        'Obtain a handle to the LoadLibrary function from kernel32.dll
        hKernel = GetProcAddress(GetModuleHandle("Kernel32"), "LoadLibraryA")
                
        If hKernel = 0 Then
            CloseHandle hProcess
            If lLinkToLibrary <> 0 Then VirtualFreeEx hProcess, lLinkToLibrary, 0, MEM_RELEASE
            
            MsgBox "hKernel returned NULL"
            Exit Function
        End If
        
        'Create the remote thread in the address space
        hThread = CreateRemoteThread(hProcess, ByVal 0&, 0&, ByVal hKernel, lLinkToLibrary, 0, ByVal 0&)
        
        If hThread = 0 Then
            CloseHandle hKernel
            CloseHandle hProcess
            If lLinkToLibrary <> 0 Then VirtualFreeEx hProcess, lLinkToLibrary, 0, MEM_RELEASE
            
            MsgBox "hThread returned NULL."
            Exit Function
        End If
        
        'Wait for it to complete, the suggested time to wait is 2000 ms, however
        'you may use INFINITE (it is declared)
        WaitForSingleObject hThread, 2000
        
        If lLinkToLibrary <> 0 Then VirtualFreeEx hProcess, lLinkToLibrary, 0, MEM_RELEASE
    
        'Close all open handles
        If hKernel <> 0 Then CloseHandle (hKernel)
        If hThread <> 0 Then CloseHandle (hThread)
        If hProcess <> 0 Then CloseHandle (hProcess)

        InjectLibrary = 1 'Success
End Function

Private Function InjectIntoMe(sLibrary$) As Long
    InjectIntoMe = LoadLibrary(sLibrary)
End Function
