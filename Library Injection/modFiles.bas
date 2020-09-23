Attribute VB_Name = "modFiles"
'Injection into a different process is a stable, but not very safe thing to be doing
'so i take no responsibility what you choose to do with this program.

'This was ported from a C++ application.

'Created by Marcin Kleczynski
'marcin@malwarebytes.org

Option Explicit

'Does the file exist, if so, report true
Public Function FileExists(sFile$) As Boolean
    If Trim(sFile) = vbNullString Then Exit Function
    
        FileExists = IIf(Dir(sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, True, False)
End Function

'Used as a better stripper function
Public Function TrimNull$(sToTrim$)
    If InStr(sToTrim, Chr(0)) > 0 Then
        TrimNull = Left(sToTrim, InStr(sToTrim, Chr(0)) - 1)
    Else
        TrimNull = sToTrim
    End If
End Function
