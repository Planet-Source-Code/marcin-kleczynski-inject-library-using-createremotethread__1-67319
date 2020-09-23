VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Injection"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetCurrentProcess 
      Caption         =   "GetCurrentProcessID()"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtPID 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdInject 
      Caption         =   "Inject Library"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "C:\Test.dll"
      Top             =   1800
      Width           =   5055
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      Caption         =   "You can obtain a test DLL from APM (Diamondcs)"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   3525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Location of library goes here."
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2040
   End
   Begin VB.Label lblPID 
      AutoSize        =   -1  'True
      Caption         =   "Process ID goes here."
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblDescription 
      Caption         =   "Uses CreateRemoteThread to create a thread in a remote process and remotely load library."
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Injection into a different process is a stable, but not very safe thing to be doing
'so i take no responsibility what you choose to do with this program.

'This was ported from a C++ application.

'Created by Marcin Kleczynski
'marcin@malwarebytes.org

Option Explicit

Private Sub cmdGetCurrentProcess_Click()
    txtPID.Text = GetCurrentProcessId
End Sub

Private Sub cmdInject_Click()
    Dim lSuccess&

    If txtPID.Text = "" Then
        MsgBox "Need a PID entered."
            Exit Sub
    End If

    'Call the injection code
    lSuccess = InjectLibrary(CLng(txtPID.Text), txtFileName.Text)
    
    If lSuccess > 0 Then
        MsgBox "InjectLibrary succeeded!", vbInformation
    Else
        MsgBox "InjectLibrary failed!", vbCritical
    End If
End Sub

Private Sub Form_Load()
    GetSeDebugPrivelege
End Sub
