VERSION 5.00
Begin VB.Form frmShutdown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shutdown & Logoff"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogoff 
      Caption         =   "Log Off"
      Height          =   435
      Left            =   1800
      TabIndex        =   3
      Top             =   705
      Width           =   1560
   End
   Begin VB.CommandButton cmdShutdown 
      Caption         =   "Shutdown"
      Height          =   435
      Left            =   150
      TabIndex        =   2
      Top             =   705
      Width           =   1560
   End
   Begin VB.CommandButton cmdForceLogoff 
      Caption         =   "Force Log Off"
      Height          =   435
      Left            =   1800
      TabIndex        =   1
      Top             =   165
      Width           =   1560
   End
   Begin VB.CommandButton cmdForceShutdown 
      Caption         =   "Force Shut Down"
      Height          =   435
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   1560
   End
End
Attribute VB_Name = "frmShutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogoff_Click()

Call LogOff

End Sub

Private Sub cmdForceShutdown_Click()

Call ForceShutDown

End Sub

Private Sub cmdForceLogoff_Click()

Call ForceLogOff

End Sub

Private Sub cmdShutdown_Click()

Call ShutDown

End Sub


Private Sub Form_Load()
'********************************************************************
'* When the project starts, check the operating system used by
'* calling the GetVersion function.
'********************************************************************
Dim lngVersion As Long

lngVersion = GetVersion()

If ((lngVersion And &H80000000) = 0) Then
   glngWhichWindows32 = mlngWindowsNT
   MsgBox "Running Windows NT"
Else
   glngWhichWindows32 = mlngWindows95
   MsgBox "Running Windows 95 or 98"
End If

End Sub

