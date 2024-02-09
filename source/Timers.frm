VERSION 5.00
Begin VB.Form Timers 
   Caption         =   "Timers"
   ClientHeight    =   1140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Timers"
   ScaleHeight     =   1140
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerPlaySound 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2040
      Top             =   240
   End
   Begin VB.Timer TimerTipCows 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   240
   End
   Begin VB.Timer LogOffDeath 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   480
      Top             =   240
   End
End
Attribute VB_Name = "Timers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LogOffDeath_Timer()

Timers.LogOffDeath.Enabled = False
hook.achooks.Logout

End Sub

Public Sub TimerPlaySound_Timer()

'    If control.chkPlayWavWhenComplete.Checked = True Then
    Hub.PlayExtWavFile (App.Path & "\" & "YMCA.WAV")
'    Hub.PlayExtWavFile (App.Path & "\" & "YMCA.MP3")
'    End If

End Sub

Public Sub TimerTipCows_Timer()

'WriteToChat "(DEBUG) - Starting Count5Min: " & Count5Min, 8
'WriteToChat "(DEBUG) - Villa Searching...", 8
' @house available villa

    If Count5Min = 0 Then
    Count5Min = 1
    
    Else
    
    Count5Min = Count5Min + 1
    End If
    
    If Count5Min >= 30 Then
    hook.achooks.InvokeChatParser ("@house available villa")
    Count5Min = 0
            If control.chkPlayWavWhenComplete.Checked = True Then
            WriteToChat "HeartBeat...", 4
            End If
    End If
    
End Sub


