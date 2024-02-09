VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents achooks As Decal.achooks
Attribute achooks.VB_VarHelpID = -1


'Private Sub AChooks_StatusTextIntercept(ByVal bstrText As String, bEat As Boolean)
'On Error GoTo err_dead
     
'   LOOKING FOR THE DAMN RED TEXT CRAP
            
'    If control.chkEatRedMessages.Checked = True Then
'    WriteToChat "RED TEXT SPAM: " & bstrText, 2
'    bEat = True
'    End If
'    Exit Sub
    
'err_dead:
'    Hub.SendText "err_StatusTextIntercept: " & Err.Description, 1
'End Sub

Private Sub AChooks_ChatTextIntercept(ByVal bstrText As String, _
ByVal lColor As Long, ByVal lTarget As Long, bEat As Boolean)

On Error GoTo err_dead

'    WriteToChat "(DEBUG) - ChatTextIntercept - bstrText:  " & bstrText & "       lColor:" & lColor, 5

    If lColor = 0 And bLoginCompleted = True Then
    
'   -- NO VILLA
'   8:35:57 AM> There are 0 villas available.

    If (InStr(1, bstrText, "villas available", vbTextCompare) <> 0) Then
'    WriteToChat "(DEBUG) - VILLA RESULT MSG!", 4
        
        If (InStr(1, bstrText, "There are 0 villas available", vbTextCompare) <> 0) Then
'        WriteToChat "(DEBUG) - 0 VILLA'S!", 4
        Timers.TimerPlaySound.Enabled = False
    
            If control.chkEatRedMessages.Checked = True Then
            bEat = True
            End If
'        End If
        Else
    
        WriteToChat "OPEN VILLA!!!", 5
    
        If Timers.TimerPlaySound.Enabled = False Then
        Timers.TimerPlaySound.Enabled = True
        End If
        End If
        
            End If
    End If

    Exit Sub
    
err_dead:
    Hub.SendText "err_AChooks_ChatTextIntercept: " & Err.Description, 1
End Sub

