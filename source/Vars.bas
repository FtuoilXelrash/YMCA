Attribute VB_Name = "Vars"
Option Explicit

Public boolSettingsLoaded       As Boolean
Public boolLoggedIn             As Boolean
Public bLoginCompleted          As Boolean
Global DEBUG_MODE               As Boolean

Public PluginSite               As Decal.IPluginSite2
Public pluginSiteOld            As DecalPlugins.IPluginSite
Public view                     As clsViewHandler
Public control                  As clsControlHandler
Public hook                     As New clsHooks
Public Csf                      As New clsCharStats
Public world                    As New clsWorldFilter

'   REGISTRY
Public Saved                    As String
Public reg                      As RegOp

'   STUFF
Public PlayerGUID               As Long

'   SOUND
Public Const SND_RESOURCE = &H40004
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2

Public Tippin_OK            As Boolean
Public SelectedObjectName   As String
Public CowGUID              As Long
Public Count5Min            As Integer
