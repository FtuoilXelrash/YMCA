VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsViewHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents MainView As DecalPlugins.view
Attribute MainView.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Dim Schema As String
      Dim i%
        Schema = ViewSchemaOld()
   
   ' Initialize Decal Controls
   Set MainView = pluginSiteOld.LoadView(Schema)
End Sub

Public Function ViewSchemaOld() As String
On Error GoTo err_dead

    ViewSchemaOld = "<?xml version=""1.0""?>"
'    ViewSchemaOld = ViewSchemaOld & "<view icon=""4151"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
'    ViewSchemaOld = ViewSchemaOld & "<view icon=""4157"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
'    ViewSchemaOld = ViewSchemaOld & "<view icon=""4158"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
' BREAD    ViewSchemaOld = ViewSchemaOld & "<view icon=""4159"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
' cabbage    ViewSchemaOld = ViewSchemaOld & "<view icon=""4160"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
' cheese    ViewSchemaOld = ViewSchemaOld & "<view icon=""4162"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
' chicken    ViewSchemaOld = ViewSchemaOld & "<view icon=""4163"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
' egg    ViewSchemaOld = ViewSchemaOld & "<view icon=""4164"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
' grapes    ViewSchemaOld = ViewSchemaOld & "<view icon=""4166"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
' steak    ViewSchemaOld = ViewSchemaOld & "<view icon=""4168"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
' apple    ViewSchemaOld = ViewSchemaOld & "<view icon=""4169"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
' green sing    ViewSchemaOld = ViewSchemaOld & "<view icon=""4173"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
' cloro lock    ViewSchemaOld = ViewSchemaOld & "<view icon=""4178"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
' key  ViewSchemaOld = ViewSchemaOld & "<view icon=""4190"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
    ViewSchemaOld = ViewSchemaOld & "<view icon=""4191"" title=""YMCA!"" top=""55"" left=""3"" width=""150"" height=""175"">"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Notebook"" name=""YMCA"">"
               
    ViewSchemaOld = ViewSchemaOld & "<page label=""YMCA!"">"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.FixedLayout"" name=""YMCAOpts"" clipped="""">"
                              
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Checkbox"" name=""chkEatRedMessages"" left=""5"" top=""10"" width=""130"" height=""20"" text=""Hide Villa Result Messages"" fontsize=""13"" textcolor=""""/>"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Checkbox"" name=""chkAniIdle"" left=""5"" top=""30"" width=""110"" height=""20"" text=""Anti-Idle"" fontsize=""13""/>"
              
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Checkbox"" name=""chkLogOnDeath"" left=""5"" top=""50"" width=""110"" height=""20"" text=""Log Off On Death"" fontsize=""13""/>"
                         
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Checkbox"" name=""chkPlayWavWhenComplete"" left=""5"" top=""70"" width=""150"" height=""20"" text=""Active HeartBeat"" fontsize=""13""/>"
                         
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.PushButton"" name=""btnStartTips"" top=""100"" left=""40"" width=""50"" height=""16"" facecolor=""64"" fontsize=""14"" textcolor=""65280"" text=""START""/>"

    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</page>"
    
    ViewSchemaOld = ViewSchemaOld & "<page label=""ABOUT"">"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.FixedLayout"" name=""About"" clipped="""">"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutName"" text=""YMCA!"" top=""10"" left=""40"" width=""225"" height=""30"" fontsize=""14"" fontstyle=""bold"" textcolor=""255""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutName2"" text=""For Decal v3"" top=""25"" left=""30"" width=""225"" height=""30"" fontsize=""14"" fontstyle=""bold"" textcolor=""255""/>"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutContact"" text=""Ftuoil Xelrash of Darktide"" top=""55"" left=""5"" width=""210"" height=""20"" fontsize=""14"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutBuildDate"" text=""Version: "" top=""70"" left=""10"" width=""210"" height=""20"" fontsize=""14"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</page>"
       
    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</view>"
    
    Exit Function

err_dead:
    Hub.SendText "err_ViewSchemaOld: " & Err.Description, 1
End Function

' Clear the MainView
Private Sub Class_Terminate()
On Error GoTo err_dead

  Set MainView = Nothing
  Exit Sub
  
err_dead:
    Hub.SendText "err_Class_terminate: " & Err.Description, 1
End Sub
