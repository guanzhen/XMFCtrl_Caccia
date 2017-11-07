'------------------------------------------------------------------------------
'--    ___   _      ___   ___  _____  _      ___     _____            _      --
'--   / __\ /_\    / __\ / __\ \_   \/_\    ( _ )   /__   \___   ___ | |___  --
'--  / /   //_\\  / /   / /     / /\//_\\   / _ \/\   / /\/ _ \ / _ \| / __| --
'-- / /___/  _  \/ /___/ /___/\/ /_/  _  \ | (_>  <  / / | (_) | (_) | \__ \ --
'-- \____/\_/ \_/\____/\____/\____/\_/ \_/  \___/\/  \/   \___/ \___/|_|___/ --
'--                                                                          --
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'--- ON LOAD FRAME ------------------------------------------------------------
'------------------------------------------------------------------------------
#define CONFIG_VERSION  "Configversion V1.00"
#define CONFIG_LEN      16

Dim configLoaded

Sub OnLoadFrame()
  Dim Paramfile
  Dim i, Config, Release, TitleString
  
  'log outputs
  Memory.Set "dispLog", true
  
  Window.Width  = 1040
  Window.Height = 640
  
  ' Create version and title ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  If Not System.Configuration ("Version", Config, "Package") Then
    Release = "missing"
  Else
    Release = "V " & Config.Param(0)
  End If
  If Not System.Configuration( "Description", Config, "Package") Then
    TitleString = "Missing title"
  End If
  TitleString = Config.Param(0) & "   " & Release
  Window.Title = TitleString
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  
  '~~~ Read params ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  'ReadConfigParam
  '~~~ Disable elements for input only ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  'DisableComponents
  'Visual.SerScopeGrid.AddRows 10,10, True, 1
  CreateDebugLogWindow
  Visual.Script("dhxWins").load_cansetup
  Visual.Script("LogGrid").load_messagebox  
  Visual.Script("SCIGrid").load_messagebox  
  Visual.Select("Layer_Main").style.display = "none"
  Visual.Select("Layer_Logs").style.display = "none"  
  Init_MFCommand
  Visual.Script("win").attachEvent "onClose" , Lang.GetRef( "btn_CanConnect" , 1)
End Sub

Sub OnUnloadFrame()
  '~~~ Disconnect COM ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  'disconnectCom
  'Memory.Set "stopRec", True
  '~~~ Save params ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  'If configLoaded = True Then
  '  WriteConfigParam
  'End If
  DebugWindowClose
End Sub


