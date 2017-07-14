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

  Memory.Set "dispLog", false
  
  Window.Width  = 1040
  Window.Height = 680
  
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
'  Visual.SerScopeGrid.AddRows 10,10, True, 1
  
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


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~ Initialize serial scope ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'-------------------------------------------------------------------------------
'--- Initialize components -----------------------------------------------------
'-------------------------------------------------------------------------------



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~ Init HEX EDIT CONTROL ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'--- Disable components --------------------------------------------------------
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub DisableComponents
End Sub 

Sub disableButtons
End Sub

Sub enableButtons
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'--- Enable COM-Interface ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub EnableInterface (enable)
  If enable = True Then
 
    
  Else

  End If
End Sub  



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'--- get reader parameter ------------------------------------------------------
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub OnClick_BUTTON_SET_DEVID_DEFAULT(Reason)

  
End Sub  


'-- Check line numbers of config file ------------------------------------------------
Function CheckConfLen (sPath)
  Dim oFso, oReg, sData
  Const ForReading = 1
  Set oReg = New RegExp
  Set oFso = CreateObject("Scripting.FileSystemObject")
  sData = oFso.OpenTextFile(sPath, ForReading).ReadAll
  With oReg
      .Pattern = "\r\n" 
      .Global = True
       CheckConfLen = .Execute(sData).Count + 1
  End With
  Set oFso = Nothing
End Function



Sub ReadConfigParam
  Dim Paramfile, RunTime, Str, Val, cnt, sPath, version
  Dim dummy, lineCnt
  Dim confOk
  
  confOk = True
  sPath = System.Environment.Path & "Config\MF_Config.txt"
  If Not File.FileExists(sPath) = True Then
    System.MessageBox  "Configuration file does not exist , tool-settings relocated to default !", "New configuration necessary" , MB_ICONEXCLAMATION
  Else
    If Not CheckConfLen (sPath) = $(CONFIG_LEN) Then
      System.MessageBox  "Configuration file corrupt , tool-settings relocated to default !", "New configuration necessary" , MB_ICONEXCLAMATION
    Else
      Set Paramfile = File.Open( sPath, "rt")
      version = Paramfile.ReadLine
      If Not version = $(CONFIG_VERSION) Then
        System.MessageBox  "New configuration order detected, tool-settings relocated to default !", "New configuration necessary" , MB_ICONEXCLAMATION
      Else
        ' COM params ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        for cnt = 1 to 5
          Visual.Script("SysParaGrid").setVal cnt,Paramfile.ReadLine
        next
        ' Download ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Visual.DwnldFile.InnerHtml         = Paramfile.ReadLine
        ' Memory ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Visual.SelMemTarget.SelectedIndex  = Paramfile.ReadLine      
        Visual.MemDataFormat.SelectedIndex = Paramfile.ReadLine      
        Visual.SelMemType.SelectedIndex    = Paramfile.ReadLine      
        Visual.MemoryStartAdr.Value        = Paramfile.ReadLine
        Visual.MemoryEndAdr.Value          = Paramfile.ReadLine
        Visual.DwnldTarget.Value           = Paramfile.ReadLine
        Visual.RepRateAckErrInp.Value      = Paramfile.ReadLine
        Visual.SelRandAckErr.SelectedIndex = Paramfile.ReadLine
      End If

      ' Close param file ~~~~
      Set Paramfile = Nothing
    End If
  End If
  configLoaded = True
End Sub

Sub WriteConfigParam
  Dim Paramfile, Str, RunTime, cnt
  Set Paramfile = File.Open( System.Environment.Path & "Config\MF_Config.txt", "wt")

  Paramfile.Write $(CONFIG_VERSION)                                    & Chr(10)
  ' COM params ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  for cnt = 1 to 5
    Paramfile.Write Visual.Script("SysParaGrid").getVal(cnt)           & Chr(10)
  next
  ' Download ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ParamFile.Write Visual.DwnldFile.InnerHtml                           & Chr(10)
  ' Memory ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ParamFile.Write Visual.SelMemTarget.SelectedIndex                    & Chr(10)
  ParamFile.Write Visual.MemDataFormat.SelectedIndex                   & Chr(10)
  ParamFile.Write Visual.SelMemType.SelectedIndex                      & Chr(10)
  ParamFile.Write Visual.MemoryStartAdr.Value                          & Chr(10)
  ParamFile.Write Visual.MemoryEndAdr.Value                            & Chr(10)
  ParamFile.Write Visual.DwnldTarget.Value                             & Chr(10)
  ParamFile.Write Visual.RepRateAckErrInp.Value                        & Chr(10)
  ParamFile.Write Visual.SelRandAckErr.SelectedIndex                   & Chr(10)
  
  ' Close param file ---------------------------------------
  Set Paramfile = Nothing
End Sub


'-------------------------------------------------------------------------------
'--- Highlight selected command in grid ----------------------------------------
'-------------------------------------------------------------------------------
Sub OnChange_SelCommand( Reason )
  Visual.Script("CommandTableGrid").SelectRow(Visual.SelCommand.SelectedIndex)
End Sub
 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~ Clear scope content ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub OnClick_BUTTON_CLEAR_SCOPE(Reason)
  Memory.SerialScope.Clear
End Sub  

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~ Open tzool description ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub OnClick_BUTTON_OpenToolSpec(Reason)

  System.OpenDocument System.Environment.Path & "Doc\CacciaTool_MF.pdf", "CAdobeReaderDocument"

End Sub

