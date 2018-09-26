
Const DEFAULT_SLOT = 28
Const COMP_TYPE_RES  = 1
Const COMP_TYPE_CAP  = 2
Const COMP_TYPE_IND  = 3
Const COMP_TYPE_DIODE= 4
Const COMP_TYPE_PCAP = 5
Const COMP_TYPE_AUTO = 6

Const PROCESS_NONE = 0
Const PREPARE_NONE = 0
Const PREPARE_ALL = 6

Const TIO_SETUPMEASURE = 10000
Const TIO_MEASURE = 10000
Const TIO_CALIBRATE = 100000
Const TIO_SELFTEST = 10000

Const ERR_SELF_TEST_INPUT_OFFSET_R93_WARN = 20
Const ERR_SELF_TEST_ADC_IPP_1KHZ_WARN     = 21
Const ERR_SELF_TEST_ADC_IPP_10KHZ_WARN    = 22
Const ACK_CM_NOT_CONNECTED            = &hD1
Const ACK_WRONG_POLARITY              = &hD2
Const ACK_MAX_VOLTAGE                 = &hD3
Const ACK_INVALID_INPUT_OFFSET        = &hD4
Const ACK_INVALID_1KHZ_ADC_I_PP       = &hD5
Const ACK_INVALID_10KHZ_ADC_I_PP      = &hD6
Const ACK_TRIM_SHORT_NOK              = &hD7
Const ACK_TRIM_OPEN_NOK               = &hD8
Const ACK_TRIM_CANNOT_SAVE            = &hD9
Const ACK_CAL_CANNOT_SAVE             = &hDA
Const ACK_CANNOT_GET_BOARD_VERS       = &hDB
Const ACK_CANNOT_CAL_DIODE_DIFF       = &hDC
Const ACK_AUTO_RANGE_DID_NOT_SUCCEED  = &hDD
Const ACK_CAP_AUTO_POL_UNDETERMINED   = &hDE
Const ACK_FWD_VOLT_ONLY_2MA_OR_10MA   = &hDF

Const Debug_Bit_SCITX = 5
Const Debug_Bit_SCIRX = 6

Const Log_SCI_TX   = 0
Const Log_SCI_RX   = 1
Const Log_SCI_TXRX = 2

Const PARAM_MB_STATUS  = &h55
Const PARAM_DATA_STATUS = &h56
Const PARAM_MB_OVERRIDE = &h57

Const CALB_COMP1 = &h8080F1
Const CALB_COMP2 = &h808001
Const CALB_COMP3 = &h808003
Const CALB_COMP4 = &h418001
Const CALB_COMP5 = &h488001
Const CALB_COMP6 = &h428001
Const CALB_COMP7 = &h508001
Const CALB_COMP8 = &h448001
Const CALB_COMP9 = &h608001

Function Init_MFCommand ( )
  Dim PrepCmd_Inprogress,PrepCmd_Error,PrepCmd_PrepID,Endurance_Inprogress
  Dim PrepCmd_MeasureInProgress,counter
  ChangeVisibility_ComponentSelect COMP_TYPE_RES
  ChangeVisibility_Result PROCESS_NONE

  PrepCmd_Inprogress = 0
  Endurance_Inprogress = 0
  PrepCmd_Error = 0
  PrepCmd_PrepID = 1
  Memory.Set "PrepCmd_Inprogress",PrepCmd_Inprogress
  Memory.Set "PrepCmd_Error",PrepCmd_Error
  Memory.Set "Endurance_Inprogress",Endurance_Inprogress
  Memory.Set "PrepCmd_PrepID",PrepCmd_PrepID
  
  Visual.Select("cbmodesel").Checked = True
  Visual.Select("ip_param_setupExpectedVal").Value = 100E-06
  Visual.Select("opt_polarity").Value = 0
  Visual.Select("ip_paramnumofcycles").Value = 0
  Visual.Select("ip_paramvoltage").Value = 3
  Visual.Select("ip_parammaxvoltage").Value = 5
  
  Visual.Select("param_numofcycle").Style.Display = "None"
    
  For counter = 1 To 60
    Visual.Select("opt_SlotNum").addItem counter,counter
  Next
   
  Visual.Select("opt_SlotNum").SelectedIndex  = DEFAULT_SLOT-1
  Memory.Set "SLOT_NO",Visual.Select("opt_SlotNum").SelectedItemAttribute("Value")
  DebugMessage "SLOT_NO :" & Memory.SLOT_NO
    
  Visual.Select("ip_param_setupExpectedVal").SetValidation VALIDATE_INPUT_MASK_R4,"Red",10
  
  PrepCmd_MeasureInProgress = 0
  Memory.Set "PrepCmd_MeasureInProgress",PrepCmd_MeasureInProgress
  
  InitEEPROMGrid
  
End Function

'------------------------------------------------------------------

Sub Get_Firmware( )
  Dim AppMaj,AppMin
  Dim Bios,App1,App2,App3
  If Command_GetFW($(PARAM_DL_ZIEL_APPL),AppMaj,AppMin) = 1 Then
    App1 = String.Format("%02X.%02X", AppMaj,AppMin)
  Else
    App1 = "??.??"
  End If
  
  If Command_GetFW($(PARAM_DL_ZIEL_BIOS),AppMaj,AppMin) = 1 Then
    Bios = String.Format("%02X.%02X", AppMaj,AppMin)
  Else
    Bios = "??.??"
  End If
  
  If Command_GetFW($(PARAM_DL_ZIEL_APPL_2),AppMaj,AppMin) = 1 Then
    App2 = String.Format("%02X.%02X", AppMaj,AppMin)
    GetSCITrace Log_SCI_TXRX,"Firmware App2:"
  Else
    App2 = "??.??"
  End If
  
  If Command_GetFW($(PARAM_DL_ZIEL_APPL_3),AppMaj,AppMin) = 1 Then
    App3 = String.Format("%02X.%02X", AppMaj,AppMin)
    GetSCITrace Log_SCI_TXRX, "Firmware App3:"
  Else
    App3 = "??.??"
  End If
  LogAdd "Firmware version: Bios:"& Bios & " App: " & App1 & " App2: " & App2 & " App3: " & App3
End Sub

'------------------------------------------------------------------

Function Wait_Measurement ( TimeOut )
Dim loop_enable,Time,measureOK
Dim cmd

  measureOK = 0
  Time = TimeOut / 100
  loop_enable = 1
  Do
        'Check Error Stop Condition
      If Memory.PrepCmd_Error = 1 Then
        LogAdd "Measurement command error"
        loop_enable = 0
      'Check button stop condition.
      Elseif Memory.PrepCmd_Inprogress = 0 Then
        If Not Memory.Exists("sig_ERexternalstop") Then
          LogAdd "Measurement Complete"
        End If
        measureOK = 1
        loop_enable = 0
      End If
      
      Time = Time - 1
      If Time = 0 Then
        loop_enable = 0
        LogAdd "Measurement timeout"
        LED_Change 1
      Else
        System.Delay(100)
      End If  
    Loop Until loop_enable = 0 
    
    If measureOK = 1 Then
      Get_Measurements
    End If
    
    Memory.Set "measureOK",measureOK
    Memory.PrepCmd_MeasureInProgress = 0
    Dim LogMsg
    Select Case Memory.PrepCmd
      Case PREPARE_NONE :
      
      Case $(CMD_PREPARE_SETUP_MEASURE) :
        LogMsg = "Setup Measure"
      Case $(CMD_PREPARE_MEASURE) :
        LogMsg = "Prepare Measure"
      Case $(CMD_PREPARE_SELFTEST) :
        LogMsg = "Selftest"
      Case $(CMD_PREPARE_CALIBRATION) :
        LogMsg = "Compensation"
      Case $(CMD_PREPARE_MEASURE_AUTO) :
        LogMsg = "Auto Meas"
    End Select
    GetSCITrace Log_SCI_TXRX, LogMsg    
    If Memory.PrepCmd_Error = 1  AND Memory.CanErr = $(PUB_MB_ERROR) Then
      GetSCIErrorQueue
    End If
    
End Function 
'------------------------------------------------------------------
Function Get_Measurements ( )
  Dim ResultLog,Value, CompType
  DebugMessage "Process Results"
  ChangeVisibility_Result(Memory.PrepCmd)
  'Read params from MF depending on the prepare command executed.
  Select Case Memory.PrepCmd
    Case PREPARE_NONE : 
      ResultLog = "Error, no prepare"
    Case $(CMD_PREPARE_SETUP_MEASURE) :
    
      ResultLog = "Setup Measure: Contact Res :" & GetFloatCanData($(PARAM_SETUP_MEAS_CONTACT_RES),"op_paramContactres")
      ResultLog = ResultLog & " Cap:" & GetFloatCanData($(PARAM_SETUP_MEAS_STRAY_CAPACITY),"op_paramCapacityCM")
      ResultLog = ResultLog & " Res:" & GetFloatCanData($(PARAM_SETUP_MEAS_RESISTANCE),"op_paramResistanceCM")
      ResultLog = ResultLog & " U:" & GetFloatCanData($(PARAM_SETUP_MEAS_U),"op_paramU")
      ResultLog = ResultLog & " I:" & GetFloatCanData($(PARAM_SETUP_MEAS_I),"op_paramI")
      ResultLog = ResultLog & " Phi:" & GetFloatCanData($(PARAM_SETUP_MEAS_PHI),"op_paramPhi")
      ResultLog = ResultLog & " Freq:" & GetFloatCanData($(PARAM_SETUP_MEAS_FREQ),"op_paramFreq")
      
    Case $(CMD_PREPARE_MEASURE) :
      
      CompType = Visual.Select("opt_MeasureCommand").Value
      'Not Diode
      If Not CompType = 4 Then
        DebugMessage "Process Measure CRL, PCAP"
        
        CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_COMPONENT_TYPE),Memory.SLOT_NO,1,0
        Value = Memory.CanData(2)
        Visual.Select("op_paramcomptype1").Value = String.Format("%c",Value)
        
        ResultLog = "Meas: CompType1 :" & String.Format("%c",Value)              
        ResultLog = ResultLog & " Value:" & GetFloatCanData($(PARAM_MEAS_VALUE),"op_paramType1Value")
        ResultLog = ResultLog & " Min:" & GetFloatCanData($(PARAM_MEAS_VALUE_MIN),"op_paramType1ValueMin")
        ResultLog = ResultLog & " Max:" & GetFloatCanData($(PARAM_MEAS_VALUE_MAX),"op_paramType1ValueMax")
        
        CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_COMPONENT_TYPE2),Memory.SLOT_NO,1,0
        Value = Memory.CanData(2)
        Visual.Select("op_paramcomptype2").Value = String.Format("%c",Value)
        ResultLog = ResultLog & " CompType2 :" & String.Format("%c",Value)      

        ResultLog = ResultLog & " Value:" & GetFloatCanData($(PARAM_MEAS_VALUE2),"op_paramType2Value")
        ResultLog = ResultLog & " Min:" & GetFloatCanData($(PARAM_MEAS_VALUE_MIN2),"op_paramType2ValueMin")
        ResultLog = ResultLog & " Max:" & GetFloatCanData($(PARAM_MEAS_VALUE_MAX2),"op_paramType2ValueMax")
        ResultLog = ResultLog & " U:" & GetFloatCanData($(PARAM_MEAS_U),"op_paramU")
        ResultLog = ResultLog & " I:" & GetFloatCanData($(PARAM_MEAS_I),"op_paramI")
        ResultLog = ResultLog & " Phi:" & GetFloatCanData($(PARAM_MEAS_PHI),"op_paramPhi")
        ResultLog = ResultLog & " Freq:" & GetFloatCanData($(PARAM_MEAS_FREQUENCY),"op_paramFreq")

      'Diode
      Else
        DebugMessage "Process Measure Diode"
        
        ResultLog = "Meas: FWD Voltage :" & GetFloatCanData($(PARAM_MEAS_FWD_VOLTAGE),"op_paramfwdvoltage")
        ResultLog = ResultLog & " Min:" & GetFloatCanData($(PARAM_MEAS_VALUE_MIN),"op_paramType1ValueMin")
        ResultLog = ResultLog & " Max:" & GetFloatCanData($(PARAM_MEAS_VALUE_MAX),"op_paramType1ValueMax")
        ResultLog = ResultLog & " FWD Current:" & GetFloatCanData($(PARAM_MEAS_FWD_CURRENT),"op_paramfwdcurrent")
        ResultLog = ResultLog & " Min:" & GetFloatCanData($(PARAM_MEAS_VALUE_MIN2),"op_paramType2ValueMin")
        ResultLog = ResultLog & " Max:" & GetFloatCanData($(PARAM_MEAS_VALUE_MAX2),"op_paramType2ValueMax")
      End If      
      'End CMD_PREPARE_MEASURE
    Case $(CMD_PREPARE_SELFTEST) :
      ResultLog = "Self Test: Contact Res :" & GetFloatCanData($(PARAM_SELFTEST_CONTACT_RES),"op_paramContactres")
      ResultLog = ResultLog & " Cap:" & GetFloatCanData($(PARAM_SELFTEST_CAPACITY_CM_ID),"op_paramCapacityCM")
      ResultLog = ResultLog & " Res:" & GetFloatCanData($(PARAM_SELFTEST_RESISTANCE),"op_paramResistanceCM")
      ResultLog = ResultLog & " U:" & GetFloatCanData($(PARAM_SELFTEST_U),"op_paramU")
      ResultLog = ResultLog & " I:" & GetFloatCanData($(PARAM_SELFTEST_I),"op_paramI")
      ResultLog = ResultLog & " Phi:" & GetFloatCanData($(PARAM_SELFTEST_PHI),"op_paramPhi")
      ResultLog = ResultLog & " Freq:" & GetFloatCanData($(PARAM_SELFTEST_FREQ),"op_paramFreq")      
      
    Case $(CMD_PREPARE_CALIBRATION) :

      ResultLog = "Compensation: R_1kHz :" & GetFloatCanData($(PARAM_CAL_R_1KHz),"op_paramres1k")
      ResultLog = ResultLog & " X_1kHz:" & GetFloatCanData($(PARAM_CAL_X_1KHz),"op_paramreac1k")
      ResultLog = ResultLog & " R_10kHz:" & GetFloatCanData($(PARAM_CAL_R_10KHz),"op_paramres10k")
      ResultLog = ResultLog & " X_10kHz:" & GetFloatCanData($(PARAM_CAL_X_10KHz),"op_paramreac10k")

    Case $(CMD_PREPARE_MEASURE_AUTO) :
    
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_COMPONENT_TYPE),Memory.SLOT_NO,1,0
      Value = Memory.CanData(2)
      Visual.Select("op_paramcomptype1").Value = String.Format("%c",Value)        
      ResultLog = "Meas: CompType1 :" & String.Format("%c",Value)              
      ResultLog = ResultLog & " Value:" & GetFloatCanData($(PARAM_MEAS_VALUE),"op_paramType1Value")
      ResultLog = ResultLog & " Min:" & GetFloatCanData($(PARAM_MEAS_VALUE_MIN),"op_paramType1ValueMin")
      ResultLog = ResultLog & " Max:" & GetFloatCanData($(PARAM_MEAS_VALUE_MAX),"op_paramType1ValueMax")        
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_COMPONENT_TYPE2),Memory.SLOT_NO,1,0
      Value = Memory.CanData(2)
      Visual.Select("op_paramcomptype2").Value = String.Format("%c",Value)
      ResultLog = ResultLog & " CompType2 :" & String.Format("%c",Value)      
      ResultLog = ResultLog & " Value:" & GetFloatCanData($(PARAM_MEAS_VALUE2),"op_paramType2Value")
      ResultLog = ResultLog & " Min:" & GetFloatCanData($(PARAM_MEAS_VALUE_MIN2),"op_paramType2ValueMin")
      ResultLog = ResultLog & " Max:" & GetFloatCanData($(PARAM_MEAS_VALUE_MAX2),"op_paramType2ValueMax")
      ResultLog = ResultLog & " U:" & GetFloatCanData($(PARAM_MEAS_U),"op_paramU")
      ResultLog = ResultLog & " I:" & GetFloatCanData($(PARAM_MEAS_I),"op_paramI")
      ResultLog = ResultLog & " Phi:" & GetFloatCanData($(PARAM_MEAS_PHI),"op_paramPhi")
      ResultLog = ResultLog & " Freq:" & GetFloatCanData($(PARAM_MEAS_FREQUENCY),"op_paramFreq")        

    End Select
    LogAdd ResultLog
    DebugMessage ResultLog
End Function 

'------------------------------------------------------------------

Function OnBlur_inputCANID ( Reason )
  Dim CANID   
  CANID = CLng("&h" & Visual.Select("inputCANID").Value)
  If CANID_Validate (CANID) = True Then  
    DebugMessage "Changed CAN ID:" & String.Format("%3X",CANID)
    CANID_Set CANID
    Memory.CANConfig.CANIDvalid = 1
  Else
    Memory.CANConfig.CANIDvalid = 0
  End If
  
End Function
'------------------------------------------------------------------

Function OnClick_btn_AssignCANID( Reason )
  Dim CanReadArg,CanID
  Set CanReadArg = CreateObject("ICAN.CanReadArg")
  
  CanID = CLng("&h" & Visual.Select("inputCANID").Value)
  'InitCAN CanID
  LogAdd "Assign CANID"
  CANID_Assign CanID
  System.Delay(100)
  If Command_GetNumOfSlots = True Then
    Get_Firmware
    'Debug_Set_Bit Debug_Bit_SCIRX
    'Set Green LED    
    Memory.CANData(0) = 1 
    Memory.CANData(1) = 2 
    CANSendGetMC $(CMD_SEND_DATA),$(MC_SET_LED),Memory.SLOT_NO,1,2
    Memory.CANData(0) = 1 
    CANSendGetMC $(CMD_SEND_DATA),$(MC_STATUS),Memory.SLOT_NO,1,1
  End If
  
End Function

'------------------------------------------------------------------

Function OnClick_btn_setupmeasure( Reason )
Dim ExpectedValue ,CM_ID, CompType, ModeSelect
  ExpectedValue = Math.CastFloat2Long(Visual.Select("ip_param_setupExpectedVal").Value)
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")
  CompType = Visual.Select("opt_MeasureCommand").Value
  ModeSelect = GetModeSelect()
  If NOT IsNumeric(ExpectedValue) Then
    LogAdd "Invalid value"
  Else    
    Command_Prepare_SetupMeasure CM_ID,ExpectedValue,CompType,ModeSelect,TIO_SETUPMEASURE
  End If
End Function

'------------------------------------------------------------------

Sub OnClick_btn_GridClear( Reason )
  Visual.Script("LogGrid").clearAll()  
  Visual.Script("SCIGrid").clearAll()
End Sub

'------------------------------------------------------------------

Sub OnClick_btn_GetApp( Reason )  
  Get_Firmware
End Sub

'------------------------------------------------------------------
Function OnClick_btn_Reset ( Reason )
  Dim CanSendArg,CanReadArg,CANConfig
  Dim encodersel
  Set CanSendArg = CreateObject("ICAN.CanSendArg")
  Set CanReadArg = CreateObject("ICAN.CanReadArg")

  If Memory.Exists( "CanManager" ) Then
    Memory.Get "CANConfig",CANConfig
    CanSendArg.CanId = CANConfig.CANIDcmd
    CanSendArg.Data(0) = &h06
    CanSendArg.Length = 1
    Memory.CanManager.Send CanSendArg
    LogAdd "Resetting Measurement Feeder." 
  End if
End Function

'------------------------------------------------------------------

Function OnClick_btn_GetSCIbuffer( Reason )
  GetSCITrace Log_SCI_TXRX, "Data: "
End Function 
'------------------------------------------------------------------

Function OnClick_btn_measure( Reason )
  Dim CompType
  DebugMessage "PrepareMeasure Command"
  CompType = Visual.Select("opt_MeasureCommand").Value
  Select Case CompType
    Case 1:
      PrepareMeasureCRL 
    Case 2:
      PrepareMeasureCRL    
    Case 3:
      PrepareMeasureCRL
    Case 4:
      PrepareMeasureDiode
    Case 5:
      PrepareMeasurePolarCap
  End Select
End Function

'------------------------------------------------------------------

Function OnClick_btn_calibrate ( Reason )
  Dim command
  Dim Reply,CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")  
  command = Visual.Select("opt_SubCmd").Value
  Memory.CANData(0) = command
  Reply = MsgBox("Are you sure you wish to start compensation?", 1 , "Confirm compensation")
  If Reply = 1 Then
    If CANSendPrepareCMD($(CMD_PREPARE_CALIBRATION),1,Memory.SLOT_NO,CM_ID,1,250) = True Then
      Memory.Set "PrepCmd", $(CMD_PREPARE_CALIBRATION)
      LogAdd "Compensation command started"
      System.Start "Wait_Measurement",TIO_CALIBRATE
    Else
      LogAdd "Compensation Error."
    End If
  End If
End Function

'------------------------------------------------------------------

Function OnClick_btn_selftest ( Reason )
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")
  If CANSendPrepareCMD($(CMD_PREPARE_SELFTEST),1,Memory.SLOT_NO,CM_ID,0,250) = True Then
    Memory.Set "PrepCmd", $(CMD_PREPARE_SELFTEST)
    LogAdd "Self Test command started"
    System.Start "Wait_Measurement",TIO_SELFTEST
  Else
    LogAdd "Self Test Error."
  End If
End Function

'------------------------------------------------------------------

Function OnClick_btn_endurance ( Reason )

  System.Start Endurance(5000)

End Function

'------------------------------------------------------------------
Function OnClick_btn_endu_stop ( Reason )

If Memory.Exists("sig_ERexternalstop") Then
  LogAdd "Endurance Run Stopping..."
  Memory.sig_ERexternalstop.Set
Else
  LogAdd "No Endurance run to stop."
End If
End Function

'------------------------------------------------------------------

Function OnClick_btn_getcover ( Reason )
  Dim CM_ID, CanReadArg
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")  
  If CANSendGetMC ($(CMD_GET_DATA),$(PARAM_INP_COVER),Memory.SLOT_NO,CM_ID,0) = True Then
    If Memory.CanData.Data(2) = 1 Then
      LED_Update "ledcover",1
      GetSCITrace Log_SCI_TXRX, "Cover Open"
    Else
      LED_Update "ledcover",0
      GetSCITrace Log_SCI_TXRX, "Cover Closed"

    End If
  Else
    LED_Update "ledcover",0
  End If
  'GetSCILog "Get Cover Slot: "
End Function

'------------------------------------------------------------------

Function OnClick_btn_gettemp ( Reason )
  Dim CM_ID, CanReadArg, Temp
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")  
  If CANSendGetMC ($(CMD_GET_DATA),$(PARAM_MB_TEMP),Memory.SLOT_NO,CM_ID,0) = True Then
    Temp = Lang.MakeInt(Memory.CanData(2),Memory.CanData(3))
    Visual.Select("op_paramtemp").Value = Temp
    LogAdd "Temperature: " & Temp

  Else    
    Temp = "??"
  End If
  Visual.Select("op_paramtemp").Value = Temp
  'GetSCILog "Get Cover Temp: "
End Function

'------------------------------------------------------------------
Function OnClick_btn_getresults ( Reason )
  If Memory.PrepCmd_Inprogress = 0 Then
    Memory.Set "PrepCmd", $(CMD_PREPARE_MEASURE)
    Get_Measurements
  Else
    LogAdd "Measurement in progress"
  End If
End Function
'------------------------------------------------------------------
Function OnClick_btn_GetSCIErrQ ( Reason )
  GetSCIErrorQueue
End Function

'------------------------------------------------------------------
Function OnClick_btn_DebugLog ( Reason )
  If Memory.Exists("DebugLogWindow") Then
    DebugWindowClose
  Else 
    CreateDebugLogWindow
  End If

End Function


'------------------------------------------------------------------
Function OnClick_btn_getstatus ( Reason )
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")    
  Dim StatusWord,DataStatusWord,OverrideWord
  If CANSendGetMC ($(CMD_GET_DATA),PARAM_MB_STATUS,Memory.SLOT_NO,CM_ID,0) = True Then
    StatusWord = Lang.MakeWord(Memory.CanData(2),Memory.CanData(3))
    UpdateStatus "ipstatus_",StatusWord,13
  End If
  If CANSendGetMC ($(CMD_GET_DATA),PARAM_DATA_STATUS,Memory.SLOT_NO,CM_ID,0) = True Then
    DataStatusWord = Lang.MakeWord(Memory.CanData(2),Memory.CanData(3))
  End If
  If CANSendGetMC ($(CMD_GET_DATA),PARAM_MB_OVERRIDE,Memory.SLOT_NO,CM_ID,0) = True Then
    OverrideWord = Lang.MakeWord(Memory.CanData(2),Memory.CanData(3))
    UpdateStatus "ipdebug_",OverrideWord,10
  End If
  DebugMessage String.Format("Status %02X",StatusWord) & String.Format(" DataStat %02X",DataStatusWord) & String.Format(" Override %02X",OverrideWord) 
End Function
'------------------------------------------------------------------

Function UpdateStatus ( handle,data,length )
  Dim iBit,handlename
  For iBit = 0 to length-1
    handlename = handle & String.Format("%01d",iBit+1)
    'DebugMessage handlename & " " & Lang.Bit(data,iBit)
    LED_Update handlename, Lang.Bit(data,iBit)
  Next
End Function
'------------------------------------------------------------------
Function OnChange_opt_SlotNum ( Reason )
  Memory.Set "SLOT_NO",Visual.Select("opt_SlotNum").SelectedItemAttribute("Value")
  DebugMessage "SLOT_NO" & Memory.SLOT_NO
End Function
'------------------------------------------------------------------
Function OnClick_cbmodesel ( Reason )
Dim StatusWord,checkvalue

  If Visual.Select("cbmodesel").Checked = True Then
    Visual.Select("optFrequency").Disabled = False
    Visual.Select("optModel").Disabled = False
  Else
    Visual.Select("optFrequency").Disabled = True
    Visual.Select("optFrequency").Value = 0
    Visual.Select("optModel").Disabled = True
    Visual.Select("optModel").Value = 0
  End If
  'Set bit 1 (mode seleect) to value in element.
  Override_SetBit 1,Visual.Select("cbmodesel").Checked
End Function


'------------------------------------------------------------------
Function OnClick_ipdebug_1 ( Reason )
  Dim Bit,NewValue
  Bit = Visual.Select("ipdebug_1").Value
  NewValue = Override_ToggleBit(Bit)
  LED_Update "ipdebug_1",NewValue
End Function
'------------------------------------------------------------------
Function OnClick_ipdebug_2 ( Reason )
  Dim Bit,NewValue
  Bit = Visual.Select("ipdebug_2").Value
  NewValue = Override_ToggleBit(Bit)
  LED_Update "ipdebug_2",NewValue
End Function
'------------------------------------------------------------------
Function OnClick_ipdebug_3 ( Reason )
  Dim Bit,NewValue
  Bit = Visual.Select("ipdebug_3").Value
  NewValue = Override_ToggleBit(Bit)
  LED_Update "ipdebug_3",NewValue
End Function
'------------------------------------------------------------------
Function OnClick_ipdebug_4 ( Reason )
  Dim Bit,NewValue
  Bit = Visual.Select("ipdebug_4").Value
  NewValue = Override_ToggleBit(Bit)
  LED_Update "ipdebug_4",NewValue
End Function
'------------------------------------------------------------------
Function OnClick_ipdebug_5 ( Reason )
  Dim Bit,NewValue
  Bit = Visual.Select("ipdebug_5").Value
  NewValue = Override_ToggleBit(Bit)
  LED_Update "ipdebug_5",NewValue
End Function

'------------------------------------------------------------------
Function OnClick_ipdebug_6 ( Reason )
  Dim Bit,NewValue
  Bit = Visual.Select("ipdebug_6").Value
  NewValue = Override_ToggleBit(Bit)
  LED_Update "ipdebug_6",NewValue
End Function
'------------------------------------------------------------------
Function OnClick_ipdebug_9 ( Reason )
  Dim Bit,NewValue
  Bit = Visual.Select("ipdebug_9").Value
  NewValue = Override_ToggleBit(Bit)
  LED_Update "ipdebug_9",NewValue
End Function

'------------------------------------------------------------------
Function OnClick_ipdebug_10 ( Reason )
  Dim Bit,NewValue
  Bit = Visual.Select("ipdebug_10").Value
  NewValue = Override_ToggleBit(Bit)
  LED_Update "ipdebug_10",NewValue
End Function
'------------------------------------------------------------------
Function OnClick_btn_Sel1 ( Reason )
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")
  Memory.CANData(0) = 1
  Memory.CANData(1) = Lang.GetByte(CALB_COMP1,0)
  If CANSendGetMC($(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2) = True Then
  Memory.CANData(0) = 2
  Memory.CANData(1) = Lang.GetByte(CALB_COMP1,1)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2
  Memory.CANData(0) = 3
  Memory.CANData(1) = Lang.GetByte(CALB_COMP1,2)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2    
  Else
    LogAdd "Check if Calibration Module is connected"
  End If
End Function
'------------------------------------------------------------------
Function OnClick_btn_Sel2 ( Reason )
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")
  Memory.CANData(0) = 1
  Memory.CANData(1) = Lang.GetByte(CALB_COMP2,0)
 If CANSendGetMC($(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2) = True Then
  Memory.CANData(0) = 2
  Memory.CANData(1) = Lang.GetByte(CALB_COMP2,1)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2
  Memory.CANData(0) = 3
  Memory.CANData(1) = Lang.GetByte(CALB_COMP2,2)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2    
  Else
    LogAdd "Check if Calibration Module is connected"
  End If
End Function
'------------------------------------------------------------------
Function OnClick_btn_Sel3 ( Reason )
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")
  Memory.CANData(0) = 1
  Memory.CANData(1) = Lang.GetByte(CALB_COMP3,0)
  If CANSendGetMC($(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2) = True Then
  Memory.CANData(0) = 2
  Memory.CANData(1) = Lang.GetByte(CALB_COMP3,1)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2
  Memory.CANData(0) = 3
  Memory.CANData(1) = Lang.GetByte(CALB_COMP3,2)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2    
    Else
    LogAdd "Check if Calibration Module is connected"
  End If
End Function
'------------------------------------------------------------------
Function OnClick_btn_Sel4 ( Reason )
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")
  Memory.CANData(0) = 1
  Memory.CANData(1) = Lang.GetByte(CALB_COMP4,0)
  If CANSendGetMC($(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2) = True Then
  Memory.CANData(0) = 2
  Memory.CANData(1) = Lang.GetByte(CALB_COMP4,1)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2
  Memory.CANData(0) = 3
  Memory.CANData(1) = Lang.GetByte(CALB_COMP4,2)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2    
  Else
    LogAdd "Check if Calibration Module is connected"
  End If  
End Function
'------------------------------------------------------------------
Function OnClick_btn_Sel5 ( Reason )
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")
  Memory.CANData(0) = 1
  Memory.CANData(1) = Lang.GetByte(CALB_COMP5,0)
  If CANSendGetMC($(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2) = True Then
  Memory.CANData(0) = 2
  Memory.CANData(1) = Lang.GetByte(CALB_COMP5,1)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2
  Memory.CANData(0) = 3
  Memory.CANData(1) = Lang.GetByte(CALB_COMP5,2)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2 
  Else
    LogAdd "Check if Calibration Module is connected"
  End If
End Function
'------------------------------------------------------------------
Function OnClick_btn_Sel6 ( Reason )
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")
  Memory.CANData(0) = 1
  Memory.CANData(1) = Lang.GetByte(CALB_COMP6,0)
  If CANSendGetMC($(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2) = True Then
  Memory.CANData(0) = 2
  Memory.CANData(1) = Lang.GetByte(CALB_COMP6,1)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2
  Memory.CANData(0) = 3
  Memory.CANData(1) = Lang.GetByte(CALB_COMP6,2)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2    
  Else
    LogAdd "Check if Calibration Module is connected"
  End If  
End Function
'------------------------------------------------------------------
Function OnClick_btn_Sel7 ( Reason )
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")
  Memory.CANData(0) = 1
  Memory.CANData(1) = Lang.GetByte(CALB_COMP7,0)
  If CANSendGetMC($(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2) = True Then
  Memory.CANData(0) = 2
  Memory.CANData(1) = Lang.GetByte(CALB_COMP7,1)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2
  Memory.CANData(0) = 3
  Memory.CANData(1) = Lang.GetByte(CALB_COMP7,2)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2    
  Else
    LogAdd "Check if Calibration Module is connected"
  End If  
End Function
'------------------------------------------------------------------
Function OnClick_btn_Sel8 ( Reason )
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")
  Memory.CANData(0) = 1
  Memory.CANData(1) = Lang.GetByte(CALB_COMP8,0)
  If CANSendGetMC($(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2) = True Then
  Memory.CANData(0) = 2
  Memory.CANData(1) = Lang.GetByte(CALB_COMP8,1)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2
  Memory.CANData(0) = 3
  Memory.CANData(1) = Lang.GetByte(CALB_COMP8,2)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2    
  Else
    LogAdd "Check if Calibration Module is connected"
  End If  
End Function
'------------------------------------------------------------------
Function OnClick_btn_Sel9 ( Reason )
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")
  Memory.CANData(0) = 1
  Memory.CANData(1) = Lang.GetByte(CALB_COMP9,0)
  If CANSendGetMC($(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2) = True Then
  Memory.CANData(0) = 2
  Memory.CANData(1) = Lang.GetByte(CALB_COMP9,1)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2
  Memory.CANData(0) = 3
  Memory.CANData(1) = Lang.GetByte(CALB_COMP9,2)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CALB_IOEXP),Memory.SLOT_NO,CM_ID,2    
  Else
    LogAdd "Check if Calibration Module is connected"
  End If  
End Function
'------------------------------------------------------------------

' Read the override word
Function Override_Get (ByRef Value)
  Dim CM_ID
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")   
  If CANSendGetMC ($(CMD_GET_DATA),PARAM_MB_OVERRIDE,Memory.SLOT_NO,CM_ID,0) = True Then
    Value = Lang.MakeWord(Memory.CanData(2),Memory.CanData(3))
  End If
End Function 
'------------------------------------------------------------------
' Set selected bit in override word, to a certain boolean value
Function Override_SetBit (Bit,Value)
  Dim CM_ID
  Dim OverrideWord
  
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")     
  Override_Get OverrideWord
  OverrideWord = Lang.SetBit(OverrideWord,Bit,Value)
  DebugMessage "Override = " & String.Format("%2X",OverrideWord)
  Memory.CANData(0) = Lang.GetByte(OverrideWord,0)
  Memory.CANData(1) = Lang.GetByte(OverrideWord,1)    
  CANSendGetMC $(CMD_SEND_DATA),PARAM_MB_OVERRIDE,Memory.SLOT_NO,CM_ID,2
End Function
'------------------------------------------------------------------
' Toggle selected bit in override word
Function Override_ToggleBit (Bit)
Dim value
Override_Get value
If Lang.Bit(value,Bit) = 1 Then
  'DebugMessage "value: " & value & "new bit: 0"  
  Override_SetBit Bit,0
  Override_ToggleBit = 0
Else 
  'DebugMessage "value: " & value & "new bit: 1"
  Override_SetBit Bit,1
  Override_ToggleBit = 1
End If
End Function 
'------------------------------------------------------------------
Function OnChange_opt_MeasureCommand ( Reason )
  Dim CompType
  DebugMessage "Select:" & Visual.select("opt_MeasureCommand").Value

  Select Case  Visual.Select("opt_MeasureCommand").Value
  Case "1" : CompType = COMP_TYPE_RES
  Case "2" : CompType = COMP_TYPE_CAP
  Case "3" : CompType = COMP_TYPE_IND
  Case "4" : CompType = COMP_TYPE_DIODE
  Case "5" : CompType = COMP_TYPE_PCAP
  Case "6" : CompType = COMP_TYPE_AUTO
  End Select
  Memory.Set "CompType", CompType
  ChangeVisibility_ComponentSelect CompType
  
End Function
'------------------------------------------------------------------
Function PrepareMeasureCRL ( )

  Dim ExpectedValue, CompType,NumofCycles,CM_ID,ModeSelect
  ExpectedValue = Math.CastFloat2Long(Visual.Select("ip_param_setupExpectedVal").Value)
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")  
  CompType = Visual.Select("opt_MeasureCommand").Value
  NumofCycles = Visual.Select("ip_paramnumofcycles").Value
  ModeSelect = GetModeSelect()

  If NOT IsNumeric(ExpectedValue) Then
    LogAdd "Invalid value"
  Else
    Command_Prepare_Measure CM_ID,ExpectedValue,CompType,NumofCycles,ModeSelect,TIO_MEASURE
  End If

End Function
'------------------------------------------------------------------

Function PrepareMeasureDiode ( )

  Dim CompType,CM_ID, Current,Voltage,Polarity
  Voltage = Math.CastFloat2Long(Visual.Select("ip_paramvoltage").Value)
  Current = Math.CastFloat2Long(Visual.Select("optCurrSel").SelectedItemAttribute("Value"))
  Polarity = Visual.Select("opt_polarity").SelectedItemAttribute("Value")
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")  
  CompType = Visual.Select("opt_MeasureCommand").Value
  If NOT IsNumeric(Voltage) Then
    LogAdd "Invalid Voltage value"
  Else
    Command_Prepare_Meas_FWDVOLTAGE CM_ID,Current,Voltage,CompType,Polarity,TIO_MEASURE
  End If

End Function
'------------------------------------------------------------------
Function PrepareMeasurePolarCap ()

  Dim CompType,CM_ID, Capacitance,Voltage,Polarity
  Voltage = Math.CastFloat2Long(Visual.Select("ip_parammaxvoltage").Value)
  Capacitance = Math.CastFloat2Long(Visual.Select("ip_param_setupExpectedVal").Value)
  Polarity = Visual.Select("opt_polarity").SelectedItemAttribute("Value")
  CM_ID = Visual.Select("opt_CMID").SelectedItemAttribute("Value")  
  CompType = Visual.Select("opt_MeasureCommand").Value
  If NOT IsNumeric(Voltage) Then
    LogAdd "Invalid Voltage value"    
  Elseif Not IsNumeric(Capacitance) Then
    LogAdd "Invalid Capacitance value"    
  Else
    Command_Prepare_Meas_PolarCap CM_ID,Voltage,Capacitance,CompType,Polarity,TIO_MEASURE
  End If

End Function

'------------------------------------------------------------------
Function Command_Prepare_SetupMeasure (CM_ID, ExpectedValue,ComponentType,ModeSelect,TimeOut)
  Memory.CANData(0) = Lang.GetByte(ExpectedValue,0)
  Memory.CANData(1) = Lang.GetByte(ExpectedValue,1)
  Memory.CANData(2) = Lang.GetByte(ExpectedValue,2)
  Memory.CANData(3) = Lang.GetByte(ExpectedValue,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_EXPECTED_RESULT),Memory.SLOT_NO,CM_ID,4
  
  Memory.CANData(0) = Lang.GetByte(ComponentType,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_COMPONENT_TYPE),Memory.SLOT_NO,CM_ID,1  

  Memory.CANData(0) = Lang.GetByte(ModeSelect,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_MODE_SELECT),Memory.SLOT_NO,CM_ID,1  
  
  If CANSendPrepareCMD($(CMD_PREPARE_SETUP_MEASURE),1,Memory.SLOT_NO,CM_ID,0,250) = True Then
    Memory.Set "PrepCmd", $(CMD_PREPARE_SETUP_MEASURE)
    LogAdd "Setup Measure command started"
    System.Start "Wait_Measurement",TimeOut
  Else
    LogAdd "Setup Measure Error."
  End If
End Function

'------------------------------------------------------------------
Function Command_Prepare_Measure (CM_ID,ExpectedValue,ComponentType,NumofCycles,ModeSelect,TimeOut)
  Memory.CANData(0) = Lang.GetByte(ExpectedValue,0)
  Memory.CANData(1) = Lang.GetByte(ExpectedValue,1)
  Memory.CANData(2) = Lang.GetByte(ExpectedValue,2)
  Memory.CANData(3) = Lang.GetByte(ExpectedValue,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_EXPECTED_RESULT),Memory.SLOT_NO,CM_ID,4
  
  Memory.CANData(0) = Lang.GetByte(ComponentType,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_COMPONENT_TYPE),Memory.SLOT_NO,CM_ID,1

  Memory.CANData(0) = Lang.GetByte(NumofCycles,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_NUM_OF_CYCLES),Memory.SLOT_NO,CM_ID,1
  
  Memory.CANData(0) = Lang.GetByte(ModeSelect,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_MODE_SELECT),Memory.SLOT_NO,CM_ID,1  
  
  If CANSendPrepareCMD($(CMD_PREPARE_MEASURE),1,Memory.SLOT_NO,CM_ID,0,250) = True Then
    If Not Memory.Exists("sig_ERexternalstop") Then        
      LogAdd "Measure command started"
    End If
    Memory.Set "PrepCmd", $(CMD_PREPARE_MEASURE)
    Memory.PrepCmd_MeasureInProgress = 1
    System.Start "Wait_Measurement",TimeOut
  Else
    LogAdd "Measure Command Error."
    LED_Change 1
  End If
End Function

'------------------------------------------------------------------
Function Command_Prepare_Meas_FWDVOLTAGE (CM_ID,Current,Voltage,ComponentType,Polarity,TimeOut)
  Memory.CANData(0) = Lang.GetByte(ComponentType,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_COMPONENT_TYPE),Memory.SLOT_NO,CM_ID,1  
  
  Memory.CANData(0) = Lang.GetByte(Polarity,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_POLARITY),Memory.SLOT_NO,CM_ID,1  
  
  Memory.CANData(0) = Lang.GetByte(Current,0)
  Memory.CANData(1) = Lang.GetByte(Current,1)
  Memory.CANData(2) = Lang.GetByte(Current,2)
  Memory.CANData(3) = Lang.GetByte(Current,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CURRENT),Memory.SLOT_NO,CM_ID,4
  
  Memory.CANData(0) = Lang.GetByte(Voltage,0)
  Memory.CANData(1) = Lang.GetByte(Voltage,1)
  Memory.CANData(2) = Lang.GetByte(Voltage,2)
  Memory.CANData(3) = Lang.GetByte(Voltage,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_VOLTAGE),Memory.SLOT_NO,CM_ID,4  

  If CANSendPrepareCMD($(CMD_PREPARE_MEASURE),1,Memory.SLOT_NO,CM_ID,0,250) = True Then
    Memory.Set "PrepCmd", $(CMD_PREPARE_MEASURE)
    LogAdd "Measure command started"
    System.Start "Wait_Measurement",TimeOut
  Else
    LogAdd "Measure Command Error."
  End If
End Function
'------------------------------------------------------------------

Function Command_Prepare_Meas_PolarCap (CM_ID,MaxVoltage,Capacity,ComponentType,Polarity,TimeOut)
  Memory.CANData(0) = Lang.GetByte(MaxVoltage,0)
  Memory.CANData(1) = Lang.GetByte(MaxVoltage,1)
  Memory.CANData(2) = Lang.GetByte(MaxVoltage,2)
  Memory.CANData(3) = Lang.GetByte(MaxVoltage,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_MAX_VOLTAGE),Memory.SLOT_NO,CM_ID,4  
  
  Memory.CANData(0) = Lang.GetByte(Capacity,0)
  Memory.CANData(1) = Lang.GetByte(Capacity,1)
  Memory.CANData(2) = Lang.GetByte(Capacity,2)
  Memory.CANData(3) = Lang.GetByte(Capacity,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_EXPECTED_RESULT),Memory.SLOT_NO,CM_ID,4
  
  Memory.CANData(0) = Lang.GetByte(ComponentType,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_COMPONENT_TYPE),Memory.SLOT_NO,CM_ID,1  
  
  Memory.CANData(0) = Lang.GetByte(Polarity,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_POLARITY),Memory.SLOT_NO,CM_ID,1  

  If CANSendPrepareCMD($(CMD_PREPARE_MEASURE),1,Memory.SLOT_NO,CM_ID,0,250) = True Then
    Memory.Set "PrepCmd", $(CMD_PREPARE_MEASURE)
    LogAdd "Measure Polar cap command started"
    System.Start "Wait_Measurement",TimeOut
  Else
    LogAdd "Measure Polar cap Command Error."
  End If
End Function
'------------------------------------------------------------------

Function Command_GetNumOfSlots( )
  If CANSendGetMC($(CMD_GET_DATA),$(MC_NUMBER_OF_SLOTS),Memory.SLOT_NO,1,0) = False Then
    LogAdd "Get Number of Slot command Error!"
    Command_GetNumOfSlots = False
  Else
    Command_GetNumOfSlots = True
  End If
End Function

'------------------------------------------------------------------

Function Command_GetFW(ByVal AppBios, ByRef MajorValue, ByRef MinValue)
  Dim CanSendArg,CanReadArg,CANConfig
  Dim CanManager, Result, CANData

  Set CanSendArg = CreateObject("ICAN.CanSendArg")
  Set CanReadArg = CreateObject("ICAN.CanReadArg")
  
  If Memory.Exists("CanConfig") Then
    Memory.Get "CanConfig",CanConfig
    Memory.Get "CANData",CANData
  End If
  
  With CanSendArg
    .CanId = CanConfig.CANIDcmd
    If CANConfig.Config = 0 Then
    'Standalone
    .Data(0) = $(CMD_DOWNLOAD_VERSION)
    .Data(1) = AppBios
    .length = 2
    Else
    'XFCU
    .Data(0) = $(CMD_DOWNLOAD_VERSION) + &h10
    .Data(1) = AppBios
    .Data(2) = Memory.SLOT_NO
    .length = 3
    End If
  End With
  If Memory.Exists("CanManager") AND CanConfig.CANIDvalid = 1 Then    
    Memory.Get "CanManager",CanManager        
    Result = CanManager.SendCmd(CanSendArg,250,SC_CHECK_ERROR_BYTE,CanReadArg)
    DebugMessage "GetVer: " & "(TX: " & CanSendArg.Format & ") (RX:" & CanReadArg.Format & ")"
    'DebugMessage "Result" & Result
    If Result = SCA_NO_ERROR Then
      Command_GetFW = 1
      If CANConfig.Config = 0 Then
        MajorValue = CanReadArg.Data(3)
        MinValue   = CanReadArg.Data(4)
      Else
        MajorValue = CanReadArg.Data(4)
        MinValue   = CanReadArg.Data(5)      
      End If
    Else
      Command_GetFW = 0
    End If
  Else
    LogAdd ("No CAN Manager")
  End If
End Function


'-------------------------------------------------------------------

Function ChangeVisibility_Result ( ProcessType )  
  Dim CompType
  Visual.Select("LayerResults").Style.Display = "None"
  'Set all to none
  Visual.Select("param_ContactRes").Style.Display = "None"
  Visual.Select("param_CapacityCM").Style.Display = "None"
  Visual.Select("param_ResistanceCM").Style.Display = "None"
  Visual.Select("param_CompType1").Style.Display = "None"
  Visual.Select("param_CompType2").Style.Display = "None"
  Visual.Select("param_Type1Value").Style.Display = "None"
  Visual.Select("param_Type1ValueMin").Style.Display = "None"
  Visual.Select("param_Type1ValueMax").Style.Display = "None"
  Visual.Select("param_Type2Value").Style.Display = "None"  
  Visual.Select("param_Type2ValueMin").Style.Display = "None"  
  Visual.Select("param_Type2ValueMax").Style.Display = "None"  
  Visual.Select("param_fwdvoltage").Style.Display = "None"  
  Visual.Select("param_fwdcurrent").Style.Display = "None"  
  Visual.Select("param_U").Style.Display = "None"  
  Visual.Select("param_I").Style.Display = "None"  
  Visual.Select("param_Phi").Style.Display = "None"  
  Visual.Select("param_Freq").Style.Display = "None"  
  Visual.Select("param_res1k").Style.Display = "None"  
  Visual.Select("param_reac1k").Style.Display = "None"  
  Visual.Select("param_res10k").Style.Display = "None"  
  Visual.Select("param_reac10k").Style.Display = "None"  
  Select Case ProcessType 
  Case PREPARE_NONE:
  'None : Should not occur
  Case $(CMD_PREPARE_SETUP_MEASURE):
  Visual.Select("param_ContactRes").Style.Display = "Block"
  Visual.Select("param_CapacityCM").Style.Display = "Block"
  Visual.Select("param_ResistanceCM").Style.Display = "Block"
  Visual.Select("param_U").Style.Display = "Block"
  Visual.Select("param_I").Style.Display = "Block"
  Visual.Select("param_Phi").Style.Display = "Block"
  Visual.Select("param_Freq").Style.Display = "Block"
  Case $(CMD_PREPARE_MEASURE):
  CompType = Visual.Select("opt_MeasureCommand").Value
  Visual.Select("param_U").Style.Display = "Block"
  Visual.Select("param_I").Style.Display = "Block"
  Visual.Select("param_Phi").Style.Display = "Block"
  Visual.Select("param_Freq").Style.Display = "Block"
  'Component = R, C or L
  If Not CompType = 4 Then
    Visual.Select("param_CompType1").Style.Display = "Block"
    Visual.Select("param_CompType2").Style.Display = "Block"
    Visual.Select("param_Type1Value").Style.Display = "Block"
    Visual.Select("param_Type1ValueMin").Style.Display = "Block"
    Visual.Select("param_Type1ValueMax").Style.Display = "Block"
    Visual.Select("param_Type2Value").Style.Display = "Block"  
    Visual.Select("param_Type2ValueMin").Style.Display = "Block"  
    Visual.Select("param_Type2ValueMax").Style.Display = "Block"
    Visual.Select("param_U").Style.Display = "Block"
    Visual.Select("param_I").Style.Display = "Block"
  'Component = Diode
  Else
    Visual.Select("param_fwdvoltage").Style.Display = "Block"
    Visual.Select("param_fwdcurrent").Style.Display = "Block"
    Visual.Select("param_Type1ValueMin").Style.Display = "Block"  
    Visual.Select("param_Type1ValueMax").Style.Display = "Block"
    Visual.Select("param_Type2ValueMin").Style.Display = "Block"  
    Visual.Select("param_Type2ValueMax").Style.Display = "Block" 
  End If
  Case $(CMD_PREPARE_SELFTEST) :  
  Visual.Select("param_ContactRes").Style.Display = "Block"
  Visual.Select("param_CapacityCM").Style.Display = "Block"
  Visual.Select("param_ResistanceCM").Style.Display = "Block"
  Visual.Select("param_U").Style.Display = "Block"
  Visual.Select("param_I").Style.Display = "Block"
  Visual.Select("param_Phi").Style.Display = "Block"
  Visual.Select("param_Freq").Style.Display = "Block"
  'For debug
  Case $(CMD_PREPARE_CALIBRATION):
  Visual.Select("param_res1k").Style.Display = "Block"  
  Visual.Select("param_reac1k").Style.Display = "Block"  
  Visual.Select("param_res10k").Style.Display = "Block"  
  Visual.Select("param_reac10k").Style.Display = "Block"
  
  Case PREPARE_ALL:
  Visual.Select("param_ContactRes").Style.Display = "Block"
  Visual.Select("param_CapacityCM").Style.Display = "Block"
  Visual.Select("param_ResistanceCM").Style.Display = "Block"
  Visual.Select("param_CompType1").Style.Display = "Block"
  Visual.Select("param_CompType2").Style.Display = "Block"
  Visual.Select("param_Type1Value").Style.Display = "Block"
  Visual.Select("param_Type1ValueMin").Style.Display = "Block"
  Visual.Select("param_Type1ValueMax").Style.Display = "Block"
  Visual.Select("param_Type2Value").Style.Display = "Block"  
  Visual.Select("param_Type2ValueMin").Style.Display = "Block"  
  Visual.Select("param_Type2ValueMax").Style.Display = "Block"  
  Visual.Select("param_fwdvoltage").Style.Display = "Block"  
  Visual.Select("param_fwdcurrent").Style.Display = "Block"  
  Visual.Select("param_U").Style.Display = "Block"  
  Visual.Select("param_I").Style.Display = "Block"  
  Visual.Select("param_Phi").Style.Display = "Block"  
  Visual.Select("param_Freq").Style.Display = "Block"  
  Visual.Select("param_res1k").Style.Display = "Block"  
  Visual.Select("param_reac1k").Style.Display = "Block"  
  Visual.Select("param_res10k").Style.Display = "Block"  
  Visual.Select("param_reac10k").Style.Display = "Block"
  
  End Select
  Visual.Select("LayerResults").Style.Display = "Block"
End Function
'------------------------------------------------------------------
Function ChangeVisibility_ComponentSelect ( CompType )
  'DebugMessage "Change:" & CompType
  'Set all fields to none
    Visual.Select("param_voltage").Style.Display = "None"
    Visual.Select("param_expectedVal").Style.Display = "None"
    Visual.Select("param_current").Style.Display  = "None"
    Visual.Select("param_maxvoltage").Style.Display  = "None"
    Visual.Select("param_numofcycle").Style.Display  = "None"
    Visual.Select("param_polarity").Style.Display  = "None"  
  Select Case CompType
  'Res
  Case 1:
  
    Visual.Select("ParamUnit").InnerHTML  = "Ohm"
    Visual.Select("param_expectedVal").Style.Display = "block"
    Visual.Select("param_numofcycle").Style.Display  = "block"
  'Cap
  Case 2:
    Visual.Select("ParamUnit").InnerHTML  = "F"
    Visual.Select("param_expectedVal").Style.Display = "block"
    Visual.Select("param_numofcycle").Style.Display  = "block"    
  'Inductor
  Case 3:
    Visual.Select("ParamUnit").InnerHTML  = "H"
    Visual.Select("param_expectedVal").Style.Display = "block"
    Visual.Select("param_numofcycle").Style.Display  = "block"    
  'Diode
  Case 4:
    Visual.Select("param_voltage").Style.Display = "block"
    Visual.Select("param_current").Style.Display  = "block"
    Visual.Select("param_polarity").Style.Display  = "block"
  'PCAP
  Case 5:
    Visual.Select("ParamUnit").InnerHTML  = "F"
    Visual.Select("param_expectedVal").Style.Display = "block"
    Visual.Select("param_maxvoltage").Style.Display  = "block"
    Visual.Select("ip_paramvoltage").Value  = "5"
    Visual.Select("param_polarity").Style.Display  = "block"
  Case 6:
    'case auto: all none
  End Select
End Function 

'-------------------------------------------------------------------

Function Debug_Set_Bit ( BitNo )

  Dim DebugStatus,DebugStatusNew
  If CANSendGetFeed( $(FEED_GET_DATA),$(PARAM_MEAS_DEBUG),Memory.SLOT_NO,1,0) = True Then
    DebugStatus = Lang.MakeInt(Memory.CANData.Data(2),Memory.CANData.Data(3))
    DebugMessage "DebugStatus: " &String.Format("0x%04X",DebugStatus)
  End If
  
  DebugStatusNew= Lang.SetBit(DebugStatus,BitNo,True)
  DebugMessage "DebugStatusNew: " &String.Format("0x%04X",DebugStatusNew)    
  
  Memory.CANData(0) = Lang.GetByte(DebugStatusNew,0)
  Memory.CANData(1) = Lang.GetByte(DebugStatusNew,1)
  
  If CANSendGetFeed ($(FEED_SEND_DATA),$(PARAM_MEAS_DEBUG),Memory.SLOT_NO,1,2) = True Then
    DebugStatus = Lang.MakeInt(Memory.CANData.Data(0),Memory.CANData.Data(1))
  End If  
End Function 

'------------------------------------------------------------------

Function GetModeSelect ( )
  Dim ModeSelect, Model, Frequency
  Model = Visual.Select("optModel").SelectedItemAttribute("Value")
  Frequency = Visual.Select("optFrequency").SelectedItemAttribute("Value")
  ModeSelect = 0
  ModeSelect = Lang.ShiftLeft(Frequency,2) OR Model
  'DebugMessage Model & " " & Frequency & " " & ModeSelect
  GetModeSelect = ModeSelect 
End Function

'------------------------------------------------------------------


'------------------------------------------------------------------
' Endurance Run Function ------------------------------------------
'------------------------------------------------------------------

Function Endurance ( TimeOut )
  Dim looping
  Dim count
  Dim sig_ERexternalstop
  Dim TIO
  
  TIO = Timeout / 100
  Set sig_ERexternalstop = Signal.Create
  Memory.Set "sig_ERexternalstop", sig_ERexternalstop
  looping = 1
  Do While looping = 1
    'LED On Cycle
    'LED_Change 2
    PrepareMeasureCRL    
    Do    
      If sig_ERexternalstop.wait(50) Then
        looping = 0
      End If
      System.Delay(100)
    Loop Until Memory.PrepCmd_MeasureInProgress = 0
    
    If Memory.measureOK = 0 Then                                                   
      'looping = 0
      'LED_Change 1
      'Exit Do
    End If    
    System.Delay(1500)
    'LED Off Cycle
    PrepareMeasureCRL    
    'LED_Change 0
    Do
      If sig_ERexternalstop.wait(50) Then
        looping = 0
      End If
      
      System.Delay(100)
    Loop Until Memory.PrepCmd_MeasureInProgress = 0
    If Memory.measureOK = 0 Then
      'looping = 0
      'LED_Change 1
      'Exit Do
    End If

    System.Delay(1500)

  Loop
  Memory.Free "sig_ERexternalstop"
  LogAdd "Endurance Run Stopped"
End Function
'------------------------------------------------------------------

'-------------------------------------------------------------------------
' Common Auxillary Functions ---------------------------------------------
'-------------------------------------------------------------------------

Function LED_Update ( Var_ID , OnOff )
		If OnOff = 1 Then
  		Visual.Select(Var_ID).Src = "./icon/led_green.png"
		ElseIf OnOff = 0 Then
      Visual.Select(Var_ID).Src = "./icon/led_black.png"
    Else
      Visual.Select(Var_ID).Src = "./icon/led_black.png"
		End If
End Function

'------------------------------------------------------------------

Function LED_Change ( Colour )
  Memory.CANData(0) = 1
  CANSendGetMC $(CMD_SEND_DATA),$(MC_STATUS_CANCEL),Memory.SLOT_NO,1,1
  Memory.CANData(0) = 1
  Memory.CANData(1) = Colour
  CANSendGetMC $(CMD_SEND_DATA),$(MC_SET_LED),Memory.SLOT_NO,1,2        
  Memory.CANData(0) = 1 
  CANSendGetMC $(CMD_SEND_DATA),$(MC_STATUS),Memory.SLOT_NO,1,1
End Function

'------------------------------------------------------------------

Function GetFloatCanData( param, displayelement )
  Dim Value,RawValue,CANData
  
  CANSendGetMC $(CMD_GET_DATA),param,Memory.SLOT_NO,1,0
  Memory.Get "CANData",CANData
  'DebugMessage String.Format("%4X,%4X,%4X,%4X",CANData(2),CANData(3),CANData(4),CANData(5))
  RawValue = Lang.MakeLong4(CANData(2),CANData(3),CANData(4),CANData(5))
  Value = String.Format("%G",Math.CastLong2Float(RawValue))
  Visual.Select(displayelement).Value = String.Format(Value)  
  GetFloatCanData = Value 
  
End Function

'-------------------------------------------------------------------
Function LogAdd ( sMessage )
  Dim Gridobj
  Set Gridobj = Visual.Script("LogGrid")
  Dim MsgId
  MsgId = Gridobj.uid()
  If NOT(sMessage = "") Then
    Gridobj.addRow MsgId, ""& FormatDateTime(Date, vbShortDate) &","& FormatDateTime(Time, vbShortTime)&":"& String.Format("%02d ", Second(Time)) &","& sMessage,0
    'Wish of SCM (automatically scroll to newest Msg)
    Gridobj.showRow( MsgId )
  End If  
  'DebugMessage sMessage
End Function

'-------------------------------------------------------------------
Function GetSCIErrorQueue ( ) 

  Dim exitloop
  Dim Debugmsg
  Dim loopcnt
  loopcnt = 32
  Debugmsg = "MB Errors: "
  exitloop = 0

  Do

    If CANSendGetMC($(CMD_GET_DATA),$(PARAM_MB_ERRORS), Memory.SLOT_NO,1,0) = True Then        
      'Exit if error = 0
      If Memory.CANData.Data(2) = 0 Then  
        exitloop = 1  
      Else
        Debugmsg = Debugmsg & String.Format ("%02X ",Memory.CANData.Data(2))
      End If
      GetSCITrace Log_SCI_TXRX, "Get MB Err:"
    End If   

    loopcnt = loopcnt - 1
    
    If loopcnt = 0 Then
      LogAdd "SCI Error Queue Overflow!"
      exitloop = 1
    'If no errors, no need to display message.
    End If  
  Loop Until exitloop = 1
  ' The last error was 00 = no error.
  If loopcnt = 19 Then
    Debugmsg = Debugmsg & "None"
  End If  
  LogAdd Debugmsg
End Function

'-------------------------------------------------------------------
Function GetSCITrace ( TxRxSetting, Log )
  Dim scitx,scirx,i,log_msg
  Dim Get_Tx, Get_Rx
  
  Get_Tx = False
  Get_Rx = False  
  
  If TxRxSetting = Log_SCI_TX OR TxRxSetting = Log_SCI_TXRX Then
    Get_Tx = True
  End If  
  If TxRxSetting = Log_SCI_RX OR TxRxSetting = Log_SCI_TXRX Then
    Get_Rx = True
  End If
  
  'Get TX sci data
  If Get_Tx = True Then
    scitx = ""
    GetSCIDataML 0  
    If Memory.SCIArray.size > 0 Then  
      scitx = scitx & Memory.SCIArray.size & " ("
      For i = 0 To Memory.SCIArray.size - 1
        scitx = scitx & String.Format ("%02X ",Memory.SCIArray.Data(i))
        If i > 50 Then
          DebugMessage "Data Overflow!"
          Exit For
        End If    
      Next  
    End If
    Log_SCIMsg "TX: " & scitx & ")"
  End If
  
  'Get RX sci data
  If Get_Rx = True Then
    scirx = ""
    GetSCIDataML 1    
    If Memory.SCIArray.size > 0 Then  
      scirx = scirx & Memory.SCIArray.size & " ("
      For i = 0 To Memory.SCIArray.size - 1
        scirx = scirx & String.Format ("%02X ",Memory.SCIArray.Data(i))
        If i > 50 Then
          DebugMessage "Data Overflow!"
          Exit For
        End If    
      Next
    End If
    Log_SCIMsg "RX: " & scirx & ")"
  End If

End Function
'-------------------------------------------------------------------

Function Log_SCIMsg( sMessage )
  Dim Gridobj
  Set Gridobj = Visual.Script("SCIGrid")
  Dim MsgId
  If NOT(sMessage = "") Then
    Gridobj.addRow MsgId, ""& FormatDateTime(Date, vbShortDate) &","& FormatDateTime(Time, vbShortTime)&":"& String.Format("%02d ", Second(Time)) &","& sMessage,0
    'Wish of SCM (automatically scroll to newest Msg)
    Gridobj.showRow( MsgId )
    DebugMessage "SCI: " & sMessage
  End If 
End Function

'-------------------------------------------------------------------

Function MF_Handle_Async_Msg_Standalone ( CanReadArg )
  If CanReadArg.Data(1) = $(PB_USER) Then
    If CanReadArg.Data(3) = 0 Then
      'LogAdd "Testing TX"
    ElseIf CanReadArg.Data(3) = 1 Then
      'LogAdd "Testing RX"    
      GetSCITrace Log_SCI_RX, ""
    End If
  End If
End Function

