
Const COMP_TYPE_RES  = 1
Const COMP_TYPE_CAP  = 2
Const COMP_TYPE_IND  = 3
Const COMP_TYPE_DIODE= 4
Const COMP_TYPE_PCAP = 5
Const COMP_TYPE_AUTO = 6

Const PROCESS_NONE = 0
Const PROCESS_SETUP = 1
Const PROCESS_MEASURE = 2
Const PROCESS_SELFTEST = 3
Const PROCESS_Calibration = 4

Const PREPARE_NONE = 0
Const PREPARE_SETUP = 1
Const PREPARE_MEASURE = 2
Const PREPARE_SELFTEST = 3
Const PREPARE_CALIBRATION = 4
Const PREPARE_AUTO = 6


Const TIO_SETUPMEASURE = 6000
Const TIO_MEASURE = 6000
Const TIO_CALIBRATE = 5000
Const TIO_SELFTEST = 5000

Function Init_MFCommand ( )
  Dim PrepCmd_Inprogress,PrepCmd_Error,PrepCmd_PrepID,Endurance_Inprogress
  
  MeasureChangeVisibility COMP_TYPE_RES
  ResultChangeVisibility PROCESS_NONE

  PrepCmd_Inprogress = 0
  Endurance_Inprogress = 0
  PrepCmd_Error = 0
  PrepCmd_PrepID = 1
  Memory.Set "PrepCmd_Inprogress",PrepCmd_Inprogress
  Memory.Set "PrepCmd_Error",PrepCmd_Error
  Memory.Set "Endurance_Inprogress",Endurance_Inprogress
  Memory.Set "PrepCmd_PrepID",PrepCmd_PrepID
  
  Visual.Select("ip_param_setupExpectedVal").Value = 100
  Visual.Select("ip_parampolarity").Value = 0
  Visual.Select("ip_paramnumofcycles").Value = 10
  Visual.Select("ip_paramresults").Value = 100
  
  Visual.Select("ip_param_setupExpectedVal").SetValidation VALIDATE_INPUT_MASK_R4,"Red",10
End Function
'------------------------------------------------------------------
Function WaitMeasure ( TimeOut )
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
        LogAdd "Measurement Complete"
        measureOK = 1
        loop_enable = 0
      End If
      
      Time = Time - 1
      If Time = 0 Then
        loop_enable = 0
        LogAdd "Measurement timeout"
      Else
        System.Delay(100)
      End If
    'End Loop Do while Memory.PrepCmd_Inprogress = 1     
    Loop Until loop_enable = 0 
    
    If measureOK = 1 Then
      ProcessResults
    End If
End Function 
'------------------------------------------------------------------
Function ProcessResults ( )
  Dim ResultLog,Value, CompType
  DebugMessage "Process Results"
  ResultChangeVisibility(Memory.PrepCmd)
  Select Case Memory.PrepCmd
    Case PREPARE_NONE : 
      ResultLog = "Error, no prepare"
    Case PREPARE_SETUP :
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_COMP_TYPE),SLOT_NO,1,0
      Value = Memory.CanData(2)
      Visual.Select("op_paramsetupcomptype").Value = String.Format("%c",Value)
      ResultLog = "MeasSetup: CompType :" & String.Format("%c",Value)
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_STRAY_CAPACITY),SLOT_NO,1,0
       Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_paramsetupcapacity").Value = Value
      ResultLog = ResultLog & " Capacity:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_U),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_paramU").Value = Value
      ResultLog = ResultLog & " U:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_I),SLOT_NO,1,0
       Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_paramI").Value = Value
      ResultLog = ResultLog & " I:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_PHI),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_paramPhi").Value = Value
      ResultLog = ResultLog & " Phi:" & Value      
      
    Case PREPARE_MEASURE :
      
      CompType = Visual.Select("optMeasureCommand").Value
      If Not CompType = 4 Then
      DebugMessage "Process Measure CRL, PCAP"
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_COMPONENT_TYPE),SLOT_NO,1,0
      Value = Memory.CanData(2)
      Visual.Select("op_parammeascomptype").Value = String.Format("%c",Value)
      ResultLog = "Meas: CompType :" & String.Format("%c",Value)      
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_parammeasValue").Value = Value
      ResultLog = ResultLog & " Value:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_U),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_parammeasU").Value = Value
      ResultLog = ResultLog & " U:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_I),SLOT_NO,1,0
       Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_parammeasI").Value = Value
      ResultLog = ResultLog & " I:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_PHI),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_parammeasPhi").Value = Value
      ResultLog = ResultLog & " Phi:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_FREQUENCY),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_parammeasFreq").Value = Value
      ResultLog = ResultLog & " Freq:" & Value      
      'Diode
      Else
      DebugMessage "Process Measure Diode"
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_FWD_VOLTAGE),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_paramfwdvoltage").Value = Value
      ResultLog = "Meas: FWDVoltage :" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_FWD_CURRENT),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_paramfwdcurrent").Value = Value
      ResultLog = ResultLog & " FWD Current:" & Value
      End If
      
      'End measure
    Case PREPARE_SELFTEST :
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SELFTEST_CONTACT_RES),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_paramSTcontactres").Value = Value
      ResultLog = ResultLog & " Value:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SELFTEST_CAPACITY_CM_ID),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_paramSTcapacitance").Value = Value
      ResultLog = ResultLog & " U:" & Value

    Case PREPARE_CALIBRATION :
    'NO param to read
    Case PREPARE_AUTO :
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_COMPONENT_TYPE),SLOT_NO,1,0
      Value = Memory.CanData(2)
      Visual.Select("op_parammeascomptype").Value = String.Format("%c",Value)
      ResultLog = "AutoMeas: CompType :" & String.Format("%c",Value)      
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_parammeasValue").Value = Value
      ResultLog = ResultLog & " Value:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_U),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_parammeasU").Value = Value
      ResultLog = ResultLog & " U:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_I),SLOT_NO,1,0
       Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_parammeasI").Value = Value
      ResultLog = ResultLog & " I:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_PHI),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_parammeasPhi").Value = Value
      ResultLog = ResultLog & " Phi:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_FREQUENCY),SLOT_NO,1,0
      Value = String.Format("%f",GetFloatCanData)
      Visual.Select("op_parammeasFreq").Value = Value
      ResultLog = ResultLog & " Freq:" & Value     
    End Select
    LogAdd ResultLog      
End Function 
'------------------------------------------------------------------
Function PrepareCommands (Cmd,Context,SlotNo,Division,DataLen,PubEndTimeout)
  Dim Command
  Command = PREPARE_NONE
  'This function is just to set the memory variable PrepCmd.
  'it is used by ProcessResults to determine which variables to read
  
  Select Case Cmd
    Case $(CMD_PREPARE_SETUP_MEASURE) : Command = PREPARE_SETUP
    Case $(CMD_PREPARE_MEASURE) : Command = PREPARE_MEASURE
    Case $(CMD_PREPARE_MEASURE_AUTO) : Command = PREPARE_AUTO
    Case $(CMD_PREPARE_SELFTEST) : Command = PREPARE_SELFTEST
    Case $(CMD_PREPARE_CALIBRATION) : Command = PREPARE_CALIBRATION
  End Select 
  Memory.Set "PrepCmd", Command
  PrepareCommands = CANSendPrepareCMD (Cmd,Context,SlotNo,Division,DataLen,PubEndTimeout)  
End Function

'------------------------------------------------------------------
Function OnClick_btnAssignCANID( Reason )
  Dim CanReadArg,CanID
  Set CanReadArg = CreateObject("ICAN.CanReadArg")
  
  CanID = CLng("&h" & Visual.Select("inputCANID").Value)
  'InitCAN CanID
  LogAdd "Assign CANID"
  CANID_Assign CanID
  System.Delay(100)
  Command_GetNumOfSlots
  GetFirmwareInfo
End Function

'------------------------------------------------------------------

Function OnClick_btn_setupmeasure( Reason )
Dim ExpectedValue ,CM_ID, CompType
  ExpectedValue = Math.CastFloat2Long(Visual.Select("ip_param_setupExpectedVal").Value)
  CM_ID = Visual.Select("opt_CMID").Value
  CompType = Visual.Select("optMeasureCommand").Value
  If NOT IsNumeric(ExpectedValue) Then
    LogAdd "Invalid value"
  Else    
    Command_Prepare_SetupMeasure CM_ID,ExpectedValue,CompType,TIO_SETUPMEASURE
  End If
End Function

'------------------------------------------------------------------

Function OnClick_btn_measure( Reason )
  Dim CompType
  DebugMessage "PrepareMeasure Button"
  CompType = Visual.Select("optMeasureCommand").Value
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
      LogAdd "PCAP not implemented"
    Case 6:
      LogAdd "Auto not implemented"    
  End Select
End Function

'------------------------------------------------------------------

Function PrepareMeasureCRL ( )

  Dim ExpectedValue, CompType,CM_ID
  ExpectedValue = Math.CastFloat2Long(Visual.Select("ip_param_setupExpectedVal").Value)
  CM_ID = Visual.Select("opt_CMID").Value  
  CompType = Visual.Select("optMeasureCommand").Value
  If NOT IsNumeric(ExpectedValue) Then
    LogAdd "Invalid value"
  Else
    Command_Prepare_Measure CM_ID,ExpectedValue,CompType,TIO_MEASURE
  End If

End Function
'------------------------------------------------------------------

Function PrepareMeasureDiode ( )

  Dim CompType,CM_ID, Current,Voltage,Polarity
  Voltage = Math.CastFloat2Long(Visual.Select("ip_paramvoltage").Value)
  Current = Math.CastFloat2Long(Visual.Select("ip_paramcurrent").Value)
  
  CM_ID = Visual.Select("opt_CMID").Value  
  CompType = Visual.Select("optMeasureCommand").Value
  If NOT IsNumeric(Voltage) Then
    LogAdd "Invalid Voltage value"    
  Elseif Not IsNumeric(Current) Then
    LogAdd "Invalid Current value"    
  Else
    Command_Prepare_Meas_FWDVOLTAGE CM_ID,Voltage,Current,CompType,TIO_MEASURE
  End If

End Function
'------------------------------------------------------------------

Function OnClick_btn_calibrate ( Reason )
  
  If PrepareCommands($(CMD_PREPARE_CALIBRATION),1,SLOT_NO,1,0,250) = True Then
    LogAdd "Calibration command started"
    System.Start "WaitMeasure",TIO_CALIBRATE
  Else
    LogAdd "Calibration Error."
  End If
End Function
'------------------------------------------------------------------

Function OnClick_btn_selftest ( Reason )
    
  If PrepareCommands($(CMD_PREPARE_SELFTEST),1,SLOT_NO,1,0,250) = True Then
    LogAdd "Self Test command started"
    System.Start "WaitMeasure",TIO_SELFTEST
  Else
    LogAdd "Self Test Error."
  End If
End Function

'------------------------------------------------------------------
Function OnChange_optMeasureCommand ( Reason )
  Dim CompType
  DebugMessage "Select:" & Visual.select("optMeasureCommand").Value

  Select Case  Visual.Select("optMeasureCommand").Value
  Case "1" : CompType = COMP_TYPE_RES
  Case "2" : CompType = COMP_TYPE_CAP
  Case "3" : CompType = COMP_TYPE_IND
  Case "4" : CompType = COMP_TYPE_DIODE
  Case "5" : CompType = COMP_TYPE_PCAP
  Case "6" : CompType = COMP_TYPE_AUTO
  End Select
  Memory.Set "CompType", CompType
  MeasureChangeVisibility CompType
  
End Function
'------------------------------------------------------------------
Function Command_Prepare_SetupMeasure (CM_ID, ExpectedValue,ComponentType,TimeOut)
  Memory.CANData(0) = Lang.GetByte(ExpectedValue,0)
  Memory.CANData(1) = Lang.GetByte(ExpectedValue,1)
  Memory.CANData(2) = Lang.GetByte(ExpectedValue,2)
  Memory.CANData(3) = Lang.GetByte(ExpectedValue,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_EXPECTED_RESULT),SLOT_NO,CM_ID,4
  
  Memory.CANData(0) = Lang.GetByte(ComponentType,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_COMPONENT_TYPE),SLOT_NO,CM_ID,1  
   
  If PrepareCommands($(CMD_PREPARE_SETUP_MEASURE),1,SLOT_NO,CM_ID,0,250) = True Then
    LogAdd "Setup Measure command started"
    System.Start "WaitMeasure",TimeOut
  Else
    LogAdd "Setup Measure Error."
  End If
End Function

'------------------------------------------------------------------
Function Command_Prepare_Measure (CM_ID,ExpectedValue,ComponentType,TimeOut)
  Memory.CANData(0) = Lang.GetByte(ExpectedValue,0)
  Memory.CANData(1) = Lang.GetByte(ExpectedValue,1)
  Memory.CANData(2) = Lang.GetByte(ExpectedValue,2)
  Memory.CANData(3) = Lang.GetByte(ExpectedValue,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_EXPECTED_RESULT),SLOT_NO,CM_ID,4
  
  Memory.CANData(0) = Lang.GetByte(ComponentType,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_COMPONENT_TYPE),SLOT_NO,CM_ID,1  
   
  If PrepareCommands($(CMD_PREPARE_MEASURE),1,SLOT_NO,CM_ID,0,250) = True Then
    LogAdd "Measure command started"
    System.Start "WaitMeasure",TimeOut
  Else
    LogAdd "Measure Command Error."
  End If
End Function

'------------------------------------------------------------------
Function Command_Prepare_Meas_FWDVOLTAGE (CM_ID,Current,Voltage,ComponentType,TimeOut)
  Memory.CANData(0) = Lang.GetByte(Current,0)
  Memory.CANData(1) = Lang.GetByte(Current,1)
  Memory.CANData(2) = Lang.GetByte(Current,2)
  Memory.CANData(3) = Lang.GetByte(Current,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_CURRENT),SLOT_NO,CM_ID,4
  
  Memory.CANData(0) = Lang.GetByte(Voltage,0)
  Memory.CANData(1) = Lang.GetByte(Voltage,1)
  Memory.CANData(2) = Lang.GetByte(Voltage,2)
  Memory.CANData(3) = Lang.GetByte(Voltage,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_VOLTAGE),SLOT_NO,CM_ID,4  
  
  Memory.CANData(0) = Lang.GetByte(ComponentType,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_COMPONENT_TYPE),SLOT_NO,CM_ID,1  
   
  If PrepareCommands($(CMD_PREPARE_MEASURE),1,SLOT_NO,CM_ID,0,250) = True Then
    LogAdd "Measure command started"
    System.Start "WaitMeasure",TimeOut
  Else
    LogAdd "Measure Command Error."
  End If
End Function
'------------------------------------------------------------------
Function Command_GetNumOfSlots( )
  If CANSendGetMC($(CMD_GET_DATA),$(MC_NUMBER_OF_SLOTS),SLOT_NO,1,0) = False Then
    LogAdd "Get Number of Slot command Error!"
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
    .Data(2) = 29
    .length = 3
    End If
  End With
  If Memory.Exists("CanManager") AND CanConfig.CANIDvalid = 1 Then    
    Memory.Get "CanManager",CanManager        
    DebugMessage "SendCmd:" & CanSendArg.Format
    Result = CanManager.SendCmd(CanSendArg,250,SC_CHECK_ERROR_BYTE,CanReadArg)
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
Function LogAdd ( sMessage )
  Dim Gridobj
  Set Gridobj = Visual.Script("LogGrid")
  Dim MsgId
  MsgId = Gridobj.uid()
  If NOT(sMessage = "") Then
    Gridobj.addRow MsgId, ""& FormatDateTime(Date, vbShortDate) &","& FormatDateTime(Time, vbShortTime)&":"& String.Format("%02d ", Second(Time)) &","& sMessage
    'Wish of SCM (automatically scroll to newest Msg)
    Gridobj.showRow( MsgId )
  End If  
  'DebugMessage sMessage
End Function

'------------------------------------------------------------------
Sub GetFirmwareInfo ( )
  Dim AppMaj,AppMin
  Dim App,Bios,App2
  If Command_GetFW($(PARAM_DL_ZIEL_APPL),AppMaj,AppMin) = 1 Then
    App = String.Format("%02X.%02X", AppMaj,AppMin)
  Else
    App = "??.??"
  End If
  
  If Command_GetFW($(PARAM_DL_ZIEL_BIOS),AppMaj,AppMin) = 1 Then
    Bios = String.Format("%02X.%02X", AppMaj,AppMin)
  Else
    Bios = "??.??"
  End If
  
  If Command_GetFW($(PARAM_DL_ZIEL_APPL_2),AppMaj,AppMin) = 1 Then
    App2 = String.Format("%02X.%02X", AppMaj,AppMin)
  Else
    App2 = "??.??"
  End If
  LogAdd "Firmware version: Bios:"& Bios & " App: " & App & " App2 " & App2
End Sub
'------------------------------------------------------------------
Function ResultChangeVisibility ( ProcessType )  
  Dim CompType
  Visual.Select("LayerResults").Style.Display = "None"
  'Set all to none
  Visual.Select("param_STcontactres").Style.Display = "None"
  Visual.Select("param_STCapacity").Style.Display = "None"
  Visual.Select("param_setupCompType").Style.Display = "None"
  Visual.Select("param_setupCapacity").Style.Display = "None"
  Visual.Select("param_setupU").Style.Display = "None"
  Visual.Select("param_setupI").Style.Display = "None"
  Visual.Select("param_setupPhi").Style.Display = "None"
  Visual.Select("param_setupFreq").Style.Display = "None"
  Visual.Select("param_measCompType").Style.Display = "None"  
  Visual.Select("param_measU").Style.Display = "None"  
  Visual.Select("param_measI").Style.Display = "None"  
  Visual.Select("param_measPhi").Style.Display = "None"  
  Visual.Select("param_measFreq").Style.Display = "None"  
  Visual.Select("param_measValue").Style.Display = "None"  
  Visual.Select("param_fwdvoltage").Style.Display = "None"  
  Visual.Select("param_fwdcurrent").Style.Display = "None"  
  Select Case ProcessType 
  Case PREPARE_NONE:
  'none
  Case PREPARE_SETUP:
  Visual.Select("param_setupCompType").Style.Display = "Block"
  Visual.Select("param_setupCapacity").Style.Display = "Block"
  Visual.Select("param_setupU").Style.Display = "Block"
  Visual.Select("param_setupI").Style.Display = "Block"
  Visual.Select("param_setupPhi").Style.Display = "Block"
  Case PREPARE_MEASURE:
  CompType = Visual.Select("optMeasureCommand").Value
  'Not Diode
  If Not CompType = 4 Then
  Visual.Select("param_measU").Style.Display = "Block"
  Visual.Select("param_measI").Style.Display = "Block"
  Visual.Select("param_measPhi").Style.Display = "Block"
  Visual.Select("param_measFreq").Style.Display = "Block" 
  Visual.Select("param_measValue").Style.Display = "Block" 
  Else
  Visual.Select("param_fwdvoltage").Style.Display = "Block"
  Visual.Select("param_fwdcurrent").Style.Display = "Block"  
  End If
  Case PREPARE_SELFTEST:  
  Visual.Select("param_STcontactres").Style.Display = "Block"
  Visual.Select("param_STCapacity").Style.Display = "Block"
  End Select  
  Visual.Select("LayerResults").Style.Display = "Block"
End Function
'------------------------------------------------------------------
Function MeasureChangeVisibility ( CompType )
  DebugMessage "Change:" & CompType
  'Set all fields to none
    Visual.Select("param_voltage").Style.Display = "None"
    Visual.Select("param_current").Style.Display  = "None"
    Visual.Select("param_result").Style.Display  =  "None"
    Visual.Select("param_maxvoltage").Style.Display  = "None"
    Visual.Select("param_numofcycle").Style.Display  = "None"
    Visual.Select("param_polarity").Style.Display  = "None"  
  Select Case CompType
  'Res
  Case 1:
    Visual.Select("param_result").Style.Display  = "block"
    Visual.Select("param_numofcycle").Style.Display  = "block"
  'Cap
  Case 2:
    Visual.Select("param_result").Style.Display  = "block"
    Visual.Select("param_numofcycle").Style.Display  = "block"
  'Inductor
  Case 3:
    Visual.Select("param_result").Style.Display  = "block"
    Visual.Select("param_numofcycle").Style.Display  = "block"
  'Diode
  Case 4:
    Visual.Select("param_voltage").Style.Display = "block"
    Visual.Select("param_current").Style.Display  = "block"
    Visual.Select("param_polarity").Style.Display  = "block"
  'PCAP
  Case 5:
    Visual.Select("param_result").Style.Display  =  "block"
    Visual.Select("param_maxvoltage").Style.Display  = "block"
    Visual.Select("param_polarity").Style.Display  = "block"
  Case 6:
    'case auto: all none
  End Select
End Function 
