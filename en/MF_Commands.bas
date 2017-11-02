
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
Const PREPARE_ALL = 6


Const TIO_SETUPMEASURE = 6000
Const TIO_MEASURE = 6000
Const TIO_CALIBRATE = 5000
Const TIO_SELFTEST = 5000

Function Init_MFCommand ( )
  Dim PrepCmd_Inprogress,PrepCmd_Error,PrepCmd_PrepID,Endurance_Inprogress
  Dim PrepCmd_MeasureInProgress,counter
  ChangeVisibility_ComponentSelect COMP_TYPE_RES
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
  Visual.Select("ip_paramnumofcycles").Value = 0
  Visual.Select("ip_paramresults").Value = 100
  Visual.Select("ip_paramvoltage").Value = 3
  Visual.Select("ip_paramcurrent").Value = 0.01
  
  For counter = 1 To 60
    Visual.Select("opt_SlotNum").addItem counter,counter
  Next
  Memory.Set "SLOT_NO",Visual.Select("opt_SlotNum").SelectedItemAttribute("Value")
  DebugMessage "SLOT_NO :" & Memory.SLOT_NO
    
  Visual.Select("ip_param_setupExpectedVal").SetValidation VALIDATE_INPUT_MASK_R4,"Red",10
  
  PrepCmd_MeasureInProgress = 0
  Memory.Set "PrepCmd_MeasureInProgress",PrepCmd_MeasureInProgress
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
  LogAdd "Firmware version: Bios:"& Bios & " App: " & App & " App2: " & App2
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
        Command_ChangeLED 1
      Else
        System.Delay(100)
      End If
    'End Loop Do while Memory.PrepCmd_Inprogress = 1     
    Loop Until loop_enable = 0 
    
    If measureOK = 1 Then
      Get_Measurements
    End If    
    Memory.Set "measureOK",measureOK
    Memory.PrepCmd_MeasureInProgress = 0
End Function 
'------------------------------------------------------------------
Function Get_Measurements ( )
  Dim ResultLog,Value, CompType
  DebugMessage "Process Results"
  ResultChangeVisibility(Memory.PrepCmd)
  Select Case Memory.PrepCmd
    Case PREPARE_NONE : 
      ResultLog = "Error, no prepare"
    Case $(CMD_PREPARE_SETUP_MEASURE) :
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_CONTACT_RES),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramContactres").Value = String.Format("%G",Value)
      ResultLog = "Setup Measure: Contact Res :" & String.Format("%c",Value)
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_CAPACITY),Memory.SLOT_NO,1,0
       Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramCapacityCM").Value = Value
      ResultLog = ResultLog & " Cap:" & Value

      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_RESISTANCE),Memory.SLOT_NO,1,0
       Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramResistanceCM").Value = Value
      ResultLog = ResultLog & " Res:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_U),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramU").Value = Value
      ResultLog = ResultLog & " U:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_I),Memory.SLOT_NO,1,0
       Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramI").Value = Value
      ResultLog = ResultLog & " I:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_PHI),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramPhi").Value = Value
      ResultLog = ResultLog & " Phi:" & Value 

      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SETUP_MEAS_FREQ),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramFreq").Value = Value
      ResultLog = ResultLog & " Freq:" & Value      
      
    Case $(CMD_PREPARE_MEASURE) :
      
      CompType = Visual.Select("optMeasureCommand").Value
      If Not CompType = 4 Then
      DebugMessage "Process Measure CRL, PCAP"
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_COMPONENT_TYPE1),Memory.SLOT_NO,1,0
      Value = Memory.CanData(2)
      Visual.Select("op_paramcomptype1").Value = String.Format("%c",Value)
      ResultLog = "Meas: CompType1 :" & String.Format("%c",Value)      
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE1),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType1Value").Value = Value
      ResultLog = ResultLog & " Value:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE_MIN1),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType1ValueMin").Value = Value
      ResultLog = ResultLog & " Min:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE_MAX1),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType1ValueMax").Value = Value
      ResultLog = ResultLog & " Max :" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_COMPONENT_TYPE2),Memory.SLOT_NO,1,0
      Value = Memory.CanData(2)
      Visual.Select("op_paramcomptype2").Value = String.Format("%c",Value)
      ResultLog = "CompType2 :" & String.Format("%c",Value)      
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE2),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType2Value").Value = Value
      ResultLog = ResultLog & " Value:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE_MIN2),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType2ValueMin").Value = Value
      ResultLog = ResultLog & " Min:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE_MAX2),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType2ValueMax").Value = Value
      ResultLog = ResultLog & " Max :" & Value      
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_U),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramU").Value = Value
      ResultLog = ResultLog & " U:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_I),Memory.SLOT_NO,1,0
       Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramI").Value = Value
      ResultLog = ResultLog & " I:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_PHI),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramPhi").Value = Value
      ResultLog = ResultLog & " Phi:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_FREQUENCY),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramFreq").Value = Value
      ResultLog = ResultLog & " Freq:" & Value      
      'Diode
      Else
      DebugMessage "Process Measure Diode"
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_FWD_VOLTAGE),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramfwdvoltage").Value = Value
      ResultLog = "Meas: FWD Voltage :" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_FWD_CURRENT),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramfwdcurrent").Value = Value
      ResultLog = ResultLog & " FWD Current:" & Value
      End If
      
      'End measure
    Case $(CMD_PREPARE_SELFTEST) :
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SELFTEST_CONTACT_RES),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramContactres").Value = Value
      ResultLog = ResultLog & " Contact Res:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SELFTEST_CAPACITY),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramCapacityCM").Value = Value
      ResultLog = ResultLog & " Cap:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SELFTEST_RESISTANCE),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramResistanceCM").Value = Value
      ResultLog = ResultLog & " Res:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SELFTEST_U),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramU").Value = Value
      ResultLog = ResultLog & " U:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SELFTEST_I),Memory.SLOT_NO,1,0
       Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramI").Value = Value
      ResultLog = ResultLog & " I:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SELFTEST_PHI),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramPhi").Value = Value
      ResultLog = ResultLog & " Phi:" & Value 

      CANSendGetMC $(CMD_GET_DATA),$(PARAM_SELFTEST_FREQ),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramFreq").Value = Value
      ResultLog = ResultLog & " Freq:" & Value    
      
      
    Case $(CMD_PREPARE_CALIBRATION) :
    'NO param to read
    Case $(CMD_PREPARE_MEASURE_AUTO) :
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_COMPONENT_TYPE1),Memory.SLOT_NO,1,0
      Value = Memory.CanData(2)
      Visual.Select("op_paramcomptype1").Value = String.Format("%c",Value)
      ResultLog = "Cal: CompType1 :" & String.Format("%c",Value)      
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE1),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType1Value").Value = Value
      ResultLog = ResultLog & " Value:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE_MIN1),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType1ValueMin").Value = Value
      ResultLog = ResultLog & " Min:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE_MAX1),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType1ValueMax").Value = Value
      ResultLog = ResultLog & " Max :" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_COMPONENT_TYPE2),Memory.SLOT_NO,1,0
      Value = Memory.CanData(2)
      Visual.Select("op_paramcomptype2").Value = String.Format("%c",Value)
      ResultLog = "CompType2 :" & String.Format("%c",Value)      
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE2),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType2Value").Value = Value
      ResultLog = ResultLog & " Value:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE_MIN2),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType2ValueMin").Value = Value
      ResultLog = ResultLog & " Min:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_VALUE_MAX2),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramType2ValueMax").Value = Value
      ResultLog = ResultLog & " Max :" & Value      
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_U),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramU").Value = Value
      ResultLog = ResultLog & " U:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_I),Memory.SLOT_NO,1,0
       Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramI").Value = Value
      ResultLog = ResultLog & " I:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_PHI),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramPhi").Value = Value
      ResultLog = ResultLog & " Phi:" & Value
      
      CANSendGetMC $(CMD_GET_DATA),$(PARAM_MEAS_FREQUENCY),Memory.SLOT_NO,1,0
      Value = String.Format("%G",GetFloatCanData)
      Visual.Select("op_paramFreq").Value = Value
      ResultLog = ResultLog & " Freq:" & Value      
    End Select
    LogAdd ResultLog      
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
  'Set Green LED
  Memory.CANData(0) = 1 
  Memory.CANData(1) = 2 
  CANSendGetMC $(CMD_SEND_DATA),$(MC_SET_LED),Memory.SLOT_NO,1,2
  Memory.CANData(0) = 1 
  CANSendGetMC $(CMD_SEND_DATA),$(MC_STATUS),Memory.SLOT_NO,1,1
  
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

Sub OnClick_ButtonGridClear( Reason )
  Visual.Script("LogGrid").clearAll()
End Sub
'------------------------------------------------------------------
Sub OnClick_btnGetApp( Reason )  
  GetFirmwareInfo
End Sub

'------------------------------------------------------------------
Function OnClick_btnReset ( Reason )
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
    LogAdd "Resetting XMF!" 
  End if
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

  Dim ExpectedValue, CompType,NumofCycles,CM_ID
  ExpectedValue = Math.CastFloat2Long(Visual.Select("ip_param_setupExpectedVal").Value)
  CM_ID = Visual.Select("opt_CMID").Value  
  CompType = Visual.Select("optMeasureCommand").Value
  NumofCycles = Visual.Select("ip_paramnumofcycles").Value
  If NOT IsNumeric(ExpectedValue) Then
    LogAdd "Invalid value"
  Else
    Command_Prepare_Measure CM_ID,ExpectedValue,CompType,NumofCycles,TIO_MEASURE
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
  Dim SubCmd
   SubCmd = Visual.Select("opt_SubCmd").Value
   Memory.CANData(0) = SubCmd
  If CANSendPrepareCMD($(CMD_PREPARE_CALIBRATION),1,Memory.SLOT_NO,1,1,250) = True Then
    Memory.Set "PrepCmd", $(CMD_PREPARE_CALIBRATION)
    LogAdd "Calibration command started"
    System.Start "Wait_Measurement",TIO_CALIBRATE
  Else
    LogAdd "Calibration Error."
  End If
End Function
'------------------------------------------------------------------

Function OnClick_btn_selftest ( Reason )
    
  If CANSendPrepareCMD($(CMD_PREPARE_SELFTEST),1,Memory.SLOT_NO,1,0,250) = True Then
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

Function OnClick_btn_endu_stop ( Reason )
If Memory.Exists("sig_ERexternalstop") Then
  LogAdd "Endurance Run Stopping..."
  Memory.sig_ERexternalstop.Set
Else
  LogAdd "No Endurance run to stop."
End If
End Function

'------------------------------------------------------------------
Function OnChange_opt_SlotNum ( Reason )
  Memory.Set "SLOT_NO",Visual.Select("opt_SlotNum").SelectedItemAttribute("Value")
  DebugMessage "SLOT_NO" & Memory.SLOT_NO
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
  ChangeVisibility_ComponentSelect CompType
  
End Function
'------------------------------------------------------------------
Function Command_Prepare_SetupMeasure (CM_ID, ExpectedValue,ComponentType,TimeOut)
  Memory.CANData(0) = Lang.GetByte(ExpectedValue,0)
  Memory.CANData(1) = Lang.GetByte(ExpectedValue,1)
  Memory.CANData(2) = Lang.GetByte(ExpectedValue,2)
  Memory.CANData(3) = Lang.GetByte(ExpectedValue,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_EXPECTED_RESULT),Memory.SLOT_NO,CM_ID,4
  
  Memory.CANData(0) = Lang.GetByte(ComponentType,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_COMPONENT_TYPE),Memory.SLOT_NO,CM_ID,1  
   
  If CANSendPrepareCMD($(CMD_PREPARE_SETUP_MEASURE),1,Memory.SLOT_NO,CM_ID,0,250) = True Then
    Memory.Set "PrepCmd", $(CMD_PREPARE_SETUP_MEASURE)
    LogAdd "Setup Measure command started"
    System.Start "Wait_Measurement",TimeOut
  Else
    LogAdd "Setup Measure Error."
  End If
End Function

'------------------------------------------------------------------
Function Command_Prepare_Measure (CM_ID,ExpectedValue,ComponentType,NumofCycles,TimeOut)
  Memory.CANData(0) = Lang.GetByte(ExpectedValue,0)
  Memory.CANData(1) = Lang.GetByte(ExpectedValue,1)
  Memory.CANData(2) = Lang.GetByte(ExpectedValue,2)
  Memory.CANData(3) = Lang.GetByte(ExpectedValue,3)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_EXPECTED_RESULT),Memory.SLOT_NO,CM_ID,4
  
  Memory.CANData(0) = Lang.GetByte(ComponentType,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_COMPONENT_TYPE),Memory.SLOT_NO,CM_ID,1
   Memory.CANData(0) = Lang.GetByte(NumofCycles,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_NUM_OF_CYCLES),Memory.SLOT_NO,CM_ID,1  
  If CANSendPrepareCMD($(CMD_PREPARE_MEASURE),1,Memory.SLOT_NO,CM_ID,0,250) = True Then
    If Not Memory.Exists("sig_ERexternalstop") Then        
      LogAdd "Measure command started"
    End If
    Memory.Set "PrepCmd", $(CMD_PREPARE_MEASURE)
    Memory.PrepCmd_MeasureInProgress = 1
    System.Start "Wait_Measurement",TimeOut
  Else
    LogAdd "Measure Command Error."
    Command_ChangeLED 1
  End If
End Function

'------------------------------------------------------------------
Function Command_Prepare_Meas_FWDVOLTAGE (CM_ID,Current,Voltage,ComponentType,TimeOut)
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
  
  Memory.CANData(0) = Lang.GetByte(ComponentType,0)
  CANSendGetMC $(CMD_SEND_DATA),$(PARAM_COMPONENT_TYPE),Memory.SLOT_NO,CM_ID,1  
   
  If CANSendPrepareCMD($(CMD_PREPARE_MEASURE),1,Memory.SLOT_NO,CM_ID,0,250) = True Then
    Memory.Set "PrepCmd", $(CMD_PREPARE_MEASURE)
    LogAdd "Measure command started"
    System.Start "Wait_Measurement",TimeOut
  Else
    LogAdd "Measure Command Error."
  End If
End Function
'------------------------------------------------------------------
Function Command_GetNumOfSlots( )
  If CANSendGetMC($(CMD_GET_DATA),$(MC_NUMBER_OF_SLOTS),Memory.SLOT_NO,1,0) = False Then
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

Function Command_ChangeLED ( Colour )
    Memory.CANData(0) = 1 
    CANSendGetMC $(CMD_SEND_DATA),$(MC_STATUS_CANCEL),Memory.SLOT_NO,1,1
    Memory.CANData(0) = 1 
    Memory.CANData(1) = Colour
    CANSendGetMC $(CMD_SEND_DATA),$(MC_SET_LED),Memory.SLOT_NO,1,2        
    Memory.CANData(0) = 1 
    CANSendGetMC $(CMD_SEND_DATA),$(MC_STATUS),Memory.SLOT_NO,1,1
End Function




Function ResultChangeVisibility ( ProcessType )  
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
  Select Case ProcessType 
  Case PREPARE_NONE:
  'none
  Case $(CMD_PREPARE_SETUP_MEASURE):
  Visual.Select("param_ContactRes").Style.Display = "Block"
  Visual.Select("param_CapacityCM").Style.Display = "Block"
  Visual.Select("param_ResistanceCM").Style.Display = "Block"
  Visual.Select("param_U").Style.Display = "Block"
  Visual.Select("param_I").Style.Display = "Block"
  Visual.Select("param_Phi").Style.Display = "Block"
  Visual.Select("param_Freq").Style.Display = "Block"
  Case $(CMD_PREPARE_MEASURE):
  CompType = Visual.Select("optMeasureCommand").Value
  'Not Diode
  If Not CompType = 4 Then
    Visual.Select("param_CompType1").Style.Display = "Block"
    Visual.Select("param_CompType2").Style.Display = "Block"
    Visual.Select("param_Type1Value").Style.Display = "Block"
    Visual.Select("param_Type1ValueMin").Style.Display = "Block"
    Visual.Select("param_Type1ValueMax").Style.Display = "Block"
    Visual.Select("param_Type2Value").Style.Display = "Block"  
    Visual.Select("param_Type2ValueMin").Style.Display = "Block"  
    Visual.Select("param_Type2ValueMax").Style.Display = "Block"  
  Else
    Visual.Select("param_fwdvoltage").Style.Display = "Block"
    Visual.Select("param_fwdcurrent").Style.Display = "Block"  
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
  
  End Select
  Visual.Select("LayerResults").Style.Display = "Block"
End Function
'------------------------------------------------------------------
Function ChangeVisibility_ComponentSelect ( CompType )
  'DebugMessage "Change:" & CompType
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
    Command_ChangeLED 2
    PrepareMeasureCRL    
    Do    
      If sig_ERexternalstop.wait(50) Then
        looping = 0
      End If
      System.Delay(100)
    Loop Until Memory.PrepCmd_MeasureInProgress = 0
    
    If Memory.measureOK = 0 Then                                                   
      'looping = 0
      Command_ChangeLED 1
      'Exit Do
    End If    
    System.Delay(1500)
    'LED Off Cycle
    PrepareMeasureCRL    
    Command_ChangeLED 0
    Do
      If sig_ERexternalstop.wait(50) Then
        looping = 0
      End If
      
      System.Delay(100)
    Loop Until Memory.PrepCmd_MeasureInProgress = 0
    If Memory.measureOK = 0 Then
      'looping = 0
      Command_ChangeLED 1
      'Exit Do
    End If

    System.Delay(1500)

  Loop
  Memory.Free "sig_ERexternalstop"
  LogAdd "Endurance Run Stopped"
End Function
'------------------------------------------------------------------

'------------------------------------------------------------------
' Auxillary Functions ---------------------------------------------
'------------------------------------------------------------------
Function GetFloatCanData( )
  Dim Value,RawValue,CANData
  Memory.Get "CANData",CANData
  'DebugMessage String.Format("%4X,%4X,%4X,%4X",CANData(2),CANData(3),CANData(4),CANData(5))
  RawValue = Lang.MakeLong4(CANData(2),CANData(3),CANData(4),CANData(5))
  Value = Math.CastLong2Float(RawValue)
  GetFloatCanData = Value
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
