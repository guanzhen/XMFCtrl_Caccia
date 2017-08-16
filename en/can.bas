
'These are definations of the CAN commands, as per the interface specification
Const CANCMD_SETBRAKE           = &H02
Const CANCMD_Command_AxisRefRun = &h04
Const CANCMD_test = &h06

Const SLOT_NO = 1

Function btn_CanConnect( id, id1 )
  Dim Net,Baud,CANConfig, CANID,CANData,i
  Dim text
  set CANConfig = Object.CreateRecord( "Net", "CANIDcmd", "CANIDAck", "CANIDPub","Baudrate","Config","CANIDvalid" )
  Set CANData = CreateObject( "MATH.Array" )
  
  For i = 0  To 7
  	CANData.Add(0)
  Next
  Memory.Set "CANData",CANData
  
  DebugMessage"Launch Can Connect"
  CANConfig.Net = Visual.Script("opt_net")
  CANConfig.Config = Visual.Script("opt_config")
  CANConfig.CANIDcmd = Visual.Script("input_CANID")

If CANConfig.Config = 0 Then
    CANConfig.Baudrate = "250"
    text = "[Standalone]"
    Visual.Select("inputCANID").Value = "608"
  Elseif CANConfig.Config = 1 Then
    CANConfig.Baudrate = "1000"
    Visual.Select("inputCANID").Value = "500"
    text = "[XFCU]"
    Visual.Select("btnAssignCANID").Style.Display  = "none"
  End If
  
  DebugMessage "Selected Config :"&CANConfig.Config
  DebugMessage "Selected CANConfig.Net :"&CANConfig.Net + 1
  
  Window.Title = Window.Title & text & " - Net " & CANConfig.Net + 1
  CanConfig.CANIDvalid = 1
  
  Memory.Set "CANConfig",CANConfig
  'Initialise can using the settings by user.
  CANID = CLng("&h" & Visual.Select("inputCANID").Value)
  DebugMessage "CANID: (d)" & CANID
  If InitCAN(CANID,CANConfig.Config) = 1 Then
    Visual.Script("dhxWins").unload()
    Visual.Select("Layer_CanSetup").Style.Display = "none"
    Visual.Select("Layer_MessageLog").Style.Display = "block"
    Visual.Select("Layer_TabStrip").Style.Display = "block"    
  Else
    MsgBox "Cannot Open CAN Manager on Net " & CANConfig.Net + 1 & "!"
  End If
  
End Function


'**********************************************************************
'* Purpose: Init CAN module listening to Async and Pub Messages (0x408,0x008)
'* Input:  none
'* Output: none
'**********************************************************************
Function InitCAN ( CANID, Config )
  Dim CanManager, CanManagerPUB, CANConfig, CanSendArg,CanReadArg

  Set CANSendArg = CreateObject("ICAN.CanSendArg")
  Set CANReadArg = CreateObject("ICAN.CanReadArg") 

  Memory.Get "CANConfig",CANConfig
  DeleteCanManager CANConfig.Net,True
  Set CanManager = LaunchCanManager( CANConfig.Net, CANConfig.BaudRate )
  If Lang.IsObject(CanManager) Then  
    CanManager.Events = True
    CanManager.Deliver = True
    If Config = 0 Then
    DebugMessage "Platform = 3"
    CanManager.Platform = 3
    Else
    DebugMessage "Platform 1"
    End If
    CanManager.SetArbitrationOrder CAN_ARBITRATION_SYNCHRONOUS
    'WithEvents.ConnectObject CanManager, "CanManager_"
    
    If Config = 0 Then
      Set CanManagerPUB = Memory.CanManager.Clone
      CanManagerPUB.Events = True
      CanManagerPUB.Deliver = True
      CanManagerPUB.Platform = 3
      CanManagerPUB.SetArbitrationOrder CAN_ARBITRATION_ASYNCHRONOUS
      WithEvents.ConnectObject CanManagerPUB, "CanManagerPUB_"
      Memory.Set "CanManagerPUB",CanManagerPUB
    Else
      Set CanManagerPUB = Memory.CanManager.Clone
      CanManagerPUB.Events = True
      CanManagerPUB.Deliver = True
      CanManagerPUB.SetArbitrationOrder CAN_ARBITRATION_ASYNCHRONOUS
      WithEvents.ConnectObject CanManagerPUB, "CanManagerPUB_"
      Memory.Set "CanManagerPUB",CanManagerPUB    
    End If
    
    CANID_Set CANID
    InitCAN = 1
  Else
    InitCAN = 0
    LogAdd "No CAN Manager!"
  End If
End Function

'------------------------------------------------------------------
Function CANID_Assign (CANID)
  Dim CanSendArg,CanReadArg
  Dim CanManager
  Set CanSendArg = CreateObject("ICAN.CanSendArg")
  Set CanReadArg = CreateObject("ICAN.CanReadArg")

  If Memory.Exists( "CanManager" ) Then
    Memory.Get "CanManager",CanManager
    CanSendArg.CanId = &h00
    CanSendArg.Data(0) = &h00
    CanSendArg.Data(1) = &h02
    CanSendArg.Data(2) = &h00
    CanSendArg.Data(3) = &h00
    CanSendArg.Data(4) = Lang.GetByte(CANID,0)
    CanSendArg.Data(5) = Lang.GetByte(CANID,1)
    CanSendArg.Data(6) = &h22
    CanSendArg.Length = 7
    CanManager.Send CanSendArg
  Else
    DebugMessage "No CAN Manager!"
  End if
End Function

'------------------------------------------------------------------

Function CANID_Validate(CanID)
  Dim xCanID
  xCanID = CanID
  If Memory.CANConfig.Config = 1 Then
    xCanID = xCanID - &h400
    DebugMessage "CANID: " & xCanID
    If xCanID < 0 Or xCanID > &h118 Or xCanID Mod 8 > 0 Then
      MsgBox "Can ID for XFCU is not valid, the legal value is 0x400, 0x408, 0x410, 0x418, 0x500, 0x508, 0x510, 0x518."
      CANID_Validate = False
    Else
      CANID_Validate = True
    End If
  Else 'UseEDIF
    If CanID < &h608 Or CanID > &h7F8 Or CanID Mod 8 > 0Then
      MsgBox "Can ID for CAN2/SS-EDIF is not valid, the legal range is 0x608...0x7F8 and in multiples of 0x008"
      CANID_Validate = False
    Else
      CANID_Validate = True
    End If
  End If
End Function

'------------------------------------------------------------------

Function CANID_Set ( CANID )
  Dim  CANConfig
  Memory.Get "CANConfig",CANConfig
  '0 = Standalone : CAN ID Cmd: 0x6e8 Ack: 0x4e8 Pub: 0xe8
  '1 = In Machine : CAN ID Cmd: 0x500 Ack: 0x501 Pub: 0x???
  If CANConfig.Config = 0 Then 
    With CANConfig
      .CANIDcmd = CANID
      .CANIDAck = CANID - &h200
      .CANIDPub = CANID - &h600
    End With
  Else
      With CANConfig
      .CANIDcmd = CANID
      .CANIDAck = CANID + 1
      .CANIDPub = CANID + 3      
    End With
  End If
  Memory.Set "CANConfig",CANConfig
  DebugMessage "Set CANID  with settings: Net:" &CANConfig.Net&" BaudRate:"& CANConfig.BaudRate&" CANID:" & String.Format("0x%03X",CANConfig.CANIDcmd)

  'Remove all existing funnels
  Memory.CanManager.ChangeFunnel "0-0x7ff",false
  'Memory.CanManager.ChangeFunnel String.Format("%d",CANConfig.CANIDAck), True
  Memory.CanManager.ChangeFunnel String.Format("%d,%d",CANConfig.CANIDAck,CANConfig.CANIDPub), True
  DebugMessage "CanManager1 FunnelSet: " &  Memory.CanManager.QueryFunnel(16)
  
  If CANConfig.Config = 0 Then
    'Remove all existing funnels
    Memory.CanManagerPUB.ChangeFunnel "0-0x7ff",false
    Memory.CanManagerPUB.ChangeFunnel String.Format("%d",CANConfig.CANIDPub),True
    DebugMessage "CanManager2 FunnelSet: " & Memory.CanManagerPub.QueryFunnel(16)
    
  Else
    Memory.CanManagerPUB.ChangeFunnel "0-0x7ff",false  
    Memory.CanManagerPUB.ChangeFunnel String.Format("%d",CANConfig.CANIDPub),True
    DebugMessage "CanManager2 FunnelSet: " & Memory.CanManagerPub.QueryFunnel(16)
  End If
  
End Function 
'------------------------------------------------------------------

Sub CanManagerPUB_Deliver( CanReadArg )
  'Ignore error frames
  If Not CanReadArg.Type = CAN_MSG_ERROR_FRAME Then
    If Memory.CANConfig.Config = 0 Then
      Handle_PubMsg_Standalone CanReadArg
    Else
      Handle_PubMsg_FCU CANReadArg
    End If
  End If
End Sub
'------------------------------------------------------------------
Function Handle_PubMsg_FCU( CanReadArg )
  Dim Bit_Active,Bit_Start,Bit_Error,Bit_Async,Bit_Posflag
  Dim OutputDebug,OutputLog
  Dim QuitEndurance  
  Dim PrepareStopTime
  Bit_Active = Lang.Bit(CanReadArg.Data(0),7)
  Bit_Async = Lang.Bit(CanReadArg.Data(0),6)
  Bit_Posflag = Lang.Bit(CanReadArg.Data(0),2)
  Bit_Error = Lang.Bit(CanReadArg.Data(0),0)
  
  If Bit_Async = 1 Then
    OutputDebug = "Feeder Direct Msg "
    If Bit_Posflag = 1 Then
      OutputDebug = OutputDebug & "Posflag "
      PrepareStopTime = CANReadArg.TimeStamp
      Memory.Set "PrepareStopTime",PrepareStopTime
      UpdateCycleTime (PrepareStopTime)
    End If
  Else
    OutputDebug = "Feeder Pub Msg "
  End If
  
  If Bit_Error = 1 Then
    OutputDebug = OutputDebug & "NOK:" & GetErrorInfo( CanReadArg )
    OutputLog = OutputLog & "NOK:" & GetErrorInfo( CanReadArg )
  Else
    OutputDebug = OutputDebug & "OK ("  & CanReadArg.Format(CFM_SHORT) & ")"
    OutputLog = OutputLog & "OK ("  & CanReadArg.Format(CFM_SHORT) & ")"
  End If    
  
  If Memory.PrepCmd_Inprogress = 1 Then
    DebugMessage "Prepare Command in progress"
    If Bit_Active = 0 AND Bit_Async = 1 Then
      If Memory.Endurance_InProgress = 1 Then
        DebugMessage "Endurance run in progress"
        If Bit_Error = 1 Then
          LogAdd "Prepare command completed with error : " & GetErrorInfo( CanReadArg )
        Else
          LogAdd "Prepare command completed"
        End If
      End If
      Memory.PrepCmd_Inprogress = 0
    Else
      DebugMessage "Endurance not in progress"
        
    End If
    
    Memory.PrepCmd_Error = Bit_Error
    If Bit_Error = 1 Then
      Memory.Set "CanErr",CanReadArg.Data(1)
    End If

  End If
  
  DebugMessage OutputDebug
End Function 
'------------------------------------------------------------------
Function Handle_PubMsg_Standalone( CanReadArg )
  Dim Bit_Active,Bit_Start,Bit_Error,Bit_Async
  Dim OutputDebug,OutputLog
  Dim QuitEndurance  
  Dim PrepareStopTime
  Bit_Active = Lang.Bit(CanReadArg.Data(0),7)
  Bit_Async = Lang.Bit(CanReadArg.Data(0),6)
  Bit_Start = Lang.Bit(CanReadArg.Data(0),4)
  Bit_Error = Lang.Bit(CanReadArg.Data(0),0)

  If Bit_Async = 1 Then
    OutputDebug = "Async Pub Msg "
    If CanReadArg.Data(1) = &hC8 Then
      PrepareStopTime = CANReadArg.TimeStamp
      Memory.Set "PrepareStopTime",PrepareStopTime
      UpdateCycleTime (PrepareStopTime)
    End If
  Else
    OutputDebug = "Sync Pub Msg "
    If Bit_Start = 1 Then
      OutputDebug = OutputDebug & "start: "
      OutputLog = OutputLog & "Operation started  "
    Else
      OutputDebug = OutputDebug & "end: "
      OutputLog = OutputLog & "Operation ended  "
    End If
  End If
  
  If Bit_Error = 1 Then
    OutputDebug = OutputDebug & "NOK:" & GetErrorInfo( CanReadArg )
    OutputLog = OutputLog & "NOK:" & GetErrorInfo( CanReadArg )
  Else
    OutputDebug = OutputDebug & "OK ("  & CanReadArg.Format(CFM_SHORT) & ")"
    OutputLog = OutputLog & "OK ("  & CanReadArg.Format(CFM_SHORT) & ")"
  End If    
  
  If Memory.PrepCmd_Inprogress = 1 Then

    'Pub End detection.
    If Bit_Active = 0 AND Bit_Async = 0 Then
      If Memory.Endurance_InProgress = 1 Then
        If Bit_Error = 1 Then
          LogAdd "Prepare command completed with error : " & GetErrorInfo( CanReadArg )
        Else
          LogAdd "Prepare command completed"
        End If
      End If
      Memory.PrepCmd_Inprogress = 0
    End If
    
    Memory.PrepCmd_Error = Bit_Error
    If Bit_Error = 1 Then
      Memory.Set "CanErr",CanReadArg.Data(1)
    End If

  End If
 'Always capture PUB messages in debug report
  DebugMessage OutputDebug
End Function

'------------------------------------------------------------------
Function GetErrorInfo ( CanReadArg )
  Dim ErrorMsg
  Dim ErrCode,ErrData1,ErrData2
  

  If Memory.CANConfig.Config = 0 Then
    ErrCode = CanReadArg.Data(1)
    If CanReadArg.Length > 2 Then
      ErrData1 = CanReadArg.Data(CanReadArg.Length-2)
      ErrData2 = CanReadArg.Data(CanReadArg.Length-1)
    End If
  Else
    ErrCode = CanReadArg.Data(1)
    If CanReadArg.Length > 2 Then
      ErrData1 = CanReadArg.Data(CanReadArg.Length-2)
      ErrData2 = CanReadArg.Data(CanReadArg.Length-1)  
    End If
  End If
  'DebugMessage "Error num: " & String.Format("%2x",ErrCode)
  Select Case ErrCode
  Case $(ACK_NOK): ErrorMsg = "Command Error (NOK)"
  Case $(ACK_UNKNOWN_CMD): ErrorMsg = "Unknown Command"
  Case $(ACK_WRONG_STATE): ErrorMsg = "Wrong State"
  Case $(ACK_INVALID_PARAM): ErrorMsg = "Invalid Parameters"
  Case $(ACK_WRONG_LENGTH): ErrorMsg = "Wrong Lenght"
  Case $(ACK_TOO_MANY_PREPARES): ErrorMsg = "Too many Prepares"
  Case $(ACK_TIMEOUT_SUBSYSTEM): ErrorMsg = "Timeout Subsystem"
  Case $(ACK_NOT_IMPLEMENTED): ErrorMsg = "Not Implemented"
  'Case $(ACK_WRONG_LENGTH): ErrorMsg = "Wrong Length"      

  Case Else: ErrorMsg = "Other errors"
  End Select 
  GetErrorInfo = ErrorMsg
End Function

'------------------------------------------------------------------
Function CANSendGetMC(Cmd,SubCmd,SlotNo,Division,DataLen)

  Dim CanManager,CanConfig,CanSendArg,CanReadArg,CANData
  Dim i,Result
  Dim CANID
  
  Set CanSendArg = CreateObject("ICAN.CanSendArg")
  Set CanReadArg = CreateObject("ICAN.CanReadArg")
  Memory.Get "CANData",CANData

  If Memory.Exists("CanConfig") Then
    Memory.Get "CanConfig",CanConfig
    CANID = CanConfig.CANIDcmd
  End If
  
  'XFCU
  If CANConfig.Config = 1 Then
    With CanSendArg
        DebugMessage "In Machine GetSendMCData command"
        .CanId = CANID
        .Data(0) = Cmd + &h10
        .Data(1) = SubCmd        
        .Data(2) = SlotNo
        .Data(3) = Division
      For i = 0 to DataLen - 1
        .Data(4+i) = CANData.Data(i)
        DebugMessage "Copy Data " & i
      Next
      .Length = 4 + DataLen      
    End With
  'Standalone
  Else
    With CanSendArg
      .CanId = CANID
        DebugMessage "Standalone GetSend MC Data command"
        .Data(0) = Cmd
        .Data(1) = SubCmd
        If DataLen > 0 Then
          For i = 0 to DataLen - 1
            .Data(2+i) = CANData.Data(i)
            DebugMessage "Copy Data " & i
          Next  
        End If
        .Length = 2 + DataLen
    End With  
  End If
  
  'Process Response
  If Memory.Exists("CanManager") AND CanConfig.CANIDvalid = 1 Then    
    Memory.Get "CanManager",CanManager        
    Result = CanManager.SendCmd(CanSendArg,1000,SC_CHECK_ERROR_BYTE,CanReadArg)
    
    If  Result = SCA_NO_ERROR Then      
      DebugMessage "GetMC OK: (TX:" & CanSendArg.Format(CFM_SHORT)&")" & " (RX:" & CanReadArg.Format & ")"
      CANSendGetMC = True
      'XFCU
      If CANConfig.Config = 1 Then
           DebugMessage "Reading XFCU GetSend MC Data Reply"
          'Data(0) = CMD
          'Data(1) = ACK
          'Data(2) = Data 1
          'Data(3) = Data 2
          'Data(4) = Data 3
          'Data(5) = Data 4
          CANData.Data(2) = CanReadArg.Data(3) 
          CANData.Data(3) = CanReadArg.Data(4)
          CANData.Data(4) = CanReadArg.Data(5)
          CANData.Data(5) = CanReadArg.Data(6)
      Else
        DebugMessage "Standalone Cmd Reply"
        'No need to process data, just copy
         For i = 0 to 7
          CANData.Data(i) = CanReadArg.Data(i)
         Next
      End If
      Memory.Set "CANData",CANData
      Memory.Set "CanReadArg",CanReadArg
      DebugMessage "CANData:" & String.Format("%02X %02X %02X %02X %02X %02X %02X %02X",CanReadArg.Data(0),CanReadArg.Data(1) ,CanReadArg.Data(2) ,CanReadArg.Data(3) ,CanReadArg.Data(4) ,CanReadArg.Data(5) ,CanReadArg.Data(6) ,CanReadArg.Data(7))
    Else
      LogAdd "Command NOK: " & GetErrorInfo( CanReadArg ) & " (" & CanReadArg.Format & ")"
      CANSendGetMC = False
    End If
    CanManager.Deliver = True
  Else
    DebugMessage "CANSendGetMC Error"
    CANSendGetMC = False
  End If
  
End Function
'------------------------------------------------------------------
Function GetFloatCanData( )
  Dim Value,RawValue,CANData
  Memory.Get "CANData",CANData
  'DebugMessage String.Format("%4X,%4X,%4X,%4X",CANData(2),CANData(3),CANData(4),CANData(5))
  RawValue = Lang.MakeLong4(CANData(2),CANData(3),CANData(4),CANData(5))
  Value = Math.CastLong2Float(RawValue)
  GetFloatCanData = Value
End Function
'------------------------------------------------------------------
Function CANSendGetFeed(Cmd,SubCmd,SlotNo,DataLen)

  Dim CanManager,CanConfig,CanSendArg,CanReadArg,CANData
  Dim i,Result
  Dim CANID
  
  Set CanSendArg = CreateObject("ICAN.CanSendArg")
  Set CanReadArg = CreateObject("ICAN.CanReadArg")
  Memory.Get "CANData",CANData

  If Memory.Exists("CanConfig") Then
    Memory.Get "CanConfig",CanConfig
    CANID = CanConfig.CANIDcmd
  End If
  
  'XFCU 
  If CANConfig.Config = 1 Then
    With CanSendArg
      DebugMessage "XFCU Get Feeder Data Command"
      .CanId = CANID+2
      .Data(0) = Cmd + &h10
      .Data(1) = SubCmd        
      .Data(2) = SlotNo
      For i = 0 to DataLen - 1
        .Data(3+i) = CANData.Data(i)
        DebugMessage "Copy Data " & i
      Next
      .Length = 3 + DataLen      
    End With
  'Standalone
  Else
    With CanSendArg
      .CanId = CANID
      DebugMessage "Standalone Get Feeder Data Cmd"
      .Data(0) = Cmd
      .Data(1) = SubCmd
      If DataLen > 0 Then
        For i = 0 to DataLen - 1
          .Data(2+i) = CANData.Data(i)
          DebugMessage "Copy Data " & i
        Next  
      End If
      .Length = 2 + DataLen
    End With  
  End If
  
  'Process Response
  If Memory.Exists("CanManager") AND CanConfig.CANIDvalid = 1 Then    
    Memory.Get "CanManager",CanManager        
    Result = CanManager.SendCmd(CanSendArg,1000,SC_CHECK_ERROR_BYTE,CanReadArg)
    
    If  Result = SCA_NO_ERROR Then      
      DebugMessage "GetFeed OK: (TX:" & CanSendArg.Format(CFM_SHORT)&")" & " (RX:" & CanReadArg.Format & ")"
      CANSendGetFeed = True
      'XFCU
      If CANConfig.Config = 1 Then
        'StandAlone Prepare Commands
           DebugMessage "Reading XFCU Get Feed Reply"
          'Data(0) = CMD
          'Data(1) = ACK
          'Data(2) = Data 1
          'Data(3) = Data 2
          'Data(4) = Data 3
          'Data(5) = Data 4
          CANData.Data(2) = CanReadArg.Data(3) 
          CANData.Data(3) = CanReadArg.Data(4)
          CANData.Data(4) = CanReadArg.Data(5)
          CANData.Data(5) = CanReadArg.Data(6)
      Else
        DebugMessage "Standalone Cmd Reply"
        'No need to process data, just copy
         For i = 0 to 7
          CANData.Data(i) = CanReadArg.Data(i)
         Next
      End If
      Memory.Set "CANData",CANData
      DebugMessage "CANData:" & String.Format("%02X %02X %02X %02X %02X %02X %02X %02X",CanReadArg.Data(0),CanReadArg.Data(1) ,CanReadArg.Data(2) ,CanReadArg.Data(3) ,CanReadArg.Data(4) ,CanReadArg.Data(5) ,CanReadArg.Data(6) ,CanReadArg.Data(7))
    ElseIf Result = SCA_ERROR_MSG Then
      LogAdd "Command NOK: " & GetErrorInfo( CanReadArg ) & " (" & CanReadArg.Format & ")"
      CANSendGetFeed = False
    End If
    CanManager.Deliver = True
  Else
    DebugMessage "CANSendGetFeed Error"
    CANSendGetFeed = False
  End If
  
End Function
'------------------------------------------------------------------

Function CANSendPrepareCMD(Cmd,Context,SlotNo,Division,DataLen,PubEndTimeout)
  Dim CanManager, CanConfig
  Dim CanSendArg, CanReadArg,CANData
  Dim i,Result
  Dim CANID
  Dim WaitPub,StartTime
  
  Set CanSendArg = CreateObject("ICAN.CanSendArg")
  Set CanReadArg = CreateObject("ICAN.CanReadArg")
  Memory.Get "CANData",CANData
  If Memory.Exists("CanConfig") Then
    Memory.Get "CanConfig",CanConfig
    CANID = CanConfig.CANIDcmd
  End If
  
  'In Machine 
  If CANConfig.Config = 1 Then
    With CanSendArg
      'Prepare Commands
      'DebugMessage "In Machine Prepare Command"
      .CanId = CANID
      .Data(0) = Cmd  
      .Data(1) = Context
      .Data(2) = SlotNo
      .Data(3) = Division
      If DataLen > 0 Then 
        'DebugMessage "Prepare Command with " & DataLen & " data"
        For i = 0 to DataLen - 1
          .Data(4+i) = CANData.Data(i)
        Next
      End If
      .Length = 4 + DataLen      
    End With
  'Standalone
  Else
    With CanSendArg
      .CanId = CANID
      'DebugMessage "Standalone Prepare Cmd"
      .CanId = CANID
      .Data(0) = Cmd      
      .Data(1) = Division
      If DataLen > 0 Then
        DebugMessage "Prepare Command with " & DataLen & " data"
        For i = 0 to DataLen - 1
          .Data(2+i) = CANData.Data(i)
        Next
      End If
      .Length = 2 + DataLen
    End With
  End If
  
  If Memory.Exists("CanManager") AND CanConfig.CANIDvalid = 1 Then    
    Memory.Get "CanManager",CanManager      
    Result = CanManager.SendCmd(CanSendArg,1000,SC_CHECK_ERROR_BYTE,CanReadArg)
    'Process ACK
    If  Result = SCA_NO_ERROR Then      
      DebugMessage "Prepare OK: (TX:" & CanSendArg.Format(CFM_SHORT)&")" & " (RX:" & CanReadArg.Format & ")"
      Memory.PrepCmd_Inprogress = 1
      Memory.PrepCmd_Error = 0
      CANSendPrepareCMD = True
      'StandAlone
      If CANConfig.Config = 1 Then
      'StandAlone Prepare Commands
        'DebugMessage "XFCU Prepare Cmd Reply"
        'Move the  Error Details down to 3 /4 (same as standalone)
        'Data(0) = CMD
        'Data(1) = ACK
        'Data(2) = Division
        'Data(3) = Error Detail 1
        'Data(4) = Error Detail 2        
        CANData.Data(3) = CanReadArg.Data(5)       
        CANData.Data(4) = CanReadArg.Data(6)       
        CANData.Data(5) = CanReadArg.Data(7)
      Else
        'DebugMessage "Standalone Cmd Reply"
        'No need to process data, just copy
         For i = 0 to 7
          CANData.Data(i) = CanReadArg.Data(i)
         Next
      End If
      Memory.Set "CANData",CANData
      'DebugMessage "CANData:" & String.Format("%02X %02X %02X %02X %02X %02X %02X %02X",CanReadArg.Data(0),CanReadArg.Data(1) ,CanReadArg.Data(2) ,CanReadArg.Data(3) ,CanReadArg.Data(4) ,CanReadArg.Data(5) ,CanReadArg.Data(6) ,CanReadArg.Data(7))
    Else
      Memory.PrepCmd_Inprogress = 0      
      Memory.PrepCmd_Error = 1
      Memory.Set "CanErr",CanReadArg.Data(1)
      DebugMessage "Command NOK: " & GetErrorInfo( CanReadArg ) & " (" & CanReadArg.Format & ")"
      LogAdd "Command NOK: " & GetErrorInfo( CanReadArg ) & " (" & CanReadArg.Format & ")"
      CANSendPrepareCMD = False
      Exit Function
    End If
    CanManager.Deliver = True   
  Else
    DebugMessage "CANSendPrepareCMD Error"
    CANSendPrepareCMD = False
  End If
  
  If CANSendPrepareCMD = False Then
    LogAdd "Prepare Cmd NOK: " & GetErrorInfo( CanReadArg ) & " (" & CanReadArg.Format & ")"
  End If
  
End Function

'------------------------------------------------------------------
Function CANSendTACMD(Cmd,SubCmd,SlotNo,Division,DataLen)

  Dim CanManager,CanConfig,CanSendArg,CanReadArg,CANData
  Dim i,Result
  Dim CANID
  
  Set CanSendArg = CreateObject("ICAN.CanSendArg")
  Set CanReadArg = CreateObject("ICAN.CanReadArg")
  Memory.Get "CANData",CANData

  If Memory.Exists("CanConfig") Then
    Memory.Get "CanConfig",CanConfig
    CANID = CanConfig.CANIDcmd
  End If
  
  'XFCU
  If CANConfig.Config = 1 Then
    With CanSendArg
        DebugMessage "In Machine GetSendMCData command"
        .CanId = CANID
        .Data(0) = Cmd + &h10
        .Data(1) = SubCmd        
        .Data(2) = SlotNo
      For i = 0 to DataLen - 1
        .Data(3+i) = CANData.Data(i)
        DebugMessage "Copy Data " & i
      Next
      .Length = 3 + DataLen      
    End With
  'Standalone
  Else
    With CanSendArg
      .CanId = CANID
        DebugMessage "Standalone GetSend MC Data command"
        .Data(0) = Cmd
        .Data(1) = SubCmd
        If DataLen > 0 Then
          For i = 0 to DataLen - 1
            .Data(2+i) = CANData.Data(i)
            DebugMessage "Copy Data " & i
          Next  
        End If
        .Length = 2 + DataLen
    End With  
  End If
  
  'Process Response
  If Memory.Exists("CanManager") AND CanConfig.CANIDvalid = 1 Then    
    Memory.Get "CanManager",CanManager        
    Result = CanManager.SendCmd(CanSendArg,1000,SC_CHECK_ERROR_BYTE,CanReadArg)
    
    If  Result = SCA_NO_ERROR Then      
      DebugMessage "GetMC OK: (TX:" & CanSendArg.Format(CFM_SHORT)&")" & " (RX:" & CanReadArg.Format & ")"
      CANSendTACMD = True
      'XFCU
      If CANConfig.Config = 1 Then
          DebugMessage "Reading XFCU GetSend MC Data Reply"
          CANData.Data(2) = CanReadArg.Data(3) 
          CANData.Data(3) = CanReadArg.Data(4)
          CANData.Data(4) = CanReadArg.Data(5)
          CANData.Data(5) = CanReadArg.Data(6)
      Else
        DebugMessage "Standalone Cmd Reply"
        'No need to process data, just copy
         For i = 0 to 7
          CANData.Data(i) = CanReadArg.Data(i)
         Next
      End If
      Memory.Set "CANData",CANData
      DebugMessage "CANData:" & String.Format("%02X %02X %02X %02X %02X %02X %02X %02X",CanReadArg.Data(0),CanReadArg.Data(1) ,CanReadArg.Data(2) ,CanReadArg.Data(3) ,CanReadArg.Data(4) ,CanReadArg.Data(5) ,CanReadArg.Data(6) ,CanReadArg.Data(7))
    Else
      LogAdd "Command NOK: " & GetErrorInfo( CanReadArg ) & " (" & CanReadArg.Format & ")"
      CANSendTACMD = False
    End If
    CanManager.Deliver = True
  Else
    DebugMessage "CANSendTACMD Error"
    CANSendTACMD = False
  End If  
End Function'------------------------------------------------------------------