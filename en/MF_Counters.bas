
Const EEPROMParamFile = "EEPROM_params.xml" 
Const CM1_TARGET = 5
Const CM2_TARGET = 6
Const CM3_TARGET = 7
Const MB_TARGET  = 8

'------------------------------------------------------------------
Function OnClick_btn_ReadCMEEPROM( Reason )
  Dim EEPROMData,i

  If Memory.Get("EEPROMData_CM",EEPROMData) = True Then
    If GetEEPROMML(0,CM1_TARGET,EEPROMData) = True Then
      UpdateCMGrid
    Else
      LogAdd "Check if Contact Module is installed properly"
    End If
  End If

End Function

'------------------------------------------------------------------
Function OnClick_btn_ReadMBEEPROM( Reason )
Dim EEPROMData
Set EEPROMData = Memory.EEPROMData_MB
      
  If GetEEPROMML(0,MB_TARGET,EEPROMData) = True Then
    UpdateMBGrid
  End If

End Function
'------------------------------------------------------------------
Function OnClick_btn_ReadCLEEPROM( Reason )
Dim EEPROMData
Set EEPROMData = Memory.EEPROMData_Calib
  If GetEEPROMML(0,CM3_TARGET,EEPROMData) = True Then
    UpdateCalibGrid
  Else
    LogAdd "Check if Calibration Module is connected"
  End If

End Function
'------------------------------------------------------------------
Function OnClick_btn_SaveCMEEPROM(Reason)
  Dim Filename
  
  If System.FileDialog( False, _
     OFN_HIDEREADONLY or OFN_CREATEPROMPT or OFN_OVERWRITEPROMPT, _
     "CSV files(*.csv)", _
     Filename ) Then
    DebugMessage Filename
    WriteFile "CMEEPROMGrid",Filename
    System.MessageBox filename & " saved.","File Saved",MB_OK
  End if
End Function
'------------------------------------------------------------------
Function OnClick_btn_SaveMBEEPROM(Reason)
  Dim Filename
  
  If System.FileDialog( False, _
     OFN_HIDEREADONLY or OFN_CREATEPROMPT or OFN_OVERWRITEPROMPT, _
     "CSV files(*.csv)", _
     Filename ) Then
    DebugMessage Filename
    WriteFile "MBEEPROMGrid",Filename
    System.MessageBox filename & " saved.","File Saved",MB_OK
  End if
End Function
'------------------------------------------------------------------
Function OnClick_btn_SaveCLEEPROM(Reason)
  Dim Filename
  
  If System.FileDialog( False, _
     OFN_HIDEREADONLY or OFN_CREATEPROMPT or OFN_OVERWRITEPROMPT, _
     "CSV files(*.csv)", _
     Filename ) Then
    DebugMessage "Save calibration data to:"& Filename
    WriteFile "CalibEEPROMGrid",Filename
    System.MessageBox filename & " saved.","File Saved",MB_OK
  End if
End Function
'------------------------------------------------------------------
Function InitEEPROMGrid()
  Dim xmlOk, XMLfilepath, ackNode, i 
  Dim Address, Length, Name, Format, Value, SizeNode
  Dim EEPROMSize
  Dim EEPROMData_CM,EEPROMData_MB
  Dim CM_Grid
  Dim MB_Grid
  Dim Calib_Grid
  Dim MsgId
  
  Set CM_Grid = Visual.Script("CMEEPROMGrid")
  Set MB_Grid = Visual.Script("MBEEPROMGrid")
  XMLfilepath = System.Environment.Path & EEPROMParamFile 
  
  Set EEPROMData_CM = CreateObject("Math.ByteArray")
  Set EEPROMData_MB = CreateObject("Math.ByteArray")
  'Checks if xml file is present
  If File.FileExists(XMLfilepath) = False Then
    DebugMessage "XML file NOK : " & XMLfilepath
    System.MessageBox "~cf3 File " & Chr(34) & EEPROMParamFile & Chr(34) & " does not exist !", "File error", MB_ICONERROR 
  Else
    DebugMessage "XML file OK : " & XMLfilepath
    xmlOk = True
    Set ackNode = CreateObject("XMLCW.XmlParser").Build( XMLfilepath )
    Memory.Set "EEPROMXMLNode",ackNode
    xmlOk = Lang.IsObject(ackNode)
    If xmlOk = True Then
      'Pupulate CM Tab
      EEPROMSize = ackNode("ContactModule").Child("memory").Attribute.Attribute("size")
      'Init memory to hold EEPROMData
      For i = 0 to EEPROMSize-1
        EEPROMData_CM.Add(0)
      Next
      DebugMessage "CM size:" & EEPROMSize & " EEPROMData_CM Size:" & EEPROMData_CM.Size
      Memory.Set "EEPROMData_CM",EEPROMData_CM
      Set ackNode = ackNode("ContactModule").SelectNodes("param")
      xmlOk = Lang.IsObject(ackNode)
    End If
    If xmlOk = True Then
      'Search for all parameters and obtain the value from contact module.
      'DebugMessage "Number of parameters: " & ackNode.size
      For i = 0 to ackNode.Size-1        
        Address =  ackNode(i).Attribute.Attribute("address") 
        Length = ackNode(i).Attribute.Attribute("length") 
        Format = ackNode(i).Attribute.Attribute("format") 
        Name = ackNode.ChildContent(i)
        'DebugMessage "Param" & i & " " & Name & " "& Address & " "& Length & " "& Format & " " & Value
        CM_Grid.addrow i,Name & "," & String.Format("0x%04X",Address) & ",",i
        'Value = GetCMParam( 1, Address, Length, Format)
      Next      
      Memory.Set "CMFORMATNODE",ackNode
    End If
    ' Populate MB tab
    Set ackNode = Memory.EEPROMXMLNode
    xmlOk = Lang.IsObject(ackNode)
    If xmlOk = True Then
      EEPROMSize = ackNode("MeasurementBoard").Child("memory").Attribute.Attribute("size")
      'Init memory to hold EEPROMData
      For i = 0 to EEPROMSize-1
        EEPROMData_MB.Add(0)
      Next
      DebugMessage "MB size:" & EEPROMSize & " EEPROMData_MB Size:" & EEPROMData_MB.Size
      Memory.Set "EEPROMData_MB",EEPROMData_MB
      Set ackNode = ackNode("MeasurementBoard").SelectNodes("param")
      xmlOk = Lang.IsObject(ackNode)
    End If
    If xmlOk = True Then
      'Search for all parameters and obtain the value from contact module.
      'DebugMessage "Number of parameters: " & ackNode.size
      For i = 0 to ackNode.Size-1        
        Address =  ackNode(i).Attribute.Attribute("address") 
        Length = ackNode(i).Attribute.Attribute("length") 
        Format = ackNode(i).Attribute.Attribute("format") 
        Name = ackNode.ChildContent(i)
        MB_Grid.addrow i,Name & "," & String.Format("0x%04X",Address) & ",",i
        'Value = GetCMParam( 1, Address, Length, Format)
        'DebugMessage "Param" & i & " " & Name & " "& Address & " "& Length & " "& Format & " " & Value
      Next      
      Memory.Set "MBFORMATNODE",ackNode
    End If
    ' Populate Calib tab
    InitEEPROMGrid_Calib
  End If  
End Function

'------------------------------------------------------------------
Function InitEEPROMGrid_Calib()
  Dim xmlOk, XMLfilepath, ackNode, i 
  Dim Address, Length, Name, Format,RowType
  Dim EEPROMSize
  Dim EEPROMData_Calib
  Dim Calib_Grid
  
  Set Calib_Grid = Visual.Script("CalibEEPROMGrid")
  Set EEPROMData_Calib = CreateObject("Math.ByteArray")
  Set ackNode = Memory.EEPROMXMLNode
  xmlOk = Lang.IsObject(ackNode)
  If xmlOk = True Then
    EEPROMSize = ackNode("CalibrationBoard").Child("memory").Attribute.Attribute("size")
    'Init memory to hold EEPROMData
    For i = 0 to EEPROMSize-1
      EEPROMData_Calib.Add(0)
    Next
    DebugMessage "Calib size:" & EEPROMSize & " EEPROMData_Calib Size:" & EEPROMData_Calib.Size
    Memory.Set "EEPROMData_Calib",EEPROMData_Calib
    Set ackNode = ackNode("CalibrationBoard").SelectNodes("param")
    xmlOk = Lang.IsObject(ackNode)
  End If
  If xmlOk = True Then
    'Search for all parameters
    'DebugMessage "Number of parameters: " & ackNode.size
    For i = 0 to ackNode.Size-1        
      RowType =  ackNode(i).Attribute.Attribute("paramtype") 
      Address =  ackNode(i).Attribute.Attribute("address") 
      Length = ackNode(i).Attribute.Attribute("length") 
      Format = ackNode(i).Attribute.Attribute("format") 
      Name = ackNode.ChildContent(i)
      Calib_Grid.addrow i,RowType & "," & Name & "," & String.Format("0x%04X",Address) & ",",i
    Next
    Memory.Set "CALIBFORMATNODE",ackNode          
    Calib_Grid.collapseAllGroups
    Calib_Grid.expandGroup "Param"
  End If
End Function
'------------------------------------------------------------------
Function UpdateCalibGrid()
  Dim Node,EEPROMData_Calib,xmlOk
  Dim Address, Length, Format, Name
  Dim i,y
  Dim Data,DataError
  Memory.Get "EEPROMData_Calib",EEPROMData_Calib
  Set Node = Memory.CALIBFORMATNODE
  xmlOk = Lang.IsObject(Node)
  If xmlOk Then
    For i = 0 to Node.Size-1
      DataError = False
      Format = Node(i).Attribute.Attribute("format")
      Address = Node(i).Attribute.Attribute("address")
      Length = Node(i).Attribute.Attribute("length")
      Name = Node.ChildContent(i)
      'DebugMessage "Node: " & i & " " & String.Format("%04X",Address) & " " & Length
      'Format data based on address and data format, from EEPROM data array read using GetEEPROMML
      Select case Format
        case "str":
          'Convert to string. 
          If GetString(EEPROMData_Calib,Address,Length,Data) = False Then
            'Error with string
            DataError = True
          End If
        case "x":
          Data = EEPROMData_Calib.Char(Address)        
        case "s":
          Data = EEPROMData_Calib.Word(Address)
        case "f":
          Data = String.Format("%G",EEPROMData_Calib.Float(Address))
        case "s1":
          Data = EEPROMData_Calib.Short(Address)
        case "s2":
          Data = EEPROMData_Calib.Long(Address)
        case "s3":
          Data = EEPROMData_Calib.Long(Address)
        case "x1":
          Data = String.Format("0x%08X",Lang.MakeLong4(EEPROMData_Calib.Data(Address),EEPROMData_Calib.Data(Address+1),EEPROMData_Calib.Data(Address+2),EEPROMData_Calib.Data(Address+3)))
          'DebugMessage String.Format("%02X,%02X,%02X,%02X",EEPROMData_Calib.Data(Address),EEPROMData_Calib.Data(Address+1),EEPROMData_Calib.Data(Address+2),EEPROMData_Calib.Data(Address+3))
          'Data = String.Format("%u",Lang.MakeLong4(EEPROMData_Calib.Data(Address),EEPROMData_Calib.Data(Address+1),EEPROMData_Calib.Data(Address+2),EEPROMData_Calib.Data(Address+3)))
        case else:
      End Select
      If Name = "Type" Then
        Select case Data 
        case 1:
          Data = "Res"
        case 2:
          Data = "Cap"
        case 3:
          Data = "Inductor"
        case 4:
          Data = "Diode"
        case 5:
          Data = "Polar Cap"
        case Else:
          Data = "Unknown"
          DataError = True        
        End Select
      End If
      Visual.Script("CalibEEPROMGrid").setVal i,Data
      If DataError = True Then
        Visual.Script("CalibEEPROMGrid").setCellRed(i)
      Else
        Visual.Script("CalibEEPROMGrid").setCellBlack(i)        
      End If
      'DebugMessage "Param" & i & " "& Address & " "& Length & " "& Format & "Data: " & Data
    Next
    'Get Component data
  End If
End Function
'------------------------------------------------------------------
Function UpdateMBGrid()
Dim Node,EEPROMData_MB
Dim xmlOk
Dim Format,Length,Address
Dim i,y
Dim Data
Dim DataError
  Memory.Get "EEPROMData_MB",EEPROMData_MB
  Set Node = Memory.MBFORMATNODE
  xmlOk = Lang.IsObject(Node)
  If xmlOk Then
    For i = 0 to Node.Size-1
      DataError = False
      Format = Node(i).Attribute.Attribute("format")
      Address = Node(i).Attribute.Attribute("address")
      Length = Node(i).Attribute.Attribute("length")       
      'Format data based on address and data format, from EEPROM data array read using GetEEPROMML
      Select case Format
        case "str":
          'Convert to string. 
          If GetString(EEPROMData_MB,Address,Length,Data) = False Then
            'Error with string
            DataError = True
          End If
        case "x":
          Data = EEPROMData_MB.Char(Address)        
        case "s":
          Data = EEPROMData_MB.Word(Address)
        case "f":
          Data = String.Format("%G",EEPROMData_MB.Float(Address))
        case "s1":
          Data = EEPROMData_MB.Short(Address)
        case "s2":
          Data = EEPROMData_MB.Long(Address)
        case "s3":
          Data = EEPROMData_MB.Long(Address)
          'DebugMessage String.Format("%02X,%02X,%02X,%02X",EEPROMData_MB.Data(Address),EEPROMData_MB.Data(Address+1),EEPROMData_MB.Data(Address+2),EEPROMData_MB.Data(Address+3))
          'Data = String.Format("%u",Lang.MakeLong4(EEPROMData_MB.Data(Address),EEPROMData_MB.Data(Address+1),EEPROMData_MB.Data(Address+2),EEPROMData_MB.Data(Address+3)))
        case else:
      End Select
      Visual.Script("MBEEPROMGrid").setVal i,Data
      If DataError = True Then
        Visual.Script("MBEEPROMGrid").setCellRed(i)
      Else
        Visual.Script("MBEEPROMGrid").setCellBlack(i)        
      End If

      'DebugMessage "Param" & i & " "& Address & " "& Length & " "& Format & "Data: " & Data
    Next
  End If

End Function
'------------------------------------------------------------------
Function UpdateCMGrid()
Dim Node,EEPROMData_CM
Dim xmlOk
Dim Format,Length,Address
Dim i,y
Dim Data
Dim DataError
  Memory.Get "EEPROMData_CM",EEPROMData_CM
  Set Node = Memory.CMFORMATNODE
  xmlOk = Lang.IsObject(Node)
  If xmlOk Then
    'loop through all the parameters in the CMFORMATNODE node and display them
    For i = 0 to Node.Size-1
      DataError = False
      Format = Node(i).Attribute.Attribute("format")
      Address = Node(i).Attribute.Attribute("address")
      Length = Node(i).Attribute.Attribute("length")       
      'Format data based on address and data format, from EEPROM data array read using GetEEPROMML
      Select case Format
        case "str":
          'TODO: Investigate better way to do this
          If GetString(EEPROMData_CM,Address,Length,Data) = False Then
            'Error with string
            DataError = True
          End If
        case "x":
          Data = EEPROMData_CM.Char(Address)        
        case "s":
          Data = EEPROMData_CM.Word(Address)
        case "f":
          Data = String.Format("%G",EEPROMData_CM.Float(Address))
        case "s1":
          Data = EEPROMData_CM.Short(Address)
        case "s2":
          Data = EEPROMData_CM.Long(Address)
        case "s3":
          Data = String.Format("%u",Lang.MakeLong4(EEPROMData_CM.Data(Address),EEPROMData_CM.Data(Address+1),EEPROMData_CM.Data(Address+2),EEPROMData_CM.Data(Address+3)))
        case else:
      End Select      
      Visual.Script("CMEEPROMGrid").setVal i,Data
      If DataError = True Then
        Visual.Script("CMEEPROMGrid").setCellRed(i)
      Else
        Visual.Script("CMEEPROMGrid").setCellBlack(i)
      End If
      'DebugMessage "Param" & i & " "& Address & " "& Length & " "& Format & "Data: " & Data
    Next
  End If

End Function
'------------------------------------------------------------------
Function GetCMParam ( CM_num,address,length,format )
  Dim Value
  Dim Buffer

  Value = 0
  GetCMParam = Value
End Function
'------------------------------------------------------------------
'read EEPROM data, based on the size and commit it into memory
Function GetEEPROMML(address,target,byref EEPROMArray)
  Dim i,exitloop,Timeout,CANData
  Dim scitx,scirx, bytesleft
  Dim DebugLine
  Dim ByteCounter
  Dim LineNumber  
  Dim ProcessError
 
 ProcessError = 0
  Memory.Get "CANData",CANData
  bytesleft = EEPROMArray.size
  DebugMessage "Bytes to Read from EEPROM: " & bytesleft
  ByteCounter = 0
  exitloop = 0
  'Arbitary timeout to limit EEPROM lines read.
  Timeout = bytesleft/6 + 10
 'Get SCI TX
  CANData.Data(0) = target
  CANData.Data(1) = Lang.GetByte(address,0)
  CANData.Data(2) = Lang.GetByte(address,1)
  CANData.Data(3) = Lang.GetByte(address,2)
  CANData.Data(4) = Lang.GetByte(address,3)
  Memory.Set "CANData",CANData
  LineNumber = 0
  DebugMessage "ML:Start"
  If CANSendGetEEPROM($(CMD_GET_DATA),$(PARAM_GET_EEPROM_START), Memory.SLOT_NO,1,5) = True Then      
    Do
      LineNumber = LineNumber + 1
      'DebugMessage "ML:Line " & LineNumber & " bytes: " & bytesleft
      DebugLine = ""
      If CANSendGetEEPROM($(CMD_GET_DATA),$(PARAM_GET_EEPROM_LINE), Memory.SLOT_NO,1,0) = True Then
        Memory.Get "CANData",CANData
        For i = 2 To Memory.CANDataLen-1     
          If bytesleft > 0 Then
            DebugLine = DebugLine & String.Format("%02X ",(CANData.Data(i)))
            EEPROMArray.Data(ByteCounter) = CANData.Data(i)
            bytesleft = bytesleft - 1
            ByteCounter = ByteCounter + 1
            exitloop = 0
          End If
        Next
        If CANData.Data(1) = $(ACK_NO_MORE_DATA) Then
          If bytesleft > 0 Then
            DebugMessage "ML:End (no more data; "& ByteCounter & " bytes read, "& bytesleft & " bytes left)"
            LogAdd "Unexpected end of data! Read: " & ByteCounter & "bytes, left: " & bytesleft & "bytes"
          Else
            DebugMessage "ML:End (no more data; "& ByteCounter & " bytes read)"
          End If         
          exitloop = 1
        End If
      Else
        'PARAM_GET_EEPROM_LINE Error
        exitloop = 1
        ProcessError = 1
      End If
      'DebugMessage "ACK:" & CANData(1) & " ML:" & DebugLine & " Bytes left:" & bytesleft & " Length:" & Memory.CANDataLen
      Timeout = Timeout - 1
      If Timeout = 0 Then
        exitloop = 1
        DebugMessage "ML:TimeOut"
      End If
    Loop Until exitloop = 1
  Else
    'PARAM_GET_EEPROM_START Error
    ProcessError = 1
  End If  
  
  If ProcessError = 1 Then
     LogAdd "Error Reading EEPROM"
     GetEEPROMML = False
  Else
     GetEEPROMML = True
  End If
     

End Function

Function WriteFile (scptobj,filename)
  Dim workFile, xmlString
  xmlString = Visual.Script(scptobj).serializeToCSV(True)
  xmlString = String.Replace(xmlString, "'", chr(34)&"")
  Set workFile = File.Open( filename, "wt")
  workFile.Write xmlString
  Set workFile = nothing
End Function

'------------------------------------------------------------------
Function GetString(ByRef array,index,length,TargetString)
  Dim TmpString,TmpWord,y
  Dim ValidString
  Dim DataByte
  ValidString = True
  TmpString = ""
  For y = 0 to length-1
    DataByte = array.Data(index+y)
    If DataByte < 32 OR DataByte > 126 Then
      TmpWord = Chr(63)
      ValidString = False
    Else
      TmpWord = Chr(DataByte)
    End If
    TmpString = TmpString & TmpWord
  Next
  GetString = ValidString
  TargetString = TmpString
End Function
'------------------------------------------------------------------
Function GetChar(input)
  'Check if character is printable
  If input < 32 OR input > 126 Then
    GetChar = Chr(63)
  Else
    GetChar = Chr(input)
  End If
  
End Function
