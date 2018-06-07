
Const MB_TARGET  = 4
Const CM1_TARGET = 5
Const CM2_TARGET = 6
Const CM3_TARGET = 7

'------------------------------------------------------------------
Function OnClick_btn_ReadCMEEPROM( Reason )
  Dim EEPROMData,i

  If Memory.Get("EEPROMData_CM",EEPROMData) = True Then
    GetEEPROMML 0,CM1_TARGET,EEPROMData
    UpdateCMGrid
  End If

End Function

'------------------------------------------------------------------
Function OnClick_btn_ReadMBEEPROM( Reason )
Dim EEPROMData
Set EEPROMData = Memory.EEPROMData_MB
GetEEPROMML 0,MB_TARGET,EEPROMData
UpdateMBGrid

End Function
'------------------------------------------------------------------
Function InitEEPROMGrid()
  Dim xmlOk, XMLfilepath, ackNode, i 
  Dim Address, Length, Name, Format, Value, SizeNode
  Dim EEPROMSize
  Dim EEPROMData_CM,EEPROMData_MB
  Dim CM_Grid
  Dim MB_Grid
  Dim MsgId
  
  Set CM_Grid = Visual.Script("CMEEPROMGrid")
  Set MB_Grid = Visual.Script("MBEEPROMGrid")
  XMLfilepath = System.Environment.Path & "parameters.xml" 
  
  Set EEPROMData_CM = CreateObject("Math.ByteArray")
  Set EEPROMData_MB = CreateObject("Math.ByteArray")
  'Checks if xml file is present
  If File.FileExists(XMLfilepath) = False Then
    DebugMessage "XML file NOK : " & XMLfilepath
    System.MessageBox "~cf3 File " & Chr(34) & "parameters.xml" & Chr(34) & " does not exist !", "File error", MB_ICONERROR 
  Else
    DebugMessage "XML file OK : " & XMLfilepath
    xmlOk = True
    Set ackNode = CreateObject("XMLCW.XmlParser").Build( XMLfilepath )
    Memory.Set "EEPROMXMLNode",ackNode
    xmlOk = Lang.IsObject(ackNode)
    If xmlOk = True Then
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
      DebugMessage "param Size: " & ackNode.size
      For i = 0 to ackNode.Size-1        
        Address =  ackNode(i).Attribute.Attribute("address") 
        Length = ackNode(i).Attribute.Attribute("length") 
        Format = ackNode(i).Attribute.Attribute("format") 
        Name = ackNode.ChildContent(i)
        DebugMessage "Param" & i & " " & Name & " "& Address & " "& Length & " "& Format & " " & Value
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
      DebugMessage "param Size: " & ackNode.size
      For i = 0 to ackNode.Size-1        
        Address =  ackNode(i).Attribute.Attribute("address") 
        Length = ackNode(i).Attribute.Attribute("length") 
        Format = ackNode(i).Attribute.Attribute("format") 
        Name = ackNode.ChildContent(i)
        MB_Grid.addrow i,Name & "," & String.Format("0x%04X",Address) & ",",i
        'Value = GetCMParam( 1, Address, Length, Format)
        DebugMessage "Param" & i & " " & Name & " "& Address & " "& Length & " "& Format & " " & Value
      Next      
      Memory.Set "MBFORMATNODE",ackNode
    End If
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
Dim TmpString,TmpWord
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
          'DebugMessage String.Format("%02X,%02X,%02X,%02X",EEPROMData_MB.Data(Address),EEPROMData_MB.Data(Address+1),EEPROMData_MB.Data(Address+2),EEPROMData_MB.Data(Address+3))
          Data = String.Format("%u",Lang.MakeLong4(EEPROMData_MB.Data(Address),EEPROMData_MB.Data(Address+1),EEPROMData_MB.Data(Address+2),EEPROMData_MB.Data(Address+3)))
        case else:
      End Select
      Visual.Script("MBEEPROMGrid").setVal i,Data
      If DataError = True Then
        Visual.Script("MBEEPROMGrid").setCellRed(i)
      End If

      DebugMessage "Param" & i & " "& Address & " "& Length & " "& Format & "Data: " & Data
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
Dim TmpString,TmpWord
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
      End If
      DebugMessage "Param" & i & " "& Address & " "& Length & " "& Format & "Data: " & Data
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
 
  Memory.Get "CANData",CANData
  bytesleft = EEPROMArray.size
  DebugMessage "Bytes to Read from EEPROM: " & bytesleft
  ByteCounter = 0
  exitloop = 0
  'Arbitary timeout to limit EEPROM lines read.
  Timeout = ( bytesleft / 6 ) + 10
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
          Else
            DebugMessage "ML: End"
            exitloop = 1
            Exit For
          End If
        Next
        If CANData(1) = $(ACK_NO_MORE_DATA) Then
          exitloop = 1
        End If
      End If
      DebugMessage "ACK:" & CANData(1) & " ML:" & DebugLine & " Bytes left:" & bytesleft & " Length:" & Memory.CANDataLen
      Timeout = Timeout - 1
      If Timeout = 0 Then
        exitloop = 1
        DebugMessage "ML:TimeOut"
      End If
    Loop Until exitloop = 1
  Else
  End If

End Function

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

Function GetChar(input)
  'Check if character is printable
  If input < 32 OR input > 126 Then
    GetChar = Chr(63)
  Else
    GetChar = Chr(input)
  End If
  
End Function
