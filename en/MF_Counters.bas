
Const MB_TARGET  = 4
Const CM1_TARGET = 5
Const CM2_TARGET = 6
Const CM3_TARGET = 7


Function OnClick_btn_ReadCounters( Reason )

ReadCounters

End Function

Sub ReadCounters()
  Dim xmlOk, XMLfilepath, ackNode, i 
  Dim Address, Length, Name, Format, Value, SizeNode
  Dim EEPROMSize

  XMLfilepath = System.Environment.Path & "parameters.xml" 
  
  'Checks if xml file is present
  If File.FileExists(XMLfilepath) = False Then
    DebugMessage "XML file NOK : " & XMLfilepath
    System.MessageBox "~cf3 File " & Chr(34) & "parameters.xml" & Chr(34) & " does not exist !", "File error", MB_ICONERROR 
  Else
    DebugMessage "XML file OK : " & XMLfilepath
    xmlOk = True
    Set ackNode = CreateObject("XMLCW.XmlParser").Build( XMLfilepath )
    xmlOk = Lang.IsObject(ackNode)
    If xmlOk = True Then      
      DebugMessage "xml Size: " & ackNode.size
      EEPROMSize = ackNode("ContactModule").Child("memory").Attribute.Attribute("size")
      DebugMessage "CM memory size " & EEPROMSize
      GetEEPROMML 0,CM1_TARGET,EEPROMSize
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
        'Value = GetCMParam( 1, Address, Length, Format)
        DebugMessage "Param" & i & " " & Name & " "& Address & " "& Length & " "& Format & " " & Value
      Next      
    End If
  End If  
End Sub

Function GetCMParam ( CM_num,address,length,format )
  Dim Value
  Dim Buffer

  Value = 0
  GetCMParam = Value
End Function

'read EEPROM data, based on the size and commit it into memory
Function GetEEPROMML(address,target,size)
  dim EEPROMArray,i,exitloop,Timeout
  dim scitx,scirx, bytesleft
  Set EEPROMArray = CreateObject( "MATH.Array" )
  Dim DebugLine
  bytesleft = size
  exitloop = 0
  Timeout = 50
 'Get SCI TX
  Memory.CANData(0) = target
  Memory.CANData(1) = Lang.GetByte(address,0)
  Memory.CANData(2) = Lang.GetByte(address,1)
  Memory.CANData(3) = Lang.GetByte(address,2)
  Memory.CANData(4) = Lang.GetByte(address,3)
  
  DebugMessage "ML:Start"
  If CANSendGetEEPROM($(CMD_GET_DATA),$(PARAM_GET_EEPROM_START), Memory.SLOT_NO,1,5) = True Then      
    Do
      DebugMessage "ML:Line"
      DebugLine = ""
      CANSendGetEEPROM $(CMD_GET_DATA),$(PARAM_GET_EEPROM_LINE), Memory.SLOT_NO,1,0      
      For i = 2 To Memory.CANDataLen-1
        DebugLine = DebugLine & Memory.CANData.Data(i)
        EEPROMArray.Add(Memory.CANData.Data(i))
      Next
      DebugMessage "ACK:" & Memory.CANData(1) & " ML:" & DebugLine
      If Memory.CANData(1) = $(ACK_NO_MORE_DATA) Then
        exitloop = 1
      End If
      bytesleft = bytesleft - (Memory.CANDataLen - 2)
      If bytesleft <= 0 Then
        exitloop = 1
      End If
      DebugMessage "ACK:" & Memory.CANData(1) & " ML:" & DebugLine & " Bytes left:" & bytesleft & " Length:" & Memory.CANDataLen
      Timeout = Timeout - 1
      If Timeout = 0 Then
        exitloop = 1
        DebugMessage "ML:TimeOut"
      End If
    Loop Until exitloop = 1
  Else
    'DebugMessage "Error"
  End If
  Memory.Set "EEPROMArray",EEPROMArray

End Function