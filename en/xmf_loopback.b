Dim CANID,CANIDRX1,CANIDRX2,CANIDDBG
Dim CANRXMsg
Dim exit_condition

CANID     = "0x608"
CANIDDBG  = "0x60A"
CANIDRX1  = "0x408"
CANIDRX2  = "0x008"

exit_condition = False

{
  CANRXMsg = WaitMsg{"0x608,0x60A,0x500,0x502,0x503"}(250)  
  If CANRXMsg.Success && ( CANRXMsg.CanId == 0x608 ||  CANRXMsg.CanId == 0x60A )
  {
    Switch (CANRXMsg.Data[0])
    {
      Case 0x05:
      {
        If (CANRXMsg.Data[1] == 0x00)
          SendMsg{CANIDRX1}( 0x05,0x00,CANRXMsg.Data[2],0x00,0x01,0x33,0x44)
        Else If (CANRXMsg.Data[1] == 0x10)
          SendMsg{CANIDRX1}( 0x05,0x00,CANRXMsg.Data[2],0x00,0x03,0x33,0x44)
        Else If (CANRXMsg.Data[1] == 0x20)
          SendMsg{CANIDRX1}( 0x05,0x00,CANRXMsg.Data[2],0x00,0x17,0x33,0x44)
        Else If (CANRXMsg.Data[1] == 0x30)
          SendMsg{CANIDRX1}( 0x05,0x00,CANRXMsg.Data[2],0x00,0x02,0x33,0x44)

      }
      Case 0x09:
      {
        SendMsg{CANIDRX1}( 0x09,0x00)
      }
      Case 0x41:
      {
        SendMsg{CANIDRX1}( CANRXMsg.Data[0],0x00,0x01)
        delay 200
        SendMsg{CANIDRX2}( 0x90,0x00,0x01)
        delay 2000
        SendMsg{CANIDRX2}( 0x00,0x00,0x01)
      }
      Case 0x42:
      {
        SendMsg{CANIDRX1}( CANRXMsg.Data[0],0x00,0x01)
        delay 200
        SendMsg{CANIDRX2}( 0x90,0x00,0x01)
        delay 2000
        SendMsg{CANIDRX2}( 0x00,0x00,0x01)
      }      
      Case 0x43:
      {
        SendMsg{CANIDRX1}( CANRXMsg.Data[0],0x00,0x01)
        delay 200
        SendMsg{CANIDRX2}( 0x90,0x00,0x01)
        delay 2000
        SendMsg{CANIDRX2}( 0x00,0x00,0x01)
      }     
      Case 0x46:
      {
        SendMsg{CANIDRX1}( CANRXMsg.Data[0],0x00,0x01)
        delay 200
        SendMsg{CANIDRX2}( 0x90,0x00,0x01)
        delay 2000
        SendMsg{CANIDRX2}( 0x00,0x00,0x01)
      }     
      Case 0x47:
      {
        SendMsg{CANIDRX1}( CANRXMsg.Data[0],0x00,0x01)
        delay 200
        SendMsg{CANIDRX2}( 0x90,0x00,0x01)
        delay 2000
        SendMsg{CANIDRX2}( 0x00,0x00,0x01)
      }
      'Read MC
      Case 0x6A:
      {
        'Switch (CANRXMsg.Data[1])
        '{
          SendMsg{CANIDRX1}( CANRXMsg.Data[0],0x00,0x17,0xB7,0xD1,0x38)
        '}
      }
      'ParamGet
      Case 0x81:
      {
        Switch (CANRXMsg.Data[1])
        {
          Case 0x70:
          {
            If (CANRXMsg.Data[2] == 0x00 || CANRXMsg.Data[2] == 0x02)
              SendMsg{CANIDRX1}(CANRXMsg.Data[0],0x00)
            Else If (CANRXMsg.Data[2] == 0x01)
              SendMsg{CANIDRX1}(CANRXMsg.Data[0],0x10,0x00,0x01,0x02,0x03)
          }
        }
      }
   
    }
  }
    delay 50
}
Until exit_condition == True
