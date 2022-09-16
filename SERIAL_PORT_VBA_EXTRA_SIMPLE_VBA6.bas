Attribute VB_Name = "SERIAL_PORT_VBA_SIMPLE"
'
' https://github.com/Serialcomms/Serial-Ports-in-VBA-Extra-Simple-2022/
' https://github.com/Serialcomms/Serial-Ports-in-VBA-Extra-Simple-2022/blob/main/SERIAL_PORT_EXTRA_SIMPLE_VBA6.bas
'
Option Explicit
'
'--------------------------------------------------------------------------
 Private Const COM_PORT_NUMBER As Long = 1  ' < Change COM_PORT_NUMBER here
' -------------------------------------------------------------------------
'
Private Const LONG_0 As Long = 0
Private Const HANDLE_INVALID As Long = -1
Private Const READ_BUFFER_LENGTH As Long = 1024

Private Type COM_PORT_TIMEOUTS

             Read_Interval_Timeout          As Long
             Read_Total_Timeout_Multiplier  As Long
             Read_Total_Timeout_Constant    As Long
             Write_Total_Timeout_Multiplier As Long
             Write_Total_Timeout_Constant   As Long
End Type

Private Type COM_PORT_PROFILE

             Handle     As Long
             Started    As Boolean
             Timeouts   As COM_PORT_TIMEOUTS
End Type

Private COM_PORT As COM_PORT_PROFILE

Private Declare Function Com_Port_Close Lib "Kernel32.dll" Alias "CloseHandle" (ByVal Port_Handle As Long) As Boolean

Private Declare Function Set_Com_Timers Lib "Kernel32.dll" Alias "SetCommTimeouts" (ByVal Port_Handle As Long, ByRef Timeouts As COM_PORT_TIMEOUTS) As Boolean

Private Declare Function Com_Port_Create Lib "Kernel32.dll" Alias "CreateFileA" _
(ByVal Port_Name As String, ByVal PORT_ACCESS As Long, ByVal SHARE_MODE As Long, ByVal SECURITY_ATTRIBUTES_NULL As Any, _
 ByVal CREATE_DISPOSITION As Long, ByVal FLAGS_AND_ATTRIBUTES As Long, Optional TEMPLATE_FILE_HANDLE_NULL) As Long
 
 Private Declare Function Synchronous_Read Lib "Kernel32.dll" Alias "ReadFile" _
(ByVal Port_Handle As Long, ByVal Buffer_Data As String, ByVal Bytes_Requested As Long, ByRef Bytes_Processed As Long, Optional Overlapped_Null) As Boolean

Private Declare Function Synchronous_Write Lib "Kernel32.dll" Alias "WriteFile" _
(ByVal Port_Handle As Long, ByVal Buffer_Data As String, ByVal Bytes_Requested As Long, ByRef Bytes_Processed As Long, Optional Overlapped_Null) As Boolean
'

Public Function START_COM_PORT() As Boolean

Dim Start_Result As Boolean

With COM_PORT

 If Not .Started Then

    If OPEN_COM_PORT Then
    
        If SET_PORT_TIMERS Then
            
                    .Started = True
            
                Start_Result = True
            
        Else
                STOP_COM_PORT
        End If
                       
    End If

 End If

End With

DoEvents

START_COM_PORT = Start_Result

End Function

Private Function OPEN_COM_PORT() As Boolean

Dim Device_Path As String
Dim Open_Result As Boolean

Const OPEN_EXISTING As Long = 3
Const OPEN_EXCLUSIVE As Long = LONG_0
Const SYNCHRONOUS_MODE As Long = LONG_0

Const GENERIC_RW As Long = &HC0000000
Const DEVICE_PREFIX As String = "\\.\COM"
        
Device_Path = DEVICE_PREFIX & CStr(COM_PORT_NUMBER)

COM_PORT.Handle = Com_Port_Create(Device_Path, GENERIC_RW, OPEN_EXCLUSIVE, LONG_0, OPEN_EXISTING, SYNCHRONOUS_MODE)

Open_Result = Not (COM_PORT.Handle = HANDLE_INVALID)

OPEN_COM_PORT = Open_Result

End Function

Public Function STOP_COM_PORT() As Boolean

Dim Stop_Result As Boolean

With COM_PORT

 If .Handle > LONG_0 Then
    
    .Started = False
    
     Stop_Result = Com_Port_Close(.Handle)
    
    .Handle = IIf(Stop_Result, LONG_0, HANDLE_INVALID)
                      
 End If

End With

DoEvents

STOP_COM_PORT = Stop_Result

End Function

Public Function READ_COM_PORT(Optional Read_Length As Long) As String

Dim Bytes_Read As Long
Dim Read_String As String
Dim Read_Buffer As String * READ_BUFFER_LENGTH  ' Important - read buffer must be fixed length.

With COM_PORT
  
 If .Started Then
    
     If Read_Length = LONG_0 Or Read_Length > READ_BUFFER_LENGTH Then Read_Length = READ_BUFFER_LENGTH
    
     Synchronous_Read .Handle, Read_Buffer, Read_Length, Bytes_Read
                   
     If Bytes_Read > LONG_0 Then Read_String = Left$(Read_Buffer, Bytes_Read)
                       
 End If
  
End With

DoEvents

READ_COM_PORT = Read_String

End Function

Public Function SEND_COM_PORT(ByVal Send_String As String) As Boolean

' Important - maximum characters transmitted may be limited by write constant timer

Dim Bytes_Sent As Long
Dim Send_Result As Boolean
Dim Send_String_Length As Long

With COM_PORT
  
 If .Started Then
 
     Send_String_Length = Len(Send_String)

     Synchronous_Write .Handle, Send_String, Send_String_Length, Bytes_Sent
    
     Send_Result = (Bytes_Sent = Send_String_Length)
 
 End If
  
End With

DoEvents

SEND_COM_PORT = Send_Result

End Function

Private Function SET_PORT_TIMERS() As Boolean

Dim Temp_Result As Boolean

Const NO_TIMEOUT As Long = -1
Const WRITE_CONSTANT As Long = 4000                           ' Should be less than approx 5000 to avoid VBA "Not Responding"
                                                              
With COM_PORT

    .Timeouts.Read_Interval_Timeout = NO_TIMEOUT              ' Timeouts not used for file reads.
    .Timeouts.Read_Total_Timeout_Constant = LONG_0            '
    .Timeouts.Read_Total_Timeout_Multiplier = LONG_0          '

    .Timeouts.Write_Total_Timeout_Constant = WRITE_CONSTANT   ' Maximum time in MilliSeconds allowed for each synchronous write
    .Timeouts.Write_Total_Timeout_Multiplier = LONG_0

     Temp_Result = Set_Com_Timers(.Handle, .Timeouts)

End With

SET_PORT_TIMERS = Temp_Result

End Function
