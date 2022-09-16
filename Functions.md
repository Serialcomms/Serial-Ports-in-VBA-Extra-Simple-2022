## Serial Port VBA Functions - Extra-simplified set

##### All functions support one pre-defined Com Port number with pre-configured settings only

| VBA Function                   |    Returns     | Description                                                                           |
| -------------------------------|----------------|---------------------------------------------------------------------------------------|
| `start_com_port`               | `Boolean` [^1] | Starts com port with existing settings                                                |
| `read_com_port`                | `String`       | Returns all [^3] waiting characters from com port                                     |
| `read_com_port(6)`             | `String`       | Returns up to [^3] specified number of waiting characters from com port               |
| `send_com_port("QWERTY")`      | `Boolean` [^1] | Sends [^2] supplied character string to com port                                      |
| `send_com_port(COMMANDS)`      | `Boolean` [^1] | Sends [^2] character string defined in VBA constant or variable COMMANDS to com port  |
| `stop_com_port`                | `Boolean` [^1] | Stops com port and returns its control back to Windows                                |

##### Com Port number defined in declarations section at start of module   
`Private Const COM_PORT_NUMBER as Long = 1`    

[^1]: Function returns `True` if successful, otherwise `False`  

[^2]: Function will block until all characters are sent or write timer expires.  
      Maximum characters sent limited by timer `Write_Total_Timeout_Constant` value.   
      Long strings may cause VBA 'Not Responding' condition until transmission complete or timer expires.    
      
[^3]: Maximum characters returned = read buffer length (fixed value)    
      More waiting characters beyond buffer length may remain unread.   
     
 
