# Installing and Using functions as Excel Formulas
## Worksheet Cell Formulas


<details><summary>Before starting</summary>
<p>

* Check that the Com Port settings match on both the PC and attached device.  
* Connect a terminal emulator or similar device to the PC's COM Port. 
* See [screenshot](MODE_COM1.jpg) in folder for Com Port settings check on host PC.

</p>
</details>   

<details><summary>Installing VBA</summary>
<p>

1.  Download SERIAL_PORT_EXTRA_SIMPLE_VBAn.bas 
2.  Open a new Excel document 
3.  Enter the VBA Environment (Alt-F11)
4.  From VBA Environment, view the Project Explorer (Control-R)
5.  From Project Explorer, right-hand click and select Import File.
6.  Import the file SERIAL_PORT_EXTRA_SIMPLE_VBAn.bas
7.  Check that a new module SERIAL_PORT_VBA_SIMPLE is created and visible in the Modules folder. 
8.  Check/Edit `COM_PORT_NUMBER` value at start of module SERIAL_PORT_VBA_SIMPLE  
9.  In function `READ_COM_PORT`, remove the comment mark before `Application.Volatile`    
10. Close and return to Excel (Alt-Q)  
11. IMPORTANT - save document as type **Macro-Enabled** with a file name of your choice.
  

</p>
</details>   

<details><summary>Excel Formula Testing</summary>
<p>

<details><summary>Start Com Port</summary>
<p>  
  
* In Cell **A3**, type the formula `=start_com_port()` and hit return
* Check that `TRUE` is now shown in cell **A3**
* `TRUE` confirms that the port has started.
  
</p>
</details>    
 
<details><summary>Send to Com Port</summary>
<p>    
  
* Enter some short text in Cell **B5** - e.g. **TEST123**
* In Cell **A5**, type the formula `=send_com_port(B5)` and hit return
* Check that `TRUE` is now shown in cell **A5**  
* Check that **TEST123** appears on your device
* Change the text in Cell **B5** - e.g. **QWERTY** and hit return
* Check that `TRUE` is still shown in cell **A5**  
* Check that **QWERTY** appears on your device
* This confirms that `send_com_port` is working.    
  
  </p>
</details>   
  
<details><summary>Read from Com Port</summary>
<p>    
  
* In Cell **B7**, type the formula `=read_com_port()` and hit return
* Enter some text on your device - e.g. **HELLO**
* Change any other cell or press F9 key on your PC 
* Check that **HELLO** appears in Cell **B7**
* Change any other cell or press F9 key on your PC for a second time
* Check that Cell **B7** returns to blank (no new data to read)
* Enter some new text on your device - e.g. **AGAIN** 
* Change any other cell or press F9 key on your PC for a third time
* Check that **AGAIN** appears in Cell **B7**  
* This confirms that `read_com_port` and `Application.Volatile` are working.  
  
</p>
</details>     

 <details><summary>Stop Com Port</summary>
<p>  
  
* In Cell **A9**, type the formula `=stop_com_port()` and hit return
* Check that `TRUE` is now shown in cell **A9**
* `TRUE` confirms that the port has stopped.
* Change any other cell or press F9 key on your PC
* Check that **FALSE** appears in Cell **A5** _(send_com_port has failed as expected)_
* This confirms that `stop_com_port` is working. 
  
</p>
</details> 
  
</p>
</details>    
  
<details><summary>Serial Devices</summary>
<p> 
  
#### Passive Devices  

These devices do **not** need a command to be sent before replying.  
Reads should function from Excel with no further action required.
 
#### Active Devices 
  
These devices  **do** need a command to be sent before replying.  

  A read delay will normally be required to allow sufficient time for the :-
  
  a) device to process the read command  
  b) serial data to be transmitted back 

Remove the comment mark from `Kernel_Read_Milliseconds` in function `read_com_port`
    
</p>
</details>   

Note that Functions may still be used in VBA routines as required. 
