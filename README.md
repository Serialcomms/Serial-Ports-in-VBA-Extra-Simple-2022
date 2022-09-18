# Serial Ports in VBA for 2022

**Minimal version for basic send and receive data use only**

This is the further simplified set of VBA Serial Port functions for use with one pre-defined Com Port only.

Intended for use with serial devices which have a well-defined set of short commands and responses.

Functions can be used directly in Excel Worksheet Cells.

_Assumes that COM Port has previously been configured correctly via command-line or other method_



<P>

No functionality provided for


- Debugging

- Multiple COM Ports

- Waiting data check before read [^1]

- Device hardware (signalling) functions 
  
- COM Port settings to be modified on starting

</P>

[4 functions only - start, stop, send, read com port](Functions.md)

[^1] Read and Write in same Excel sheet may require read delay
