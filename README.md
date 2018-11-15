# rs232_monitor2
Collect COM ports (up to 4) data into 1 readable file for analyzing

Software intended for support engineers who needs to collect several data flows into a timed database to analyse the process flow.
T-cable drawing attached is the way to collect 2-way data from any RS232 connection. Sometimes one needs to collect up to several channels simultaneously so the maximum COM-port number extended to 4.
Take care of sorting data by time because Windows is not a real-time OS and data chunks can be mixed while saving them in the file even being captured correctly. Timestamps helps to keep the timeline correct.
User can disable log output to textbox and grid table to reduce CPU load. Also RS232 signals monitoring (DCD, RTS, CTS, DTR, DSR) can be disabled in settings.
See .config files for some system-related settings.
----
Real world example: to analyze behavior of PC->fiscal controller->printer setup. One needs to monitor:
  1) Commands: PC -> Fiscal
  2) Replies: Fiscal -> PC
  3) Commands: Fiscal -> Printer
  4) Replies: Printer -> Fiscal
  
It takes 2 T-cables, connected between PC and Fiscal and Fiscal and Printer. Connect T-cables to PC via USB hub + 4xUSB-COM adapters. Run the software and collect the data.
Load it into Open/LibreOffice Calc via CSV import and sort by Time+milliseconds columns to be sure it's ordered. Now data can be filtered by source port and other parameters to be analyzed.
Extracted command data can be sent to device with ComPrnControl or UsbPrnControl to replay or get analyzed with EscPosParser2.
