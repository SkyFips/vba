# Logger
Two different approaches to provide logging in vba.  

## Event driven
The main class is _cLogger.cls_, which fires _"printLog"_ whenever a Log is called.  
Availalbe _"log printers"_ are:
- file
- watch window
possible to add one or more printers to that class via _"addPrinters"_.

## interface
Each logger implements iLogger interface and prints to the defined output, like watch window, file or userform.  
