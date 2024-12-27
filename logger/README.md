# Logger
Two different approaches to provide logging in vba.  

## [Event driven](./event/)
The main class is _cLogger.cls_, with event _"printLog"_ which is raised whenever a Log is printed.  
Availalbe _"log printers"_ are:
- file
- watch window
- user form  

Possible to add one or more printers via _"addPrinters"_, which implements the interface _"iLogPrinter"_.

## [interface](./interface/)
Each available _logger_ implements the interface _iLogger_ and uses the defined output. Depending on the _logger_, a different output is created.  
Available loggers are:  
- file
- watch window
- user form
