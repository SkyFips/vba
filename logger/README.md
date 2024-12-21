# Logger
A logging mechanism is essential, when providing small tools to the public/others.

## Event driven logging
Event driven logging is able via _cLogger.cls_, which fires an event, whenever a log is printed.  
Printers can be added do that logger.

## interface logging
Each logger implements iLogger interface and prints to the defined output, like watch window, file or userform.  
