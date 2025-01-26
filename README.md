# vba
Contains code to be shared between workbooks/projects.  
They are mainly outcome of personal projects and so not continously maintained.  
General principles:
- late binding
- no additional software
- no releases
- semantic versioning on module/cls/form level
- no mac support (maybe later)

## projects:
each _"project"_ has its own directory and a readme about it.
- [comparer](https://github.com/SkyFips/vba/tree/main/comparer)  
extends the possibility to compare objects, based on your own (implemented) rule
- [enumerator](https://github.com/SkyFips/vba/tree/main/enumerator)  
enumerator interface to be able to sort objects
- [export/import](https://github.com/SkyFips/vba/tree/main/exportImport)  
possibility to export/import classes/modules/forms, based on a dedicated list
- [helper](https://github.com/SkyFips/vba/tree/main/helper)  
helper modules with no(t yet) clear assignment like ISO8601, JSON, ... conversion
- [logger](https://github.com/SkyFips/vba/tree/main/logger)  
a logging mechanism based on events or interface implemenation  
- [sorter](https://github.com/SkyFips/vba/tree/main/sorter)  
implements the bubble or quick sort algorithm for collection or iEnumerator


## list of other vba sources
during my journey in VBA, many different domains/repos, crossed my path and below an incomplete list of some.
### repositories:
- [VBA-tools](https://github.com/VBA-tools)
- [list of vba repos](https://github.com/sancarn/awesome-vba)

### domains:
- [mrExcel](https://www.mrexcel.com/)
- [betterSolutions](https://bettersolutions.com/)
