the comparer interface compares _A to B_ and returns the result as enum _"compareResult"_.
```
 0 = A equal B
 1 = A greater than B
-1 = A less than B
```

### colorRGB
_very close_ to reality but not achieve _100%_ as the comparison is done on the _long_ value of the colors.

### caseSensitive
uses _[StrComp](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/strcomp-function)_ function of vba with binary compare.
### ignorCase
uses _[StrComp](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/strcomp-function)_ function of vba with text compare.
