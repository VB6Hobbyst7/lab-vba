# Common Excel Addin
<sup>[Overview](#) · [Quickstart](#Quickstart) · [Features](#Features)</sup>

Common Excel Addin aims to provide commonly used formatting features in a single tab with customisable settings. 

## Quickstart[<sup>·</sup>]
- Download [Common Excel Addin v0.1.0](https://github.com/ry4nyeo/addins/blob/main/assets/commonaddin-r010.xlam)  
- Install the downloaded .xlam file. See [Excel guide] from Microsoft. 

## Features[<sup>·</sup>]

<img src ="../assets/img-commonaddin-tabmain.png" width="85%" height="85%">

### Custom Functions[<sup>·</sup>]

|Functions|Description|
|:--|:--|
|[**XCOMPARE**](../excel/source/XCOMPARE.bas)|Returns the word difference between two ranges.|
|[**XEXTRACTAFTER**](../excel/source/XEXTRACTAFTER.bas)|Returns the part of a selected range after a specified word.|
|[**XEXTRACTBEFORE**](../excel/source/XEXTRACTBEFORE.bas)|Returns the part of a selected range before a specified word.|
|[**XFIND**](../excel/source/XFIND.bas)|Returns the word search results on a selected range based on a specified word list.|
|[**XHASNUMBER**](../excel/source/XHASNUMBER.bas)|Returns True if there is a number in the selected range.|
|[**XLOOKUP**](../excel/source/XLOOKUP.bas)|Returns the matched lookup value from a search list.|
|[**XREPLACEWORDS**](../excel/source/XREPLACEWORDS.bas)|Replaces words in a selected range based on specified replacement word list.|
|[**XSPELLNUMBER**](../excel/source/XSPELLNUMBER.bas)|Spells monetary values in dollar and cents.|
|[**XSUBSTITUTEPREFIX**](../excel/source/XSUBSTITUTEPREFIX.bas)|Replaces the prefix of a selected range based on a specified replacement.|
|[**XSUBSTITUTESUFFIX**](../excel/source/XSUBSTITUTESUFFIX.bas)|Replaces the suffix of a selected range based on a specified replacement.|
|[**XTRANSLATE**](../excel/source/XTRANSLATE.bas)|Returns the Google Translation result on a selected range.|
|[**XCELLFORMULA**](../excel/source/XCELLFORMULA.bas)|Returns formula of the selected cell.|
|[**XCLEANTEXT**](../excel/source/XCLEANTEXT.bas)|Removes excess non-alphanumeric characters|
|[**XGETPAGENUMBER**](../excel/source/XGETPAGENUMBER.bas)|Returns page number.|
|[**XIFDATE**](../excel/source/XIFDATE.bas)|Returns True if it is date format.|
|[**XREMOVEBETWEEN**](../excel/source/XREMOVEBETWEEN.bas)|Removes text between two specified delimiters.|
|[**XREMOVESYMBOLS**](../excel/source/XREMOVESYMBOLS.bas)|Removes leading and trailing symbols from text.|
|[**XSHEETNAME**](../excel/source/XSHEETNAME.bas)|Returns worksheet name.|
|[**XSUBSTITUTEMULTIPLE**](../excel/source/XSUBSTITUTEMULTIPLE.bas)|Substitutes multiple words.|

### Examples[<sup>·</sup>]
<img src ="../assets/img-commonaddin-010.gif" width="60%" height="60%">

<br>

---
<sup>[Addins] > [Common Excel Addin]</sup>

[Addins]: https://github.com/ry4nyeo/addins
[Common Excel Addin]: ../assets/pgs-commonaddin-010.md
[<sup>·</sup>]: #

[Excel guide]: https://support.microsoft.com/en-us/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460#:~:text=COM%20add%2Din-,Click%20the%20File%20tab%2C%20click%20Options%2C%20and%20then%20click%20the,install%2C%20and%20then%20click%20OK.
