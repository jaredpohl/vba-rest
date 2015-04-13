# VBA Project: **vba-rest**
## VBA Module: **[cDataSets](/libraries/cDataSets.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (vba-rest) was automatically created on 28/03/2015 7:37:56 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cDataSets

---
VBA Procedure: **dataSets**  
Type: **Get**  
Returns: **Collection**  
Scope: **Public**  
Description: ****  

*Public Property Get dataSets() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **dataSet**  
Type: **Get**  
Returns: **[cDataSet](/libraries/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Property Get dataSet(sn As String, Optional complain As Boolean = False) As cDataSet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sn|String|False||
complain|Boolean|True| False|


---
VBA Procedure: **name**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get name() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **create**  
Type: **Function**  
Returns: **[cDataSets](/libraries/cDataSets_cls.md "cDataSets")**  
Scope: **Public**  
Description: ****  

*Public Function create(Optional sname As String = "DataSets") As cDataSets*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sname|String|True| "DataSets"|


---
VBA Procedure: **init**  
Type: **Function**  
Returns: **[cDataSet](/libraries/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Function init(Optional rInput As Range = Nothing, Optional keepFresh As Boolean = False, Optional sn As String = vbNullString, Optional blab As Boolean = False, Optional blockstarts As String, Optional bLikely As Boolean = False, Optional sKey As String = vbNullString, Optional respectFilter As Boolean = False) As cDataSet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rInput|Range|True| Nothing|
keepFresh|Boolean|True| False|
sn|String|True| vbNullString|
blab|Boolean|True| False|
blockstarts|String|True||
bLikely|Boolean|True| False|
sKey|String|True| vbNullString|
respectFilter|Boolean|True| False|


---
VBA Procedure: **exists**  
Type: **Function**  
Returns: **[cDataSet](/libraries/cDataSet_cls.md "cDataSet")**  
Scope: **Private**  
Description: ****  

*Private Function exists(sid As Variant) As cDataSet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sid|Variant|False||


---
VBA Procedure: **teardown**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Sub teardown()*  

**no arguments required for this procedure**


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
