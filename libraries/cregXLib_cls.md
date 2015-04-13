# VBA Project: **vba-rest**
## VBA Module: **[cregXLib](/libraries/cregXLib.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (vba-rest) was automatically created on 28/03/2015 7:37:56 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cregXLib

---
VBA Procedure: **Pattern**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get Pattern() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **Pattern**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Property Let Pattern(p As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|String|False||


---
VBA Procedure: **name**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get name() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **name**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Property Let name(p As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|String|False||


---
VBA Procedure: **ignorecase**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Public Property Get ignorecase() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **ignorecase**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Property Let ignorecase(p As Boolean)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Boolean|False||


---
VBA Procedure: **rGlobal**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Public Property Get rGlobal() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **rGlobal**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Property Let rGlobal(p As Boolean)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Boolean|False||


---
VBA Procedure: **init**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Sub init(sname As String, Optional spat As String = "", Optional bIgnoreSpaces As Boolean = True, Optional bIgnoreCase As Boolean = True, Optional bGlobal As Boolean = True)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sname|String|False||
spat|String|True| ""|
bIgnoreSpaces|Boolean|True| True|
bIgnoreCase|Boolean|True| True|
bGlobal|Boolean|True| True|


---
VBA Procedure: **getString**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Function getString(sFrom As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sFrom|String|False||


---
VBA Procedure: **getGroup**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Function getGroup(sFrom As String, groupNumber As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sFrom|String|False||
groupNumber|Long|False||


---
VBA Procedure: **getReplace**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Function getReplace(sFrom As String, sTo As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sFrom|String|False||
sTo|String|False||


---
VBA Procedure: **getTest**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Function getTest(sFrom As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sFrom|String|False||
