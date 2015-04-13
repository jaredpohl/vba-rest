# VBA Project: **vba-rest**
## VBA Module: **[Main](/scripts/Main.vba "source is here")**
### Type: StdModule  

This procedure list for repo (vba-rest) was automatically created on 28/03/2015 7:37:57 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in Main

---
VBA Procedure: **uploadToDB**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function uploadToDB()*  

**no arguments required for this procedure**


---
VBA Procedure: **gtExampleLoad**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function gtExampleLoad()*  

**no arguments required for this procedure**


---
VBA Procedure: **gtDeadDropLoad**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function gtDeadDropLoad()*  

**no arguments required for this procedure**


---
VBA Procedure: **gtExampleMakeManifestScriptDbCom**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  
Description: ****  

*Private Function gtExampleMakeManifestScriptDbCom()*  

**no arguments required for this procedure**


---
VBA Procedure: **gtExampleMakeManifest**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  
Description: ****  

*Private Function gtExampleMakeManifest()*  

**no arguments required for this procedure**


---
VBA Procedure: **gtExampleMakeManifestCrest**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  
Description: ****  

*Private Function gtExampleMakeManifestCrest()*  

**no arguments required for this procedure**


---
VBA Procedure: **gtClassDocumenter**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  
Description: ****  

*Private Function gtClassDocumenter()*  

**no arguments required for this procedure**


---
VBA Procedure: **gtCreateReferences**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  
Description: ****  

*Private Function gtCreateReferences(dom As Object)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dom|Object|False||


---
VBA Procedure: **gtUpdateAll**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  
Description: ****  

*Private Function gtUpdateAll()*  

**no arguments required for this procedure**


---
VBA Procedure: **gtCoIndex**  
Type: **Function**  
Returns: **Long**  
Scope: **Private**  
Description: ****  

*Private Function gtCoIndex(sid As Variant, co As Collection) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sid|Variant|False||
co|Collection|False||


---
VBA Procedure: **gtPreventCaching**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function gtPreventCaching(url As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||


---
VBA Procedure: **gtDoit**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Public Function gtDoit(gtDoitmanifestID As String, Optional greenField As Boolean = False) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
gtDoitmanifestID|String|False||
greenField|Boolean|True| False|


---
VBA Procedure: **gtAddReference**  
Type: **Function**  
Returns: **Object**  
Scope: **Private**  
Description: ****  

*Private Function gtAddReference(name As String, guid As String, major As String, minor As String) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||
guid|String|False||
major|String|False||
minor|String|False||


---
VBA Procedure: **gtStampManifest**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function gtStampManifest(vbCom As Object, line As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
vbCom|Object|False||
line|Long|False||


---
VBA Procedure: **gtInsertStamp**  
Type: **Function**  
Returns: **Long**  
Scope: **Private**  
Description: ****  

*Private Function gtInsertStamp(vbCom As Object, manifest As String, rawUrl As String) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
vbCom|Object|False||
manifest|String|False||
rawUrl|String|False||


---
VBA Procedure: **gtWillItWork**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  
Description: ****  

*Private Function gtWillItWork(dom As Object, Optional greenField As Boolean = False) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dom|Object|False||
greenField|Boolean|True| False|


---
VBA Procedure: **gtAddStr**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function gtAddStr(t As String, n As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
t|String|False||
n|String|False||


---
VBA Procedure: **gtRecreateManifest**  
Type: **Function**  
Returns: **Object**  
Scope: **Private**  
Description: ****  

*Private Function gtRecreateManifest(manifestID As String) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
manifestID|String|False||


---
VBA Procedure: **gtModuleExists**  
Type: **Function**  
Returns: **Object**  
Scope: **Private**  
Description: ****  

*Private Function gtModuleExists(name As String, wb As Workbook) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||
wb|Workbook|False||


---
VBA Procedure: **gtAddModule**  
Type: **Function**  
Returns: **Object**  
Scope: **Private**  
Description: ****  

*Private Function gtAddModule(name As String, wb As Workbook, modType As String) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
name|String|False||
wb|Workbook|False||
modType|String|False||


---
VBA Procedure: **gtConstructRawUrl**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function gtConstructRawUrl(gistID As String, Optional gistFileName As String = vbNullString) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
gistID|String|False||
gistFileName|String|True| vbNullString|


---
VBA Procedure: **gtAddToManifest**  
Type: **Function**  
Returns: **Object**  
Scope: **Private**  
Description: ****  

*Private Function gtAddToManifest(dom As Object, gistID As String, modType As String, modle As String, Optional Filename As String = vbNullString, Optional version As String = vbNullString ) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dom|Object|False||
gistID|String|False||
modType|String|False||
modle|String|False||
Filename|String|True| vbNullString|
version|String|True| vbNullString|


---
VBA Procedure: **gtAddRefToManifest**  
Type: **Function**  
Returns: **Object**  
Scope: **Private**  
Description: ****  

*Private Function gtAddRefToManifest(dom As Object, r As Object) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dom|Object|False||
r|Object|False||


---
VBA Procedure: **gtInitManifest**  
Type: **Function**  
Returns: **Object**  
Scope: **Private**  
Description: ****  

*Private Function gtInitManifest(Optional description As String = vbNullString, Optional contact As String = vbNullString) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
description|String|True| vbNullString|
contact|String|True| vbNullString|


---
VBA Procedure: **gtHttpGet**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function gtHttpGet(url As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||


---
VBA Procedure: **gtStampLog**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function gtStampLog(manifest As String, rawUrl As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
manifest|String|False||
rawUrl|String|False||


---
VBA Procedure: **gtStamp**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function gtStamp() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **gtManageable**  
Type: **Function**  
Returns: **Long**  
Scope: **Private**  
Description: ****  

*Private Function gtManageable(vbCom As Object) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
vbCom|Object|False||
