# VBA Project: vba-rest
This cross reference list for repo (vba-rest) was automatically created on 28/03/2015 7:37:57 PM by VBAGit.For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")
You can see [library and dependency information here](dependencies.md)

###Below is a cross reference showing which modules and procedures reference which others
*module*|*proc*|*referenced by module*|*proc*
---|---|---|---
cBrowser||Main|uploadToDB
cCell||cDataRow|create
cDataColumn||cDataSet|create
cDataRow||cDataSet|filterOk
cDataRow||cDataSet|create
cDataSet||Main|uploadToDB
cDataSets||cDataSet|populateData
cHeadingRow||cDataSet|Class_Initialize
cJobject||Main|uploadToDB
cregXLib||regXLib|rxMakeRxLib
cStringChunker||cJobject|recurseSerialize
cStringChunker||cJobject|unSplitToString
cStringChunker||cJobject|serialize
regXLib|rxReplace|cDataSet|populateGoogleWire
usefulcJobject|toISODateTime|cDataSet|jObject
usefulStuff|Base64Encode|cBrowser|httpGET
