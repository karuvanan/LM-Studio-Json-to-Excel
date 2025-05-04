JSON conversion and parsing for VBA (Windows and Mac Excel, Access, and other Office apps). 
It is derived from the excellent project vba-json, with some additions and improvements to fix bugs and improve performance (as part of VBA-Web).

Tested in Windows Excel 2021 and  but should work on 2007, 2011, 2013

For Windows-only support, include a reference to "Microsoft Scripting Runtime"
For Mac and Windows support, includes VBA-Dictionary

Extract the VBA-JSON-2.3.1.zip
Open the Microsoft Excel 2021
Import JsonConverter.bas into your project (Open VBA Editor, Alt + F11 File > Import File)
For Windows only, include reference to "Microsoft Scripting Runtime"
Import the lm studio json to excel.bas into your project

Click run the macro
goto Folder C:\Users\user\.lmstudio\conversations
Select the json file

The result will automatic post into sheet1

Final - Save excel into xlsm
