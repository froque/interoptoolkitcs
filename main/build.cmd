if exist output rd output /s /q
.\tools\7-Zip-9.20\7za.exe a "output\ProjectTemplates\Visual C#\Windows\CSInteropUserControlProjectTemplate.zip" .\src\Templates\CSInteropUserControlProjectTemplate\*
.\tools\7-Zip-9.20\7za.exe a "output\ItemTemplates\Visual C#\CSInteropUserControlItemTemplate.zip" .\src\Templates\CSInteropUserControlItemTemplate\*
.\tools\7-Zip-9.20\7za.exe a output\InteropFormsToolkitCS.zip .\output\*
if exist output\ProjectTemplates rd output\ProjectTemplates /s /q
if exist output\ItemTemplates rd output\ItemTemplates /s /q