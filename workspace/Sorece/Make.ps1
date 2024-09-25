csc.exe /target:winexe /win32icon:..\Icon\MD_Explorer.ico /out:..\MD_Explorer.exe `
/reference:"C:\Program Files (x86)\Reference Assemblies\Microsoft\WindowsPowerShell\3.0\System.Management.Automation.dll" `
/reference:"Dll\NPOI.dll" `
/reference:"Dll\NPOI.OOXML.dll" `
/reference:"Dll\NPOI.OpenXml4Net.dll" `
/reference:"Dll\NPOI.OpenXmlFormats.dll" `
/reference:"Dll\itextsharp.dll" `
/reference:"Dll\Xceed.Words.NET.dll" `
"Event\*.cs"  "Main\*.cs" "Setting\*.cs"