Set WshShell = WScript.CreateObject("WScript.Shell")
WSHShell.Run chr(34) & "C:\Users\chbapie\Desktop\Bartosz\apiquitous\ngrok.exe" & chr(34) & " -subdomain=apiquitous 8001" & " /silent",0,true