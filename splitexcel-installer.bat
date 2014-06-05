@ECHO OFF
mkdir C:\Progr\splitexcel
copy splitexcel.vbs C:\Progr\splitexcel\

reg add HKEY_CURRENT_USER\Software\Classes\Excel.Sheet.12\shell\split /ve /d "Split every sheet in a separate Excel file"
reg add HKEY_CURRENT_USER\Software\Classes\Excel.Sheet.12\shell\split\command /ve /d "C:\Windows\System32\wscript.exe \"C:\\Progr\\splitexcel\\splitexcel.vbs\" \"%%1\""

reg add HKEY_CURRENT_USER\Software\Classes\Excel.Sheet.8\shell\split /ve /d "Split every sheet in a separate Excel file"
reg add HKEY_CURRENT_USER\Software\Classes\Excel.Sheet.8\shell\split\command /ve /d "C:\Windows\System32\wscript.exe \"C:\\Progr\\splitexcel\\splitexcel.vbs\" \"%%1\""
