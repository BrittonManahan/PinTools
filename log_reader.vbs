dim objRegExp2 : set objRegExp2 = new RegExp
objRegExp2.Pattern = "KB[0-9]{4,8}"
dim temp,y
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("C:\\Windows\\SoftwareDistribution\\ReportingEvents.log",1,False,-1)
Set dict = CreateObject("Scripting.Dictionary")
Do Until f.AtEndOfStream
	temp = f.ReadLine
	If InStr(1,temp, "Successful:",1) > 0 Then
		'WScript.Echo temp
		'WScript.Echo " "
		'WScript.Echo " "
		Set result = objRegExp2.Execute(temp)
		If result.Count > 0 then
			If NOT dict.Exists(Trim(result.Item(0))) Then
					dict.add Trim(result.Item(0)),0
			End If
			'WScript.Echo temp
			'WScript.Echo " "
			'WScript.Echo
			'WScript.Echo (Trim(result.Item(0)))
		End If
	End If
Loop

f.Close

For each y in dict
	WScript.Echo y
Next