Dim dict,strValueName
Set dict = CreateObject("Scripting.Dictionary")
dim result
dim good
dim objRegExp : set objRegExp = new RegExp
objRegExp.Pattern = "KB[0-9]{4,8}"
Const HKEY_LOCAL_MACHINE = &H80000002
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
strValueName = "CurrentState"
Set reg = GetObject("winmgmts://./root/default:StdRegProv")
reg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
For Each sKey In arrSubKeys
	Set result = objRegExp.Execute(sKey)
	if result.Count > 0 then
		If NOT dict.Exists(Trim(result.Item(0))) Then
			reg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath & "\" & sKey,strValueName,dwValue
			'WScript.Echo CStr(dwValue)
			dict.add Trim(result.Item(0)),dwValue
		Else
			reg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath & "\" & sKey,strValueName,dwValue
			'WScript.Echo CStr(dwValue)
			dict(Trim(result.Item(0))) = dict(Trim(result.Item(0))) & "," & dwValue
		End If
	End If
Next
good = 0
for each yo in dict.Keys
	'WScript.Echo "                 "
	'WScript.Echo yo
	'WScript.Echo "--------------------"
	wha = Split(dict(yo),",")
	for each w in wha
		'WScript.Echo w
		if w = "112" Then
			good = 1
		else
			good = 0
		End If
	Next
	if good = 1 then
		WScript.Echo yo
	End If
Next

