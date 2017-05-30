Dim objSession, AutoUpdate,g, result,hot,count, h, UpdateSearcher,qtd, SearchResult, Updatestodownload, Update, Downloader, UpdatesToInstall, Installer
Set objSession = CreateObject("Microsoft.Update.Session")
Set UpdateSearcher = objSession.CreateUpdateSearcher()
qtd = UpdateSearcher.GetTotalHistoryCount()
dim objRegExp : set objRegExp = new RegExp
objRegExp.Pattern = "KB[0-9]{4,8}"
Set dict = CreateObject("Scripting.Dictionary")
Set hot = UpdateSearcher.QueryHistory(0, qtd)
count = 0
On Error Resume Next
For each h in hot
	If (IsNull(h.operation) = 0 And IsNull(h.resultcode) = 0) Then
		If (h.resultcode = 2) Then
			Set result = objRegExp.Execute(h.Title)
			If result.Count > 0 then
				If NOT dict.Exists(Trim(result.Item(0))) Then
					dict.add Trim(result.Item(0)),0
					count = count + 1
				End If
			End If
		End if
	End If
Next

WScript.Echo count
For each g in dict
	WScript.Echo g
Next