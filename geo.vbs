GoogleAPIKey = "PASTE YOUR GOOGLE API KEY HERE"

Set objShell = WScript.CreateObject("WScript.Shell")
Set objExecObject = objShell.Exec("netsh wlan show networks mode=BSSID")

buildStations = "{" & chr(34) & "wifiAccessPoints" & chr(34) & ":["
Do While Not objExecObject.StdOut.AtEndOfStream
	currLine = objExecObject.StdOut.ReadLine()
	if InStr(currLine, "not running") then 'No WiFi card enabled, get out
		Wscript.Quit
	end if
	if InStr(currLine, "BSSID") then
		buildStations = buildStations & "{" & chr(34) & "macAddress" & chr(34) & ":" & chr(34) & replace(Split(currLine, ": ")(1), ":", "-") & chr(34)
	end if

	if InStr(currLine, "Signal") then
		buildStations = buildStations & "," & chr(34) & "signalStrength" & chr(34) & ":-" & 100 - replace(Split(currLine, ": ")(1), "%", "") & "},"
	end if
Loop

buildStations = buildStations & "]}": buildStations = replace(buildStations, ",]}", "]}")

bcount = 0
countBSSIDs = Split(buildStations, ":" )
For Each BSSID in countBSSIDs
	If InStr(BSSID, "macAddress") Then
		bcount = bcount + 1
	End If
Next

If bcount < 2 then 'Not enough BSSIDs, get out of here
	Wscript.Quit
End if

LatLongResponse = POSTRequest("https://www.googleapis.com/geolocation/v1/geolocate?key=" & GoogleAPIKey, buildStations)

LatLongResponse = split(LatLongResponse, ",")

for each item in LatLongResponse
	if InStr(item, "lat") then
		latitude = replace(replace(Split(item, ": ")(2), "}", ""), vbLf, "")
	end if
	if InStr(item, "lng") then
		longtitude = replace(replace(Split(item, ": ")(1), "}", ""), vbLf, "")
	end if
	if InStr(item, "accuracy") then
		accuracy = replace(replace(Split(item, ": ")(1), "}", ""), vbLf, "")
	end if
Next

physical_address = POSTRequest("https://maps.googleapis.com/maps/api/geocode/json?latlng=" & latitude & "," & longtitude, "")
physical_address = split(physical_address, vbLf)

for each item in physical_address
	if InStr(item, "formatted_address") then
		found = replace(replace(replace(Split(item, ": ")(1), "}", ""), vbLf, ""), chr(34), "")
		wscript.echo "ADDRESS: " & LEFT(found, (LEN(found)-1))
		wscript.echo "MAP: https://www.google.com/maps/place/" & latitude & "," & longtitude
		wscript.echo "ACCURACY: " & accuracy & "m"
		exit for 'Got the first formatted address, get out.
	end if
Next

Function POSTRequest(url, request)
	set http = CreateObject("Microsoft.XMLHTTP")
	http.open "POST", url,false
	http.setRequestHeader "Content-Type", "application/json"
	http.setRequestHeader "Content-Length", Len(request)
	http.send request
	POSTRequest = http.responseText
End Function
