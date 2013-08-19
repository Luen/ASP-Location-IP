<% 'http://en.utrace.de/api.php

		'...The old API will be disabled on Nov 15th 12:00am GMT 2010. Existing users are requested to update their APIs as soon as possible...
		'API key
		key = "INPUT API KEY HERE"


		' fraud detection - http://ipinfodb.com/fraud_detection.php
		'http://api.ipinfodb.com/v2/fraud_query.php?key=<your_api_key>&ip=74.125.45.100&country_code=us&district=94043&area_code=650&mail_domain=ipinfodb.com&city=Mountain+View


		'for information on HTTP_X_FORWARDED_FOR and the problems with it.
		'http://bytes.com/topic/asp-classic/answers/532376-http_x_forwarded_for
		'http://meatballwiki.org/wiki/AnonymousProxy
		'http://www.freeproxy.ru/en/free_proxy/faq/proxy_anonymity.htm
		Dim CustomerIP, key
			CustomerIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If CustomerIP = "" or Trim(LCase(CustomerIP)) = "unknown" Then
			CustomerIP = Request.ServerVariables("REMOTE_ADDR")
		End If

		' http://www.ip-adress.com/whois/150.101.181.14
		' http://en.utrace.de/ip-address/150.101.181.14
		' http://ip-whois-lookup.com/lookup.php?ip=41.211.217.80



		'--------------------
		' LOCATION script

		' Variables
		Dim objXmlHttp, xmlDoc, strMainServer, strBackupServer
		' xml Variables
		Dim Status, CountryCode, CountryName, RegionName, City, Latitude, Longitude, Timezone

		If (Request.Cookies("location_cookie_" & CustomerIP) = "" OR InStr(Request.Cookies("location_cookie_" & CustomerIP), "ip=&status=") > 0) Then
			' Create XMLHTTP and XML DOM Document objects and set initialize other variable values
			Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
			Set xmlDoc = Server.CreateObject( "MSXML2.DOMDocument.3.0" )
			'strMainServer = "http://ipinfodb.com/ip_query.php?ip="
			'strBackupServer = "http://backup.ipinfodb.com/ip_query.php?ip="
			strMainServer = "http://api.ipinfodb.com/v2/ip_query.php?key=" & key & "&ip="
			strBackupServer = "http://backup.ipinfodb.com/ip_query.php?key=" & key & "&ip="
			Status = ""
			CountryCode = ""
			CountryName = ""
			RegionName = ""
			City = ""
			Latitude = ""
			Longitude = ""
			Timezone = ""

			' Query the main server
			objXmlHttp.open "GET", strMainServer & CustomerIP, False
			objXmlHttp.send

			' If we do not get an HTTP 200 response from the main server, query the backup server
			If objXmlHttp.status <> 200 Then
				objXmlHttp.open "GET", strBackupServer & CustomerIP, False
				objXmlHttp.send
			End If

			' If we got an HTTP 200 response from either the main server or backup server, read and parse the XML response
			If objXmlHttp.status = 200 Then
				xmlDoc.loadXML(objXmlHttp.responseXML.xml)

				' If there are no parsing errors while reading the XML response, set the country code based on the CountryCode node value
				If xmlDoc.parseError.errorcode = 0 Then
					Status = xmlDoc.documentElement.selectSingleNode("Status").text
					CountryCode = xmlDoc.documentElement.selectSingleNode("CountryCode").text
					CountryName = xmlDoc.documentElement.selectSingleNode("CountryName").text
					RegionName = xmlDoc.documentElement.selectSingleNode("RegionName").text
					City = xmlDoc.documentElement.selectSingleNode("City").text
					Latitude = xmlDoc.documentElement.selectSingleNode("Latitude").text
					Longitude = xmlDoc.documentElement.selectSingleNode("Longitude").text
				End If
			End If

			' Clean up
			Set xmlDoc = Nothing
			Set objXmlHttp = Nothing
			
			
			'COOKIE
			'create a cookie so the servers do not have to request the xml file again.
			'set cookie to reduce trafic to the servers
			location_cookie=request.cookies("location_cookie_" & CustomerIP)
			response.cookies("location_cookie_" & CustomerIP).Expires=date+30
			response.cookies("location_cookie_" & CustomerIP)="ip=" & CustomerIP & "&status=" & Status & "&countrycode=" & CountryCode & "&countryname=" & CountryName & "&regionname=" & RegionName & "&city=" & City & "&latitude=" & Latitude & "&longitude=" & Longitude & "&timezone=" & Timezone
		End If


		Dim location_cookie
		location_cookie = Request.Cookies("location_cookie_" & CustomerIP)





		parseQueryString(location_cookie)
		If ip = "" Then
			ip = CustomerIP
		End If

		'response.write location_cookie & "<br>"
		'response.write "ip=" & ip & "&status=" & status & "&countrycode=" & countrycode & "&countryname=" & countryname & "&regionname=" & regionname & "&city=" & city & "&latitude=" & latitude & "&longitude=" & longitude & "&timezone=" & timezone & "<br><br>"

		If LCase(status) = "ok" Then

		End If


		'dim location_cookie_array
		'location_cookie_array = Split(location_cookie, "&")
		'If UBound(location_cookie_array) = "7" Then
		'	If CustomerIP = location_cookie_array(0) Then
		'		If LCase(location_cookie_array(1)) = "ok" Then
		'			Status = location_cookie_array(1)
		'			CountryCode = location_cookie_array(2)
		'			CountryName = location_cookie_array(3)
		'			RegionName = location_cookie_array(4)
		'			City = location_cookie_array(5)
		'			Latitude = location_cookie_array(6)
		'			Longitude = location_cookie_array(7)
		'		End If
		'	End If
		'End If





		Dim location
			location = ""
		If (City <> RegionName) Then
			If (City <> "") Then
				location = location & City & ", "
			End If
			If (RegionName <> "") Then
				location = location & RegionName & ", "
			End If
		Else
			If (RegionName <> "" AND CountryName <> RegionName) Then
				location = location & RegionName & ", "
			End If
		End If
		If (CountryName <> "") Then
			location = location & CountryName
		End If

		' End LOCATION script
		'--------------------





		'http://www.codingforums.com/archive/index.php/t-59878.html
		function parseQueryString(temp) 'function to parse the string begins
			thepairs = Split(temp, "&")'Split thequery at the comma
			for each item in thepairs 'Begin loop through the querystring
				pair = item
				thevalue = Split(pair, "=")
				If ubound(thevalue) = 2 Then
					If (thevalue(0) = "ip") Then
						ip = thevalue(1)
					Elseif (thevalue(0) = "status") Then
						status = thevalue(1)
					Elseif (thevalue(0) = "countrycode") Then
						countrycode = thevalue(1)
					Elseif (thevalue(0) = "countryname") Then
						countryname = thevalue(1)
					Elseif (thevalue(0) = "regionname") Then
						regionname = thevalue(1)
					Elseif (thevalue(0) = "city") Then
						city = thevalue(1)
					Elseif (thevalue(0) = "latitude") Then
						latitude = thevalue(1)
					Elseif (thevalue(0) = "longitude") Then
						longitude = thevalue(1)
					Elseif (thevalue(0) = "timezone") Then
						timezone= thevalue(1)
					End If
				End If
			next
		End Function
		
		
%>
