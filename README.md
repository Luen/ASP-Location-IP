ASP Detect Location by IP
===============

Simple ASP function for getting the geo location for an IP address using ipinfodb database.


Usage
=============

```asp
<!--#include file="GeoIP.asp" -->
<%
response.write "Originating IP address: http://en.utrace.de/ip-address/" & CustomerIP & vbCrLf
response.write "Originating from: " & location & vbCrLf
%>
```
