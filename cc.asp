<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>AKSA Phone System Cost Center Summary</title>
<link href='phone.css' rel='stylesheet' type='text/css'>
</head>
<body>
<h1>AKSA Phone System Reports - Past 30 Days</h1>
<%
'on error resume next
dim Start
dim cc
dim objCfg
dim objCalldata
dim objRS
dim objRSdo
dim intTalkTime
dim intHour
dim intMin
dim intSec
dim strTalkTime

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3

Start = DateAdd("d", -30, Date)
cc = Request("cc")

set objCfg = CreateObject("adodb.connection")
'objCfg.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Shoreline Data\Database\ShoreWare.mdb"
objCfg.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=omaha;database=shoreware;port=4308;user=st_configread; pwd=passwordconfigread;"

set objCalldata = CreateObject("adodb.connection")
objCalldata.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\AssetMgt\Phone.mdb"

set objRS = CreateObject("adodb.recordset")
set objRSdo = CreateObject("adodb.recordset")

objRS.CursorLocation = adUseClient
objRSdo.CursorLocation = adUseClient

if cc = "" then
  objRS.open "select ExtensionLists.Name, ExtensionListItems.UserDN from ExtensionListItems " & _
    "inner join (select ExtensionLists.Name, ExtensionLists.ExtensionListID from ExtensionLists) " & _
    "as ExtensionLists on ExtensionLists.ExtensionListID=ExtensionListItems.ExtensionListID " & _
    "order by ExtensionListItems.UserDN" _
  , objCfg, adOpenStatic, adLockOptimistic

  Response.Write "<h1>Summary for All Cost Centers</h1>"

else
  objRS.open "select ExtensionLists.Name, ExtensionListItems.UserDN from ExtensionListItems " & _
    "inner join (select ExtensionLists.Name, ExtensionLists.ExtensionListID from ExtensionLists) " & _
    "as ExtensionLists on ExtensionLists.ExtensionListID=ExtensionListItems.ExtensionListID " & _
    "where ExtensionLists.ExtensionListID=" & cc & " " & _
    "order by ExtensionListItems.UserDN" _
  , objCfg, adOpenStatic, adLockOptimistic

Response.Write "<h1>Summary for Cost Center " & objRS("Name") & "</h1>"
end if

Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Cost Center Summary</caption>"
Response.Write "<tr><th>Extension</th><th>Description</th><th>Internal Calls</th><th>Inbound Calls</th><th>Outbound Calls</th><th>Talk Time</th></tr>"

Do While not objRS.EOF

  intTalkTime = 0

  Response.Write "<tr><td class='center'>" &  objRS("UserDN") & "</td>"

  objRSdo.open "select Description from DN where DN='" & objRS("UserDN") & "'", objCfg, adOpenStatic, adLockOptimistic
  Response.Write "<td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("UserDN") & chr(34) & ">" & objRSdo("Description") & "</a></td>"
  objRSdo.close

  objRSdo.open "select Extension, Description, Count(StartTime) as CountOfStartTime, Sum(TalkTime) as SumOfTalkTime from Calldata " & _
    "where Extension='" & objRS("UserDN") & "' and Internal=True " & _ 
    "group by Extension, Description" _
  , objCalldata, adOpenStatic, adLockOptimistic

  if not objRSdo.eof then
    if objRSdo("SumOfTalkTime") > 0 then
      intTalkTime = objRSdo("SumOfTalkTime")
    end if
    Response.Write "<td class='right'>" &  formatnumber(objRSdo("CountOfStartTime"), 0) & "</td>"
  else
    Response.Write "<td class='right'>0</td>"
  end if

  objRSdo.close

  objRSdo.open "select Extension, Description, Count(StartTime) as CountOfStartTime, Sum(TalkTime) as SumOfTalkTime from Calldata " & _
    "where Extension='" & objRS("UserDN") & "' and Internal=False and Inbound=True " & _ 
    "group by Extension, Description" _
  , objCalldata, adOpenStatic, adLockOptimistic

  if not objRSdo.eof then
    if objRSdo("SumOfTalkTime") > 0 then
      intTalkTime = intTalkTime + objRSdo("SumOfTalkTime")
    end if
    Response.Write "<td class='right'>" &  formatnumber(objRSdo("CountOfStartTime"), 0) & "</td>"
  else
    Response.Write "<td class='right'>0</td>"
  end if

  objRSdo.close

  objRSdo.open "select Extension, Description, Count(StartTime) as CountOfStartTime, Sum(TalkTime) as SumOfTalkTime from Calldata " & _
    "where Extension='" & objRS("UserDN") & "' and Internal=False and Inbound=False " & _ 
    "group by Extension, Description" _
  , objCalldata, adOpenStatic, adLockOptimistic

  if not objRSdo.eof then
    if objRSdo("SumOfTalkTime") > 0 then
      intTalkTime = intTalkTime + objRSdo("SumOfTalkTime")
    end if
    Response.Write "<td class='right'>" &  formatnumber(objRSdo("CountOfStartTime"), 0) & "</td>"
  else
    Response.Write "<td class='right'>0</td>"
  end if

  objRSdo.close

  intHour = intTalkTime \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (intTalkTime mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (intTalkTime mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec

  Response.Write "<td class='right'>" & strTalkTime & "</td></tr>"

  objRS.MoveNext

loop

objRS.close
Response.Write "</table></p>"

%>

</body>
</html>
