<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>AKSA Phone System Overview</title>
<link href='phone.css' rel='stylesheet' type='text/css'>
</head>
<body>
<h1>AKSA Phone System Reports - Past 30 Days</h1>
<%
'****************
'Bob Rhett - Thursday, August 12, 2010
'  Added trunk usage display
'Keith Brooks - Monday, October 3, 2011
'  Updated Shoreware database connection string to MySQL.
'****************
'on error resume next
dim Start
dim objCalldata
dim objCfg
dim objRS
dim objRSdo
dim objRSdodo
dim objRSCalldata
dim objRSCfg
dim strSQL
dim heightfactor

Const Top = 5
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3
Const admi = .2035
Const corp = .0678
Const dist = .0344
Const main = .1160
Const mark = .1547
Const poly = .0788
Const prod = .1489
Const qual = .0605
Const spin = .1168
Const warp = .0186

Start = DateAdd("d", -30, Date)

set objCalldata = CreateObject("adodb.connection")
objCalldata.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\Assetmgt\Phone.mdb"

set objCfg = CreateObject("adodb.connection")
'objCfg.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Shoreline Data\Database\ShoreWare.mdb"
objCfg.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=omaha;database=shoreware;port=4308;user=st_configread; pwd=passwordconfigread;"

set rs = CreateObject("adodb.recordset")
set objRS = CreateObject("adodb.recordset")
set objRSdo = CreateObject("adodb.recordset")
set objRSdodo = CreateObject("adodb.recordset")
set objRSCalldata = CreateObject("adodb.recordset")
set objRSCfg = CreateObject("adodb.recordset")

objRS.CursorLocation = adUseClient
objRSdo.CursorLocation = adUseClient
objRSdodo.CursorLocation = adUseClient
objRSCalldata.CursorLocation = adUseClient

'****************
'Trunk Usage
'****************
dim StartDate
dim strTitle
dim iHt
dim iWd
dim iHtFactor
dim iWdMin
dim strIMG

StartDate = DateAdd("d", -400, Date)
iHtFactor = 10
iWdMin = 12

response.write "<hr/>"
response.write "<p>"
response.write "<table border='1'>"
response.write "<caption>Trunk Usage</caption>"
response.write "<tr>"
response.write "<td>"
response.write "<table border='0' cellpadding='0' cellspacing='0'>"
response.write "<tr>"

strSQL = "select * from TrunkUsage where RecDate>#" & StartDate & "# order by RecDate"
rs.Open strSQL, objCalldata, 2, adLockOptimistic
do while not rs.eof
  strTitle = "On "
  strTitle = strTitle & formatdatetime(rs("RecDate"), vbLongDate)
  strTitle = strTitle & " peak usage was " & rs("Max")
  if rs("Max") = 1 then
    strTitle = strTitle & " channel"
  else
    strTitle = strTitle & " channels"
  end if
  strTitle = strTitle & " starting at "
  strTitle = strTitle & formatdatetime(rs("Peak"), vbLongTime)
  strTitle = strTitle & " and lasting " & rs("Dur")
  if rs("Dur") = 1 then
    strTitle = strTitle & " second. "
  else
    if rs("Cont") = True then
      strTitle = strTitle & " contiguous seconds. "
    else
      strTitle = strTitle & " non-contiguous seconds. "
    end if
  end if
  strTitle = strTitle & "Average usage for the day was " & rs("Avg")
  if rs("Avg") = 1 then
    strTitle = strTitle & " channel."
  else
    strTitle = strTitle & " channels."
  end if

'  iHt = (23 - rs("Max")) * iHtFactor
  iWd = iWdMin

  select case rs("Max")
    case 23
'      strIMG = "VerticalGraph/redline.gif"
      strIMG = "VerticalGraph/dark_green23bof23.gif"
      iWd = iWdMin + rs("Dur")
    case 22
      strIMG = "VerticalGraph/dark_green22bof23.gif"
    case 21
      strIMG = "VerticalGraph/dark_green21bof23.gif"
    case 20
      strIMG = "VerticalGraph/dark_green20bof23.gif"
    case 19
      strIMG = "VerticalGraph/dark_green19of23.gif"
    case 18
      strIMG = "VerticalGraph/dark_green18of23.gif"
    case 17
      strIMG = "VerticalGraph/dark_green17of23.gif"
    case 16
      strIMG = "VerticalGraph/dark_green16of23.gif"
    case 15
      strIMG = "VerticalGraph/dark_green15of23.gif"
    case 14
      strIMG = "VerticalGraph/dark_green14of23.gif"
    case 13
      strIMG = "VerticalGraph/dark_green13of23.gif"
    case 12
      strIMG = "VerticalGraph/dark_green12of23.gif"
    case 11
      strIMG = "VerticalGraph/dark_green11of23.gif"
    case 10
      strIMG = "VerticalGraph/dark_green10of23.gif"
    case 9
      strIMG = "VerticalGraph/dark_green9of23.gif"
    case 8
      strIMG = "VerticalGraph/dark_green8of23.gif"
    case 7
      strIMG = "VerticalGraph/dark_green7of23.gif"
    case 6
      strIMG = "VerticalGraph/dark_green6of23.gif"
    case 5
      strIMG = "VerticalGraph/dark_green5of23.gif"
    case 4
      strIMG = "VerticalGraph/dark_green4of23.gif"
    case 3
      strIMG = "VerticalGraph/dark_green3of23.gif"
    case 2
      strIMG = "VerticalGraph/dark_green2of23.gif"
    case 1
      strIMG = "VerticalGraph/dark_green1of23.gif"
    case 0
      strIMG = "VerticalGraph/dark_green0of23.gif"
    case else
      strIMG = "VerticalGraph/dark_green23.gif"
  end select

  response.write "<td class='center'>"


'  response.write "<img src='VerticalGraph/trans.gif' width='" & iWd + 2 & "' height='" & iHt & "' border='0' title='" & strTitle & "'>"
'  response.write "<br />"
'  iHt = (rs("Max") - rs("Avg")) * iHtFactor
'  iHt = rs("Max") * iHtFactor
  iHt = 184
  response.write "<img src='" & strIMG & "' width='" & iWd & "' height='" & iHt & "' border='0' title='" & strTitle & "'>"
'  response.write "<br />"
'  if rs("Avg") > 0 then
'    iHt = (rs("Avg") * iHtFactor) + 1
'    response.write "<img src='VerticalGraph/light_green.gif' width='" & iWd & "' height='" & iHt & "' border='1' title='" & strTitle & "'>"
'  else
'    response.write "<img src='VerticalGraph/light_green.gif' width='" & iWd & "' height='0' border='1'>"
'  end if
  response.write "<br />"
  if left(weekdayname(weekday(rs("RecDate")), True), 1) <> "S" then
    response.write "<h5 />"
  else
    response.write "<h6 />"
  end if
  response.write left(weekdayname(weekday(rs("RecDate")), True), 1)
  response.write "</td>"
  rs.movenext
loop
response.write "</tr>"
response.write "</table>"
response.write "</td>"
response.write "</tr>"
response.write "</table>"
response.write "</p>"
rs.close

'****************
'Emergency Calls to 911
'****************
objRS.Open "select Extension, Description, StartTime, TalkTime from Calldata " & _ 
  "where Inbound = False and PartyID = '9911' " & _
  "order by StartTime desc", objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Emergency Calls to 911</caption>"
Response.Write "<tr><th>Calling Number</th><th>Caller ID</th><th>Start Time</th><th>Talk Time</th></tr>"
Do While not objRS.EOF
  intHour = objRS("TalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("TalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("TalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & chr(34) & ">" & objRS("Description") & "</td><td>" & objRS("StartTime") & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "</table></p>"

'****************
'Emergency Calls to 2222
'****************
objRS.Open "select Extension, Description, StartTime, TalkTime from Calldata " & _ 
  "where Inbound = False and PartyID = '2222' " & _
  "order by StartTime desc", objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Emergency Calls to 2222</caption>"
Response.Write "<tr><th>Calling Number</th><th>Caller ID</th><th>Start Time</th><th>Talk Time</th></tr>"
Do While not objRS.EOF
  intHour = objRS("TalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("TalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("TalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & chr(34) & ">" & objRS("Description") & "</td><td>" & objRS("StartTime") & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "</table></p>"

'****************
'Elevator Calls from 5510 & 5511
'****************
objRS.Open "select Extension, Description, StartTime, TalkTime from Calldata " & _ 
  "where Internal = True and Inbound = False and (Extension = '5510' OR Extension = '5511') " & _
  "order by StartTime desc", objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Elevator Calls from 5510 & 5511</caption>"
Response.Write "<tr><th>Calling Number</th><th>Caller ID</th><th>Start Time</th><th>Talk Time</th></tr>"
Do While not objRS.EOF
  intHour = objRS("TalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("TalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("TalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & chr(34) & ">" & objRS("Description") & "</td><td>" & objRS("StartTime") & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "</table></p>"

'****************
'Information Line Calls to 1911
'****************
objRS.Open "select PartyID, PartyIDName, StartTime, EndTime, Internal from Calldata " & _ 
  "where Inbound = True and Extension = '1911' " & _
  "order by StartTime desc", objCalldata, adOpenStatic, adLockOptimistic'
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption><a href='\\mogsa4\DorIS\Processes\Emergency Information Line Script.doc' class='hidelink'>Information Line Calls to 1911</a></caption>"
Response.Write "<tr bgcolor='powderblue'><th>Calling Number</th><th>Caller ID</th><th>Start Time</th><th>Talk Time</th></tr>"
Do While not objRS.EOF
  if objRS("PartyID") = "1911" then
    strPartyID = objRS("Extension")
    strPartyIDName = objRS("Description")
  else
    strPartyID = objRS("PartyID")
    strPartyIDName = objRS("PartyIDName")
  end if
  if len(strPartyID) = 12 then
    strPartyID = "(" & mid(objRS("PartyID"),3,3) & ") " & mid(objRS("PartyID"),6,3) & "-" & mid(objRS("PartyID"),9,4)
  end if
  TalkTime = DateDiff("s", objRS("StartTime"), objRS("EndTime"))
  intHour = TalkTime \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (TalkTime mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (TalkTime mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'>" & strPartyID & "</td>"
  if objRS("Internal") = False then
    objRSdo.Open "SELECT PartyIDName FROM CallerID WHERE PartyID = '" & objRS("PartyID") & "'", objCalldata, adOpenStatic, adLockOptimistic
    if not objRSdo.EOF then
      if objRSdo("PartyIDName") <> " " then
        Response.Write "<td>" & objRSdo("PartyIDName") & "</td>"
      else
        Response.Write "<td>&nbsp</td>"
      end if
    else
      Response.Write "<td>&nbsp</td>"
    end if
    objRSdo.Close
  else
    Response.Write "<td>" & objRS("PartyIDName") & "</td>"
  end if
  Response.Write "<td>" & objRS("StartTime") & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "</table></p>"

'****************
'Most Frequently Called Numbers - Inbound
'****************
objRS.Open "select top " & Top & " Extension, Description, Count(StartTime) AS CountOfStartTime, Sum(TalkTime) AS SumOfTalkTime from Calldata " & _
  "where Internal = False and Inbound = True " & _
  "group by Extension, Description " & _
  "order by Count(StartTime) desc" _
, objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Top " & Top & " Most Frequently Called Numbers - Inbound</caption>"
Response.Write "<tr><th>Called Number</th><th>Description</th><th>Number of Calls</th><th>Talk Time</th></tr>"
Do While not objRS.EOF
  intHour = objRS("SumOfTalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("SumOfTalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("SumOfTalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & "&typ=ei" & chr(34) & ">" & objRS("Description") & "</td><td class='right'>" & formatnumber(objRS("CountOfStartTime"), 0) & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "</table></p>"

'****************
'Most Frequently Calling Numbers - Inbound
'****************
objRS.Open "select top " & Top & " PartyID, PartyIDName, Count(StartTime) AS CountOfStartTime, Sum(TalkTime) AS SumOfTalkTime from Calldata " & _
  "where Internal = False and Inbound = True  and (PartyID <> '' and PartyID <> 'O' and PartyID <> 'P') " & _
  "group by PartyID, PartyIDName " & _
  "order by Count(StartTime) desc" _
, objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Top " & Top & " Most Frequently Calling Numbers - Inbound</caption>"
Response.Write "<tr><th>Calling Number</th><th>Caller ID</th><th>Number of Calls</th><th>Talk Time</th></tr>"
Do While not objRS.EOF
  strPartyID = objRS("PartyID")
  if len(strPartyID) = 12 then
    strPartyID = "(" & mid(objRS("PartyID"),3,3) & ") " & mid(objRS("PartyID"),6,3) & "-" & mid(objRS("PartyID"),9,4)
  end if
  intHour = objRS("SumOfTalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("SumOfTalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("SumOfTalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'><a href=" & chr(34) & "detail.asp?ext=" &  objRS("PartyID") & "&typ=o" & chr(34) & ">" & strPartyID & "</td>"
  objRSdo.Open "select PartyIDName from CallerID where PartyID = '" & objRS("PartyID") & "'", objCalldata, adOpenStatic, adLockOptimistic
  if not objRSdo.EOF then
    if objRSdo("PartyIDName") <> " " then
      Response.Write "<td>" & objRSdo("PartyIDName") & "</td>"
    else
      Response.Write "<td>&nbsp</td>"
    end if
  else
    Response.Write "<td>&nbsp</td>"
  end if
  objRSdo.Close
  Response.Write "<td class='right'>" & formatnumber(objRS("CountOfStartTime"), 0) & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "</table></p>"

'****************
'Most Frequently Called Numbers - Outbound
'****************
objRS.Open "select top " & Top & " PartyID, Count(StartTime) as CountOfStartTime, Sum(TalkTime) as SumOfTalkTime from Calldata " & _
  "where Internal = False and Inbound = False " & _
  "group by PartyID " & _
  "order by Count(StartTime) desc" _
, objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Top " & Top & " Most Frequently Called Numbers - Outbound</caption>"
Response.Write "<tr><th>Called Number</th><th>Caller ID</th><th>Number of Calls</th><th>Talk Time</th></tr>"
Do While not objRS.EOF
  strPartyID = objRS("PartyID")
  if len(strPartyID) = 12 then
    strPartyID = "(" & mid(objRS("PartyID"),3,3) & ") " & mid(objRS("PartyID"),6,3) & "-" & mid(objRS("PartyID"),9,4)
  end if
  intHour = objRS("SumOfTalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("SumOfTalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("SumOfTalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'><a href=" & chr(34) & "detail.asp?ext=" &  objRS("PartyID") & "&typ=o" & chr(34) & ">" & strPartyID & "</td>"
  objRSdo.Open "select PartyIDName from CallerID where PartyID = '" & objRS("PartyID") & "'", objCalldata, adOpenStatic, adLockOptimistic
  if not objRSdo.EOF then
    if objRSdo("PartyIDName") <> " " then
      Response.Write "<td>" & objRSdo("PartyIDName") & "</td>"
    else
      Response.Write "<td>&nbsp</td>"
    end if
  else
    Response.Write "<td>&nbsp</td>"
  end if
  objRSdo.Close
  Response.Write "<td class='right'>" & formatnumber(objRS("CountOfStartTime"), 0) & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "</table></p>"

'****************
'Most Frequently Calling Numbers - Outbound
'****************
objRS.Open "select top " & Top & " Extension, Description, Count(StartTime) as CountOfStartTime, Sum(TalkTime) as SumOfTalkTime from Calldata " & _
  "where Internal = False and Inbound = False " & _
  "group by Extension, Description " & _
  "order by Count(StartTime) desc" _
, objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Top " & Top & " Most Frequently Calling Numbers - Outbound</caption>"
Response.Write "<tr><th>Calling Number</th><th>Description</th><th>Number of Calls</th><th>Talk Time</th></tr>"
Do While not objRS.EOF
  intHour = objRS("SumOfTalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("SumOfTalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("SumOfTalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & "&typ=eo" & chr(34) & ">" & objRS("Description") & "</td><td class='right'>" & formatnumber(objRS("CountOfStartTime"), 0) & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "</table></p>"

'****************
'Most Frequently Called Numbers - Internal
'****************
objRS.Open "select top " & Top & " Extension, Description, Count(StartTime) as CountOfStartTime, Sum(TalkTime) as SumOfTalkTime from Calldata " & _
  "where Internal = True and Inbound = True " & _
  "group by Extension, Description " & _
  "order by Count(StartTime) desc" _
, objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Top " & Top & " Most Frequently Called Numbers - Internal</caption>"
Response.Write "<tr><th>Called Number</th><th>Description</th><th>Number of Calls</th><th>Talk Time</th></tr>"
Do While not objRS.EOF
  intHour = objRS("SumOfTalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("SumOfTalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("SumOfTalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & "&typ=ii" & chr(34) & ">" & objRS("Description") & "</td><td class='right'>" & formatnumber(objRS("CountOfStartTime"), 0) & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "</table></p>"

'****************
'Most Frequently Calling Numbers - Internal
'****************
objRS.Open "select top " & Top & " Extension, Description, Count(StartTime) as CountOfStartTime, Sum(TalkTime) as SumOfTalkTime from Calldata " & _
  "where Internal = True and Inbound = False " & _
  "group by Extension, Description " & _
  "order by Count(StartTime) desc" _
, objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Top " & Top & " Most Frequently Calling Numbers - Internal</caption>"
Response.Write "<tr><th>Calling Number</th><th>Description</th><th>Number of Calls</th><th>Talk Time</th></tr>"
Do While not objRS.EOF
  intHour = objRS("SumOfTalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("SumOfTalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("SumOfTalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & "&typ=io" & chr(34) & ">" & objRS("Description") & "</td><td class='right'>" & formatnumber(objRS("CountOfStartTime"), 0) & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "</table></p>"

'****************
'Most Talk Time
'****************
objRS.Open "select top " & Top & " Extension, Description, Count(StartTime) as CountOfStartTime, Sum(TalkTime) as SumOfTalkTime from Calldata " & _
  "group by Extension, Description " & _
  "order by Sum(TalkTime) desc" _
, objCalldata, adOpenStatic, adLockOptimistic'
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Top " & Top & " Most Talk Time</caption>"
Response.Write "<tr><th>Extension</th><th>Description</th><th>Number of Calls</th><th>Talk Time</th></tr>"
Do While not objRS.EOF
  intHour = objRS("SumOfTalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("SumOfTalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("SumOfTalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & chr(34) & ">" & objRS("Description") & "</td><td class='right'>" & formatnumber(objRS("CountOfStartTime"), 0) & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "</table></p>"

'****************
'Longest Duration Calls
'****************
objRS.Open "select top " & Top & " Extension, Description, PartyID, PartyIDName, StartTime, TalkTime, Internal, Inbound from Calldata " & _
  "order by TalkTime desc" _
, objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Top " & Top & " Longest Duration Calls</caption>"
Response.Write "<tr><th>Extension</th><th>Description</th><th>Caller ID</th><th>Party</th><th>Start Time</th><th>Talk Time</th><th>Call Type</th></tr>"
do until objRS.eof
  strPartyID = objRS("PartyID")
  if len(strPartyID) = 12 then
    strPartyID = "(" & mid(objRS("PartyID"),3,3) & ") " & mid(objRS("PartyID"),6,3) & "-" & mid(objRS("PartyID"),9,4)
  end if
  intHour = objRS("TalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("TalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("TalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  if objRS("Internal") = True then
    strCallType = "Internal"
  elseif objRS("Inbound") = True then
    strCallType = "Inbound"
  elseif objRS("Inbound") = False then
    strCallType = "Outbound"
  end if
  Response.Write "<tr><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & chr(34) & ">" & objRS("Description") & "</td><td class='center'><a href=" & chr(34) & "detail.asp?ext=" &  objRS("PartyID") & "&typ=o" & chr(34) & ">" & strPartyID & "</td>"
  objRSdo.Open "select PartyIDName from CallerID where PartyID = '" & objRS("PartyID") & "'", objCalldata, adOpenStatic, adLockOptimistic
  if not objRSdo.EOF then
    if objRSdo("PartyIDName") <> " " then
      Response.Write "<td>" & objRSdo("PartyIDName") & "</td>"
    else
      Response.Write "<td>&nbsp</td>"
    end if
  else
    Response.Write "<td>&nbsp</td>"
  end if
  objRSdo.close
  Response.Write "<td>" & objRS("StartTime") & "</td><td class='right'>" & strTalkTime & "</td><td class='center'>" & strCallType & "</td></tr>"
  objRS.movenext
loop
objRS.Close
Response.Write "</table></p>"

'****************
'Longest Hold Times
'****************
objRS.Open "SELECT TOP " & Top & " * FROM Calldata " & _
  "WHERE Inbound = True AND Internal = False " & _
  "ORDER BY HoldTime DESC" _
, objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Top " & Top & " Longest Hold Times - Inbound Calls Only</caption>"
Response.Write "<tr><th>Extension</th><th>Description</th><th>Caller ID</th><th>Party</th><th>Start Time</th><th>Talk Time</th><th>Hold Time</th></tr>"
on error resume next
do until objRS.EOF
  strPartyID = objRS("PartyID")
  if len(strPartyID) = 12 then
    strPartyID = "(" & mid(objRS("PartyID"),3,3) & ") " & mid(objRS("PartyID"),6,3) & "-" & mid(objRS("PartyID"),9,4)
  elseif isnull(strPartyID) or strPartyID = " " or strPartyID = "" then
    strPartyID = "&nbsp"
  end if
  intHour = objRS("TalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("TalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("TalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec

  intHour = objRS("HoldTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("HoldTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("HoldTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strHoldTime = intHour & ":" & intMin & ":" & intSec

  if objRS("Internal") = True then
    strCallType = "Internal"
  elseif objRS("Inbound") = True then
    strCallType = "Inbound"
  elseif objRS("Inbound") = False then
    strCallType = "Outbound"
  end if
  Response.Write "<td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & chr(34) & ">" & objRS("Description") & "</td><td class='center'><a href=" & chr(34) & "detail.asp?ext=" &  objRS("PartyID") & "&typ=o" & chr(34) & ">" & strPartyID & "</td>"
  objRSdo.Open "SELECT PartyIDName FROM CallerID WHERE PartyID = '" & objRS("PartyID") & "'", objCalldata, adOpenStatic, adLockOptimistic
  if not objRSdo.EOF then
    if objRSdo("PartyIDName") <> " " then
      Response.Write "<td>" & objRSdo("PartyIDName") & "</td>"
    else
      Response.Write "<td>&nbsp</td>"
    end if
  else
    Response.Write "<td>&nbsp</td>"
  end if
  objRSdo.Close
  Response.Write "<td>" & objRS("StartTime") & "</td>"
  Response.Write "<td class='right'>" & strTalkTime & "</td>"
  Response.Write "<td class='right'>" & strHoldTime & "</td>"
  Response.Write "</tr>"
  objRS.movenext
loop
objRS.close
Response.Write "</table></p>"

'****************
'Longest Duration Long Distance Calls
'****************
objRS.open "select top " & Top & " Extension, Description, PartyID, PartyIDName, StartTime, TalkTime, Internal, Inbound from Calldata " & _
  "where LongDistance=True " & _
  "order by TalkTime desc" _
, objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Top " & Top & " Longest Duration Long Distance Calls</caption>"
Response.Write "<tr><th>Extension</th><th>Description</th><th>Caller ID</th><th>Party</th><th>Start Time</th><th>Talk Time</th></tr>"
do until objRS.eof
  strPartyID = objRS("PartyID")
  if len(strPartyID) = 12 then
    strPartyID = "(" & mid(objRS("PartyID"),3,3) & ") " & mid(objRS("PartyID"),6,3) & "-" & mid(objRS("PartyID"),9,4)
  end if
  intHour = objRS("TalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("TalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("TalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  if objRS("Internal") = True then
    strCallType = "Internal"
  elseif objRS("Inbound") = True then
    strCallType = "Inbound"
  elseif objRS("Inbound") = False then
    strCallType = "Outbound"
  end if
  Response.Write "<tr><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & chr(34) & ">" & objRS("Description") & "</td><td class='center'><a href=" & chr(34) & "detail.asp?ext=" &  objRS("PartyID") & "&typ=o" & chr(34) & ">" & strPartyID & "</td>"
  objRSdo.Open "select PartyIDName from CallerID where PartyID = '" & objRS("PartyID") & "'", objCalldata, adOpenStatic, adLockOptimistic
  if not objRSdo.eof then
    if objRSdo("PartyIDName") <> " " then
      Response.Write "<td>" & objRSdo("PartyIDName") & "</td>"
    else
      Response.Write "<td>&nbsp</td>"
    end if
  else
    Response.Write "<td>&nbsp</td>"
  end if
  objRSdo.close
  Response.Write "<td>" & objRS("StartTime") & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.movenext
loop
objRS.close
Response.Write "</table></p>"

'****************
'Most Frequently Calling Numbers - Long Distance
'****************
objRS.open "select top " & Top & " Extension, Description, Count(StartTime) as CountOfStartTime, Sum(TalkTime) as SumOfTalkTime from Calldata " & _
  "where Internal=False and Inbound=False and LongDistance=True " & _
  "group by Extension, Description " & _
  "order by Count(StartTime) desc" _
, objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Top " & Top & " Most Frequently Calling Numbers - Long Distance</caption>"
Response.Write "<tr><th>Calling Number</th><th>Description</th><th>Number of Calls</th><th>Talk Time</th></tr>"
do until objRS.eof
  intHour = objRS("SumOfTalkTime") \ 3600
  if intHour < 10 then
    intHour = "0" & intHour
  end if
  intMin = (objRS("SumOfTalkTime") mod 3600) \ 60
  if intMin < 10 then
    intMin = "0" & intMin
  end if
  intSec = (objRS("SumOfTalkTime") mod 3600) mod 60
  if intSec < 10 then
    intSec = "0" & intSec
  end if
  strTalkTime = intHour & ":" & intMin & ":" & intSec
  Response.Write "<tr><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & "&typ=eo" & chr(34) & ">" & objRS("Description") & "</td><td class='right'>" & formatnumber(objRS("CountOfStartTime"), 0) & "</td><td class='right'>" & strTalkTime & "</td></tr>"
  objRS.MoveNext
loop
objRS.close
Response.Write "</table></p>"

'****************
'System Statistics
'****************
objRS.open "select * from VoiceCost", objCalldata, adOpenStatic, adLockOptimistic
do until objRS.eof
  'January
  if objRS("Mth1") = 0 then
    Mth1Total = Mth1Total + objRS("Budget") / 12
  else
    Mth1Total = Mth1Total + objRS("Mth1")
  end if
  'February
  if objRS("Mth2") = 0 then
    Mth2Total = Mth2Total + objRS("Budget") / 12
  else
    Mth2Total = Mth2Total + objRS("Mth2")
  end if
  'March
  if objRS("Mth3") = 0 then
    Mth3Total = Mth3Total + objRS("Budget") / 12
  else
    Mth3Total = Mth3Total + objRS("Mth3")
  end if
  'April
  if objRS("Mth4") = 0 then
    Mth4Total = Mth4Total + objRS("Budget") / 12
  else
    Mth4Total = Mth4Total + objRS("Mth4")
  end if
  'May
  if objRS("Mth5") = 0 then
    Mth5Total = Mth5Total + objRS("Budget") / 12
  else
    Mth5Total = Mth5Total + objRS("Mth5")
  end if
  'June
  if objRS("Mth6") = 0 then
    Mth6Total = Mth6Total + objRS("Budget") / 12
  else
    Mth6Total = Mth6Total + objRS("Mth6")
  end if
  'July
  if objRS("Mth7") = 0 then
    Mth7Total = Mth7Total + objRS("Budget") / 12
  else
    Mth7Total = Mth7Total + objRS("Mth7")
  end if
  'August
  if objRS("Mth8") = 0 then
    Mth8Total = Mth8Total + objRS("Budget") / 12
  else
    Mth8Total = Mth8Total + objRS("Mth8")
  end if
  'September
  if objRS("Mth9") = 0 then
    Mth9Total = Mth9Total + objRS("Budget") / 12
  else
    Mth9Total = Mth9Total + objRS("Mth9")
  end if
  'October
  if objRS("Mth10") = 0 then
    Mth10Total = Mth10Total + objRS("Budget") / 12
  else
    Mth10Total = Mth10Total + objRS("Mth10")
  end if
  'November
  if objRS("Mth11") = 0 then
    Mth11Total = Mth11Total + objRS("Budget") / 12
  else
    Mth11Total = Mth11Total + objRS("Mth11")
  end if
  'December
  if objRS("Mth12") = 0 then
    Mth12Total = Mth12Total + objRS("Budget") / 12
  else
    Mth12Total = Mth12Total + objRS("Mth12")
  end if
  objRS.movenext
loop
CostOfSystem = Mth1Total + Mth2Total + Mth3Total + Mth4Total + Mth5Total + Mth6Total + Mth7Total + Mth8Total + Mth9Total + Mth10Total + Mth11Total + Mth12Total
objRS.close

'Determine the total number of talk time seconds for the past 30 days
objRS.open "select Sum(TalkTime) as SumOfTalkTime from Calldata", objCalldata, adOpenStatic, adLockOptimistic
SumOfTalkTime = objRS("SumOfTalkTime")
objRS.close

'Determine the cost per second per month
CostPerSec = ((CostOfSystem / 12) / (SumOfTalkTime)) * 1.014583

'Determine the Total numbers of Extensions in the System
objRS.open "select Count(ExtensionListItems.UserDN) as CountOfUserDN from ExtensionListItems where ExtensionListID<50", objCfg, adOpenStatic, adLockOptimistic
SumOfSystemExt = objRS("CountOfUserDN")
objRS.close

'Determine the cost per extension per month
CostPerExt = ((CostOfSystem / 12) / (SumOfSystemExt)) * 1.014583

'****************
'Give paging groups an ExtensionListID > 50
'****************
objRS.open "select ExtensionLists.ExtensionListID, ExtensionLists.Name, Count(ExtensionListItems.UserDN) as CountOfUserDN from ExtensionLists " & _
  "inner join (select ExtensionListItems.UserDN, ExtensionListItems.ExtensionListID from ExtensionListItems) " & _
    "as ExtensionListItems on ExtensionLists.ExtensionListID=ExtensionListItems.ExtensionListID where ExtensionLists.ExtensionListID<50 " & _
  "group by ExtensionLists.ExtensionListID, ExtensionLists.Name " & _
  "order by ExtensionLists.Name" _
, objCfg, adOpenStatic, adLockOptimistic
''****************
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Statistics</caption>"
Response.Write "<tr><th colspan='5'>Monthly accrual is based on 50% usage and 50% cost per extension</th></tr>"
Response.Write "<tr><td>Cost of the Phone System per Month</td><td><a href='phonesystemcost.asp'>" & FormatCurrency(CostOfSystem / 12, 2) & "</a></td></tr>"
'<a href=" & chr(34) & "cc.asp?cc=" &  objRS("ExtensionListID") & chr(34) & ">
Response.Write "<tr><td>Sum of Talk Time for the Past 30 Days</td><td>" & FormatNumber(SumOfTalkTime / 60, 0) & " Minutes</td></tr>"
Response.Write "<tr><td>Cost of Talk Time per Minute (Annualized)</td><td>" & FormatCurrency((CostPerSec * 60) / 2, 4) & "</td></tr>"
Response.Write "<tr><td>Total Number of Extensions in System</td><td>" & FormatNumber(SumOfSystemExt, 0) & "</td></tr>"
Response.Write "<tr><td>Cost of an Extension per Month (Annualized)</td><td>" & FormatCurrency(CostPerExt / 2, 2) & "</td></tr>"
Response.Write "</table></p>"

'****************
'Cost Center Overview
'****************
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Cost Center Overview</caption>"
Response.Write "<tr><th>Cost Center</th><th>Number of Extensions</th><th>Number of Calls</th><th>Total Talk Time</th><th>Monthly Accrual</th><th>Pct of Total</th><th>Standard</th></tr>"
do until objRS.eof
  objRSdo.Open "select ExtensionLists.Name, ExtensionListItems.UserDN from ExtensionListItems " & _
    "inner join (select ExtensionLists.Name, ExtensionLists.ExtensionListID from ExtensionLists) " & _
      "as ExtensionLists on ExtensionLists.ExtensionListID=ExtensionListItems.ExtensionListID " & _
    "where ExtensionListItems.ExtensionListID=" & objRS("ExtensionListID") & " " & _
    "order by ExtensionLists.Name" _
  , objCfg, adOpenStatic, adLockOptimistic
  intStartTime = 0
  intTalkTime = 0
  do until objRSdo.eof
    objRSdodo.open "select Count(StartTime) as CountOfStartTime, Sum(TalkTime) as SumOfTalkTime from Calldata " & _
      "where Extension='" & objRSdo("UserDN") & "'" _
    , objCalldata, adOpenStatic, adLockOptimistic
    if not objRSdodo.eof then
      if objRSdodo("CountOfStartTime") > 0 then
        intStartTime = intStartTime + objRSdodo("CountOfStartTime")
      end if
      if objRSdodo("SumOfTalkTime") > 0 then
        intTalkTime = intTalkTime + objRSdodo("SumOfTalkTime")
      end if
    end if
    objRSdodo.close
    objRSdo.movenext
  loop
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
  Response.Write "<tr>"
  Response.Write "<td><a href=" & chr(34) & "cc.asp?cc=" &  objRS("ExtensionListID") & chr(34) & ">" & objRS("Name") & "</a></td>"
  Response.Write "<td class='right'>" & objRS("CountOfUserDN") & "</td>"
  Response.Write "<td class='right'>" & FormatNumber(intStartTime, 0) & "</td>"
  Response.Write "<td class='right'>" & strTalkTime & "</td>"
  Response.Write "<td class='right'>" & FormatCurrency(Ccur((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / 2, 2) & "</td>"
  Response.Write "<td class='right'>" & FormatPercent((((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2, 1) & "</td>"
  select case objRS("ExtensionListID")
    case 3
      if poly = (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2 then
        Response.Write "<td class='right'>" & FormatPercent(poly, 1) & "</td>"
      else
        pct_diff = poly - (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2
        if pct_diff > 0 then
          cr = "FF"
          cg = hex(255 - (abs(pct_diff) * 10000))
          cb = hex(255 - (abs(pct_diff) * 10000))
        else
          cr = hex(255 - (abs(pct_diff) * 10000))
          cg = "FF"
          cb = hex(255 - (abs(pct_diff) * 10000))
        end if
        Response.Write "<td class='right' style='background-color:#" & cr & cg & cb & "'>" & FormatPercent(poly, 1) & "</style></td>"
      end if
    case 4
      if spin = (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2 then
        Response.Write "<td class='right'>" & FormatPercent(spin, 1) & "</td>"
      else
        pct_diff = spin - (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2
        if pct_diff > 0 then
          cr = "FF"
          cg = hex(255 - (abs(pct_diff) * 10000))
          cb = hex(255 - (abs(pct_diff) * 10000))
        else
          cr = hex(255 - (abs(pct_diff) * 10000))
          cg = "FF"
          cb = hex(255 - (abs(pct_diff) * 10000))
        end if
        Response.Write "<td class='right' style='background-color:#" & cr & cg & cb & "'>" & FormatPercent(spin, 1) & "</td>"
      end if
    case 7
      if warp = (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2 then
        Response.Write "<td class='right'>" & FormatPercent(warp, 1) & "</td>"
      else
        pct_diff = warp - (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2
        if pct_diff > 0 then
          cr = "FF"
          cg = hex(255 - (abs(pct_diff) * 10000))
          cb = hex(255 - (abs(pct_diff) * 10000))
        else
          cr = hex(255 - (abs(pct_diff) * 10000))
          cg = "FF"
          cb = hex(255 - (abs(pct_diff) * 10000))
        end if
        Response.Write "<td class='right' style='background-color:#" & cr & cg & cb & "'>" & FormatPercent(warp, 1) & "</td>"
      end if
    case 8
      if corp = (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2 then
        Response.Write "<td class='right'>" & FormatPercent(corp, 1) & "</td>"
      else
        pct_diff = corp - (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2
        if pct_diff > 0 then
          cr = "FF"
          cg = hex(255 - (abs(pct_diff) * 10000))
          cb = hex(255 - (abs(pct_diff) * 10000))
        else
          cr = hex(255 - (abs(pct_diff) * 10000))
          cg = "FF"
          cb = hex(255 - (abs(pct_diff) * 10000))
        end if
        Response.Write "<td class='right' style='background-color:#" & cr & cg & cb & "'>" & FormatPercent(corp, 1) & "</td>"
      end if
    case 10
      if admi = (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2 then
        Response.Write "<td class='right'>" & FormatPercent(admi, 1) & "</td>"
      else
        pct_diff = admi - (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2
        if pct_diff > 0 then
          cr = "FF"
          cg = hex(255 - (abs(pct_diff) * 10000))
          cb = hex(255 - (abs(pct_diff) * 10000))
        else
          cr = hex(255 - (abs(pct_diff) * 10000))
          cg = "FF"
          cb = hex(255 - (abs(pct_diff) * 10000))
        end if
        Response.Write "<td class='right' style='background-color:#" & cr & cg & cb & "'>" & FormatPercent(admi, 1) & "</td>"
      end if
    case 11
      if prod = (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2 then
        Response.Write "<td class='right'>" & FormatPercent(prod, 1) & "</td>"
      else
        pct_diff = prod - (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2
        if pct_diff > 0 then
          cr = "FF"
          cg = hex(255 - (abs(pct_diff) * 10000))
          cb = hex(255 - (abs(pct_diff) * 10000))
        else
          cr = hex(255 - (abs(pct_diff) * 10000))
          cg = "FF"
          cb = hex(255 - (abs(pct_diff) * 10000))
        end if
        Response.Write "<td class='right' style='background-color:#" & cr & cg & cb & "'>" & FormatPercent(prod, 1) & "</td>"
      end if
    case 15
      if dist = (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2 then
        Response.Write "<td class='right'>" & FormatPercent(dist, 1) & "</td>"
      else
        pct_diff = dist - (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2
        if pct_diff > 0 then
          cr = "FF"
          cg = hex(255 - (abs(pct_diff) * 10000))
          cb = hex(255 - (abs(pct_diff) * 10000))
        else
          cr = hex(255 - (abs(pct_diff) * 10000))
          cg = "FF"
          cb = hex(255 - (abs(pct_diff) * 10000))
        end if
        Response.Write "<td class='right' style='background-color:#" & cr & cg & cb & "'>" & FormatPercent(dist, 1) & "</td>"
      end if
    case 14
      if mark = (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2 then
        Response.Write "<td class='right'>" & FormatPercent(mark, 1) & "</td>"
      else
        pct_diff = mark - (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2
        if pct_diff > 0 then
          cr = "FF"
          cg = hex(255 - (abs(pct_diff) * 10000))
          cb = hex(255 - (abs(pct_diff) * 10000))
        else
          cr = hex(255 - (abs(pct_diff) * 10000))
          cg = "FF"
          cb = hex(255 - (abs(pct_diff) * 10000))
        end if
        Response.Write "<td class='right' style='background-color:#" & cr & cg & cb & "'>" & FormatPercent(mark, 1) & "</td>"
      end if
    case 16
      if qual = (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2 then
        Response.Write "<td class='right'>" & FormatPercent(qual, 1) & "</td>"
      else
        pct_diff = qual - (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2
        if pct_diff > 0 then
          cr = "FF"
          cg = hex(255 - (abs(pct_diff) * 10000))
          cb = hex(255 - (abs(pct_diff) * 10000))
        else
          cr = hex(255 - (abs(pct_diff) * 10000))
          cg = "FF"
          cb = hex(255 - (abs(pct_diff) * 10000))
        end if
        Response.Write "<td class='right' style='background-color:#" & cr & cg & cb & "'>" & FormatPercent(qual, 1) & "</td>"
      end if
    case 17
      if main = (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2 then
        Response.Write "<td class='right'>" & FormatPercent(main, 1) & "</td>"
      else
        pct_diff = main - (((intTalkTime * CostPerSec) + (objRS("CountOfUserDN") * CostPerExt)) / (CostOfSystem / 12)) / 2
        if pct_diff > 0 then
          cr = "FF"
          cg = hex(255 - (abs(pct_diff) * 10000))
          cb = hex(255 - (abs(pct_diff) * 10000))
        else
          cr = hex(255 - (abs(pct_diff) * 10000))
          cg = "FF"
          cb = hex(255 - (abs(pct_diff) * 10000))
        end if
        Response.Write "<td class='right' style='background-color:#" & cr & cg & cb & "'>" & FormatPercent(main, 1) & "</td>"
      end if
  end select
  Response.Write "</tr>"
  objRSdo.close
  objRS.movenext
loop
objRS.close
Response.Write "</table></p>"
%>

</font>
</p>
</body>
</html>
