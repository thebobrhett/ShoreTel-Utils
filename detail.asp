<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>AKSA Phone System Detail Report</title>
<link href='phone.css' rel='stylesheet' type='text/css'>
</head>
<body>

<%
'on error resume next
dim Past
dim ext
dim typ
dim NightColor
dim objCall
dim objCfg
dim objCalldata
dim objRS
dim objRSdo
dim objRSdodo

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3
Const sel_conn = "Connect.PartyType, Connect.ConnectReason"
Const int_all = "Call.CallType = 1 AND Connect.PartyType = 1 AND Connect.ConnectReason = 17"
Const ext_in = "Call.CallType = 2 AND Connect.PartyType = 2 AND Connect.ConnectReason = 19"
Const ext_out = "Call.CallType = 3 AND Connect.PartyType = 2 AND Connect.ConnectReason = 17"

Past = Request("past")
if Past = 0 then
  Past = 30
end if
Response.Write "<h1>AKSA Phone System Reports - Past " & Past & " Days<br />"
ext = Request("ext")
if len(ext) <> 4 and len(ext) <> 11 then
  ext = "+" & right(Request("ext"), len(Request("ext")) - 1)
end if
typ = Request("typ")

set objCfg = CreateObject("adodb.connection")
'objCfg.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Shoreline Data\Database\ShoreWare.mdb"
objCfg.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=omaha;database=shoreware;port=4308;user=st_configread; pwd=passwordconfigread;"

set objCalldata = CreateObject("adodb.connection")
objCalldata.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\Assetmgt\Phone.mdb"

'set objCall = CreateObject("adodb.connection")
'objCall.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Shoreline Data\Call Records 2\CDR.mdb"
'otherdb = "D:\Shoreline Data\Database\ShoreWare.mdb"

set objRS = CreateObject("adodb.recordset")
set objRSdo = CreateObject("adodb.recordset")
set objRSdodo = CreateObject("adodb.recordset")

objRS.CursorLocation = adUseClient
objRSdo.CursorLocation = adUseClient
objRSdodo.CursorLocation = adUseClient

if typ <> "o" then
  objRS.open "select Description from DN where DN='" & ext & "'", objCfg, adOpenStatic, adLockOptimistic
  Response.Write "Activity Detail for Extension " & ext & " - " & objRS("Description") & "</h1>"
  objRS.close

  Response.Write "<hr/>"
  Response.Write "<p>"
  Response.Write "<table border='1'>"
  Response.Write "<caption>Call Detail"
  if typ = "ei" then
    Response.Write " - Inbound Calls Received Only"
  elseif typ = "eo" then
    Response.Write " - Outbound Calls Placed Only"
  elseif typ = "ii" then
    Response.Write " - Internal Calls Received Only"
  elseif typ = "io" then
    Response.Write " - Internal Calls Placed Only"
  end if
  Response.Write "</caption>"
  Response.Write "<tr><th>Start Time</th><th>Activity</th><th>Caller ID</th><th>Party</th><th><a href=top10.asp?ext=" &  ext & ">Talk Time</a></th></tr>"

  objRS.Open "select Extension, Description, PartyID, PartyIDName, StartTime, TalkTime, Internal, Inbound from Calldata " & _
    "where Extension='" & ext & "' " & _
    "order by StartTime desc" _
  , objCalldata, adOpenStatic, adLockOptimistic

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
    if objRS("Internal") = True and objRS("Inbound") = True and (typ = "" or typ = "ii")  then
      'Internal call received
      CurrentTime = TimeValue(DatePart("h", objRS("StartTime")) & ":" & DatePart("n", objRS("StartTime")) & ":" & DatePart("s", objRS("StartTime")))
      if CurrentTime < TimeValue("6:00:00") or CurrentTime > TimeValue("18:00:00") then
        Response.Write "<tr bgcolor= '" & NightColor & "'>"
      else
        Response.Write "<tr>"
      end if
      Response.Write "<td>" & objRS("StartTime") & "</td><td> Internal call received from </td>"
      if strPartyID = ext then
        Response.Write "<td>Voicemail System</td><td>Message Notification</td>"
      else
        Response.Write "<td class='center'>" & strPartyID & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("PartyID") & chr(34) & ">" & objRS("PartyIDName") & "</td>"
      end if
      Response.Write "<td class='right'>" & strTalkTime & "</td></tr>"
    elseif objRS("Internal") = True and objRS("Inbound") = False and (typ = "" or typ = "io")  then
      'Internal call placed
      CurrentTime = TimeValue(DatePart("h", objRS("StartTime")) & ":" & DatePart("n", objRS("StartTime")) & ":" & DatePart("s", objRS("StartTime")))
      if CurrentTime < TimeValue("6:00:00") or CurrentTime > TimeValue("18:00:00") then
        Response.Write "<tr class='night'>"
      else
        Response.Write "<tr>"
      end if
      Response.Write "<td>" & objRS("StartTime") & "</td><td> Internal call placed to </td>"
      if strPartyID = ext then
        Response.Write "<td>Voicemail System</td><td>Message Retreival</td>"
      else
        Response.Write "<td class='center'>" & strPartyID & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("PartyID") & chr(34) & ">" & objRS("PartyIDName") & "</td>"
      end if
      Response.Write "<td class='right'>" & strTalkTime & "</td></tr>"
    elseif objRS("Internal") = False and objRS("Inbound") = True and (typ = "" or typ = "ei") then
      'External call received
      CurrentTime = TimeValue(DatePart("h", objRS("StartTime")) & ":" & DatePart("n", objRS("StartTime")) & ":" & DatePart("s", objRS("StartTime")))
      if CurrentTime < TimeValue("6:00:00") or CurrentTime > TimeValue("18:00:00") then
        Response.Write "<tr class='night'>"
      else
        Response.Write "<tr>"
      end if
      Response.Write "<td>" & objRS("StartTime") & "</td><td> External call received from </td>"
      if isnull(objRS("PartyID")) or objRS("PartyID") = "" then
        Response.Write "<td>&nbsp</td>"
      else
        Response.Write "<td class='center'><a href=" & chr(34) & "detail.asp?ext=" &  objRS("PartyID") & "&typ=o" & chr(34) & ">" & strPartyID & "</td>"
      end if
      objRSdo.open "select PartyIDName from CallerID where PartyID='" & objRS("PartyID") & "'", objCalldata, adOpenStatic, adLockOptimistic
      if not objRSdo.eof then
        if objRSdo("PartyIDName") <> " " then
          Response.Write "<td>" & objRSdo("PartyIDName") & "</td>"
        else
          Response.Write "<td>&nbsp</td>"
        end if
      end if
      objRSdo.close

      Response.Write "<td class='right'>" & strTalkTime & "</td></tr>"
    elseif objRS("Internal") = False and objRS("Inbound") = False and (typ = "" or typ = "eo")  then
      'External call placed
      CurrentTime = TimeValue(DatePart("h", objRS("StartTime")) & ":" & DatePart("n", objRS("StartTime")) & ":" & DatePart("s", objRS("StartTime")))
      if CurrentTime < TimeValue("6:00:00") or CurrentTime > TimeValue("18:00:00") then
        Response.Write "<tr class='night'>"
      else
        Response.Write "<tr>"
      end if
      Response.Write "<td>" & objRS("StartTime") & "</td><td> External call placed to </td>"
      if isnull(objRS("PartyID")) or objRS("PartyID") = "" or objRS("PartyID") = " " then
        Response.Write "<td>&nbsp</td>"
      else
        Response.Write "<td class='center'><a href=" & chr(34) & "detail.asp?ext=" &  objRS("PartyID") & "&typ=o" & chr(34) & ">" & strPartyID & "</td>"
      end if
      objRSdo.open "select PartyIDName from CallerID where PartyID='" & objRS("PartyID") & "'", objCalldata, adOpenStatic, adLockOptimistic
      if not objRSdo.eof then
        if objRSdo("PartyIDName") <> " " then
          Response.Write "<td>" & objRSdo("PartyIDName") & "</td>"
        else
          Response.Write "<td>&nbsp</td>"
        end if
      end if
      objRSdo.close
      Response.Write "<td class='right'>" & strTalkTime & "</td></tr>"
    end if
    objRS.movenext
  loop
else
  strPartyID = ext
  if len(strPartyID) = 12 then
    strPartyID = "(" & mid(strPartyID,3,3) & ") " & mid(strPartyID,6,3) & "-" & mid(strPartyID,9,4)
  end if
  Response.Write "<h1>Activity Detail for Number " & strPartyID & " - "
  objRS.open "select PartyIDName, Comment from CallerID where PartyID='" & ext & "'", objCalldata, adOpenStatic, adLockOptimistic
  if not objRS.eof then
    if isnull(objRS("PartyIDName")) or objRS("PartyIDName") = "" or objRS("PartyIDName") = " " then
      Response.Write "unknown"
      if not isnull(objRS("Comment")) then
        Response.Write "<br />" & objRS("Comment")
      end if
    else
      Response.Write objRS("PartyIDName")
      if not isnull(objRS("Comment")) then
        Response.Write "<br />" & objRS("Comment")
      end if
    end if
  end if
  Response.Write "</h1>"
  objRS.Close

  Response.Write "<hr/>"
  Response.Write "<p>"
  Response.Write "<table border='1'>"
  Response.Write "<caption>Call Detail - External Calls</caption>"
  Response.Write "<tr><th>Start Time</th><th>Activity</th><th>Extension</th><th>Description</th><th><a href=top10.asp?ext=" &  ext & "&typ=" & typ & ">Talk Time</a></th></tr>"

  objRS.Open "select Extension, Description, PartyID, PartyIDName, StartTime, TalkTime, Internal, Inbound from Calldata " & _
    "where PartyID='" & ext & "' " & _
    "order by StartTime desc" _
  , objCalldata, adOpenStatic, adLockOptimistic

  do until objRS.eof
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
    if objRS("Internal") = False and objRS("Inbound") = True then
      'External call received
      CurrentTime = TimeValue(DatePart("h", objRS("StartTime")) & ":" & DatePart("n", objRS("StartTime")) & ":" & DatePart("s", objRS("StartTime")))
      if CurrentTime < TimeValue("6:00:00") or CurrentTime > TimeValue("18:00:00") then
        Response.Write "<tr class='night'>"
      else
        Response.Write "<tr>"
      end if
      Response.Write "<td>" & objRS("StartTime") & "</td><td> External call received by </td><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & chr(34) & ">" & objRS("Description") & "</td><td class='right'>" & strTalkTime & "</td></tr>"
    elseif objRS("Internal") = False and objRS("Inbound") = False then
      'External call placed
      CurrentTime = TimeValue(DatePart("h", objRS("StartTime")) & ":" & DatePart("n", objRS("StartTime")) & ":" & DatePart("s", objRS("StartTime")))
      if CurrentTime < TimeValue("6:00:00") or CurrentTime > TimeValue("18:00:00") then
        Response.Write "<tr class='night'>"
      else
        Response.Write "<tr>"
      end if
      Response.Write "<td>" & objRS("StartTime") & "</td><td> External call placed by </td><td class='center'>" & objRS("Extension") & "</td><td><a href=" & chr(34) & "detail.asp?ext=" &  objRS("Extension") & chr(34) & ">" & objRS("Description") & "</td><td class='right'>" & strTalkTime & "</td></tr>"
    end if

    objRS.movenext
  loop

end if

objRS.close
Response.Write "</table></p>"
%>

</font>
</p>
</body>
</html>
