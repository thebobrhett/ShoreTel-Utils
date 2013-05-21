'****************
'Bob Rhett - Monday, November 1, 2009
'  Modified for MySQL and Shoreware version 7.5
'Keith Brooks - Monday, October 3, 2011
'  Modified for Shoreware configuration database move to MySQL after upgrade to
'  Shoreware version 11.2
'****************
'on error resume next

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3

dim Test
dim Past
dim Start
dim LastTime
dim objCall
dim objCfg
dim objCalldata
dim objRSCallData
dim objRSCall
dim objRSConnect
dim objRSDN
dim objRSdo
dim Description
dim PartyID
dim PartyIDName
dim Hits
dim TalkTime
dim HoldTime

Test = False
'Top = 5
Past = 30
Start = DateAdd("d", 0 - Past, Date)

set objCall = CreateObject("adodb.connection")
set objCfg = CreateObject("adodb.connection")
set objCalldata = CreateObject("adodb.connection")

objCall.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=Omaha;database=shorewarecdr;port=4309;user=st_cdrreport;password=passwordcdrreport;"
'objCfg.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Shoreline Data\Database\ShoreWare.mdb"
objCfg.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=omaha;database=shoreware;port=4308;user=st_configread; pwd=passwordconfigread;"
objCalldata.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\Assetmgt\Phone.mdb"

set objRS = CreateObject("adodb.recordset")
set objRSdo = CreateObject("adodb.recordset")
set objRSCalldata = CreateObject("adodb.recordset")
set objRSCall = CreateObject("adodb.recordset")
set objRSConnect = CreateObject("adodb.recordset")
set objRSDN = CreateObject("adodb.recordset")

objRSCalldata.CursorLocation = adUseClient

if Test = False then

strSQL = "delete from Calldata where EndTime < #" & Start & "#"
objRSCalldata.open strSQL, objCalldata, adOpenStatic, adLockOptimistic
strSQL = "select * FROM Calldata order by EndTime desc"
objRSCalldata.open strSQL, objCalldata, adOpenStatic, adLockOptimistic

LastTime = Start
if not objRSCalldata.eof then
  LastTime = objRSCalldata("EndTime")
end if
LastTime = cstr(year(LastTime)) & "-" & cstr(month(LastTime)) & "-" & cstr(day(LastTime)) & " " & cstr(hour(LastTime)) & ":" & cstr(minute(LastTime)) & ":" & cstr(second(LastTime))

'****************
'Gather External, Inbound data
'****************
strSQL = "select ID, Extension, StartTime, EndTime from `Call` where CallType=2 and EndTime > '" & LastTime & "'"
objRSCall.open strSQL, objCall
do until objRSCall.eof
  strSQL = "select PartyID, PartyIDName, PartyIDLastName, sum(TalkTimeSeconds), sum(HoldTime), PartyType from `Connect` where (ConnectReason=9 or ConnectReason=19) and CallTableID=" & objRSCall("ID") & " group by CallTableID"
  objRSConnect.open strSQL, objCall
  strSQL = "select Description from DN where DN='" & objRSCall("Extension") & "'"
  objRSDN.open strSQL, objCfg
  'post info to Phone.mdb
  PartyID = objRSConnect("PartyID")
  if left(PartyID, 2) <> "+1" then
    if left(PartyID, 1) = "+" then
      PartyID = "+1" & right(PartyID, len(PartyID) - 1)
    else
      PartyID = "+1" & PartyID
    end if
  end if
  PartyIDName = objRSConnect("PartyIDName") & " " & objRSConnect("PartyIDLastName")
'  PartyIDName = replace(PartyIDName, "\", "\\")
'  PartyIDName = replace(PartyIDName, "'", "\'")
  PartyIDName = replace(PartyIDName, "'", "")
  wscript.echo objRSCall("ID") & " " & objRSCall("Extension") & " " & objRSCall("StartTime") & " " & objRSCall("EndTime")
  if not objRSDN.eof then
    wscript.echo "  Description:" & objRSDN("Description")
    Description = objRSDN("Description")
  else
    wscript.echo "  Description: (none)"
    Description = ""
  end if
  wscript.echo "  PartyID:" & PartyID
  wscript.echo "  PartyIDName:" & PartyIDName
  wscript.echo "  TalkTimeSeconds:" & objRSConnect("sum(TalkTimeSeconds)")
'****************
  HoldTime = objRSConnect("sum(HoldTime)")
  do until HoldTime < 18991230000000
    HoldTime = HoldTime - 18991230000000
  loop
  wscript.echo "  HoldTime:" & HoldTime
'****************
  wscript.echo "  PartyType:" & objRSConnect("PartyType")
  wscript.echo
  objRSdo.Open "select * from CallerID where PartyID = '" & PartyID & "'", objCallData, adOpenStatic, adLockOptimistic
  if objRSdo.eof then
    Hits = 1
    strSQL = "insert into CallerID values ('" & PartyID & "', '" & PartyIDName & "', False, '" & Date() & "', '', " & Hits & ")"
  else
    Hits = objRSdo("HitCount") + 1
    if not isnull(objRSdo("PartyIDName")) or objRSdo("PartyIDName") = "" or objRSdo("PartyIDName") = " " then
      PartyIDName = objRSdo("PartyIDName")
'      PartyIDName = replace(PartyIDName, "\", "\\")
'      PartyIDName = replace(PartyIDName, "'", "\'")
      PartyIDName = replace(PartyIDName, "'", "")
    end if
    strSQL = "update CallerID set PartyIDName='" & PartyIDName & "', Hit='" & Date() & "', HitCount=" & Hits & " where PartyID='" & PartyID & "'"
  end if
  wscript.echo strSQL
  objRS.open strSQL, objCalldata, adOpenStatic, adLockOptimistic
  objRSdo.Close
  strSQL = "select sum(TalkTimeSeconds), sum(HoldTime) from `Connect` where (ConnectReason=9 or ConnectReason=19) and CallTableID=" & objRSCall("ID") & " group by CallTableID"
  objRSdo.Open strSQL, objCall
  if not objRSdo.eof then
    TalkTime = objRSdo("sum(TalkTimeSeconds)")
'    on error resume next
'****************
'    HoldTime = 0
    HoldTime = objRSdo("sum(HoldTime)")
    do until HoldTime < 18991230000000
      HoldTime = HoldTime - 18991230000000
    loop
'****************
    on error goto 0
  else
    TalkTime = objRSConnect("sum(TalkTimeSeconds)")
    HoldTime = "0:00:00"
  end if
  objRSdo.Close
  strSQL = "insert into Calldata values ('" & objRSCall("Extension") & "', '" & Description & "', '" & PartyID & "', '" & PartyIDName & "', False, True, '" & objRSCall("StartTime") & "', '" & objRSCall("EndTime") & "', '" & TalkTime & "', '" & HoldTime & "', False)"
  wscript.echo strSQL
  objRSdo.open strSQL, objCalldata, adOpenStatic, adLockOptimistic
  objRSDN.close
  objRSConnect.close
  objRSCall.movenext
loop
objRSCall.close

'****************
'Gather External, Outbound data
'****************
strSQL = "select ID, Extension, StartTime, EndTime from `Call` where CallType=3 and EndTime > '" & LastTime & "'"
objRSCall.open strSQL, objCall
do until objRSCall.eof
  strSQL = "select PartyID, PartyIDName, PartyIDLastName, sum(TalkTimeSeconds), sum(HoldTime), PartyType, LongDistance from `Connect` where ConnectReason=17 and CallTableID=" & objRSCall("ID") & " group by CallTableID"
  objRSConnect.open strSQL, objCall
  strSQL = "select Description from DN where DN='" & objRSCall("Extension") & "'"
  objRSDN.open strSQL, objCfg
  'post info to Phone.mdb
  PartyID = objRSConnect("PartyID")
  if left(PartyID, 2) = "9+" then
    PartyID = right(PartyID, len(PartyID) - 1)
  end if
  PartyIDName = objRSConnect("PartyIDName") & " " & objRSConnect("PartyIDLastName")
'  PartyIDName = replace(PartyIDName, "\", "\\")
'  PartyIDName = replace(PartyIDName, "'", "\'")
  PartyIDName = replace(PartyIDName, "'", "")
  wscript.echo objRSCall("ID") & " " & objRSCall("Extension") & " " & objRSCall("StartTime") & " " & objRSCall("EndTime")
  if not objRSDN.eof then
    wscript.echo "  Description:" & objRSDN("Description")
    Description = objRSDN("Description")
  else
    wscript.echo "  Description: (none)"
    Description = ""
  end if
  wscript.echo "  PartyID:" & PartyID
  wscript.echo "  PartyIDName:" & PartyIDName
  wscript.echo "  TalkTimeSeconds:" & objRSConnect("sum(TalkTimeSeconds)")
'****************
  HoldTime = objRSConnect("sum(HoldTime)")
  do until HoldTime < 18991230000000
    HoldTime = HoldTime - 18991230000000
  loop
  wscript.echo "  HoldTime:" & HoldTime
'****************
  wscript.echo "  PartyType:" & objRSConnect("PartyType")
  wscript.echo
  objRSdo.Open "select * from CallerID where PartyID = '" & PartyID & "'", objCallData, adOpenStatic, adLockOptimistic
  if objRSdo.eof then
    Hits = 1
    strSQL = "insert into CallerID values ('" & PartyID & "', '" & PartyIDName & "', False, '" & Date() & "', '', " & Hits & ")"
  else
    Hits = objRSdo("HitCount") + 1
    if not isnull(objRSdo("PartyIDName")) or objRSdo("PartyIDName") = "" or objRSdo("PartyIDName") = " " then
      PartyIDName = objRSdo("PartyIDName")
'      PartyIDName = replace(PartyIDName, "\", "\\")
'      PartyIDName = replace(PartyIDName, "'", "\'")
      PartyIDName = replace(PartyIDName, "'", "")
    end if
    strSQL = "update CallerID set PartyIDName='" & PartyIDName & "', Hit='" & Date() & "', HitCount=" & Hits & " where PartyID='" & PartyID & "'"
  end if
  objRS.open strSQL, objCalldata, adOpenStatic, adLockOptimistic
  objRSdo.Close
  strSQL = "select sum(TalkTimeSeconds), sum(HoldTime) from `Connect` where ConnectReason='17' and CallTableID=" & objRSCall("ID") & " group by CallTableID"
  objRSdo.Open strSQL, objCall
  if not objRSdo.eof then
    TalkTime = objRSdo("sum(TalkTimeSeconds)")
    on error resume next
'****************
    HoldTime = objRSdo("sum(HoldTime)")
    do until HoldTime < 18991230000000
      HoldTime = HoldTime - 18991230000000
    loop
'    HoldTime = objRSdo("sum(HoldTime)")
'****************
    on error goto 0
  else
    TalkTime = objRSConnect("sum(TalkTimeSeconds)")
    HoldTime = "0:00:00"
  end if
  objRSdo.Close
  strSQL = "insert into Calldata values ('" & objRSCall("Extension") & "', '" & Description & "', '" & PartyID & "', '" & PartyIDName & "', False, False, '" & objRSCall("StartTime") & "', '" & objRSCall("EndTime") & "', '" & TalkTime & "', '" & HoldTime & "', " & objRSConnect("LongDistance") & ")"
  objRSdo.open strSQL, objCalldata, adOpenStatic, adLockOptimistic
  objRSDN.close
  objRSConnect.close
  objRSCall.movenext
loop
objRSCall.close

'****************
'Gather Internal, Outbound & Inbound data
'****************
strSQL = "select ID, Extension, StartTime, EndTime from `Call` where CallType=1 and EndTime > '" & LastTime & "'"
objRSCall.open strSQL, objCall
do until objRSCall.eof
  strSQL = "select PartyID, PartyIDName, PartyIDLastName, sum(TalkTimeSeconds), sum(HoldTime), PartyType from `Connect` where ConnectReason=17 and CallTableID=" & objRSCall("ID") & " group by CallTableID"
  objRSConnect.open strSQL, objCall
  strSQL = "select Description from DN where DN='" & objRSCall("Extension") & "'"
  objRSDN.open strSQL, objCfg
  'post info to Phone.mdb
  If Not objRSConnect.BOF Then
    PartyID = objRSConnect("PartyID")
    PartyIDName = objRSConnect("PartyIDName") & " " & objRSConnect("PartyIDLastName")
'    PartyIDName = replace(PartyIDName, "\", "\\")
'    PartyIDName = replace(PartyIDName, "'", "\'")
    PartyIDName = replace(PartyIDName, "'", "")
  Else
    PartyID=""
    PartyIDName = ""
  End If
  wscript.echo objRSCall("ID") & " " & objRSCall("Extension") & " " & objRSCall("StartTime") & " " & objRSCall("EndTime")
  if not objRSDN.eof then
    wscript.echo "  Description:" & objRSDN("Description")
    Description = objRSDN("Description")
  else
    wscript.echo "  Description: (none)"
    Description = ""
  end if
  wscript.echo "  PartyID:" & PartyID
  wscript.echo "  PartyIDName:" & PartyIDName
  wscript.echo "  TalkTimeSeconds:" & objRSConnect("sum(TalkTimeSeconds)")
'****************
  HoldTime = objRSConnect("sum(HoldTime)")
  do until HoldTime < 18991230000000
    HoldTime = HoldTime - 18991230000000
  loop
  wscript.echo "  HoldTime:" & HoldTime
'****************
  wscript.echo "  PartyType:" & objRSConnect("PartyType")
  wscript.echo

  objRSCalldata.AddNew
  objRSCalldata("Extension") = objRSCall("Extension")
  if not objRSDN.eof then
    objRSCalldata("Description") = objRSDN("Description")
  else
    objRSCalldata("Description") = ""
  end if
  objRSCalldata("PartyID") = objRSConnect("PartyID")
  objRSCalldata("PartyIDName") = objRSConnect("PartyIDName") & " " & objRSConnect("PartyIDLastName")
  objRSCalldata("Internal") = True
  objRSCalldata("Inbound") = False
  objRSCalldata("StartTime") = objRSCall("StartTime")
  objRSCalldata("EndTime") = objRSCall("EndTime")
  objRSCalldata("TalkTime") = objRSConnect("sum(TalkTimeSeconds)")
  objRSCalldata.Update

  objRSCalldata.AddNew
  objRSCalldata("Extension") = objRSConnect("PartyID")
  objRSCalldata("Description") = objRSConnect("PartyIDName") & " " & objRSConnect("PartyIDLastName")
  objRSCalldata("PartyID") = objRSCall("Extension")
  if not objRSDN.eof then
    objRSCalldata("PartyIDName") = objRSDN("Description")
  else
    objRSCalldata("PartyIDName") = ""
  end if
  objRSCalldata("Internal") = True
  objRSCalldata("Inbound") = True
  objRSCalldata("StartTime") = objRSCall("StartTime")
  objRSCalldata("EndTime") = objRSCall("EndTime")
  objRSCalldata("TalkTime") = objRSConnect("sum(TalkTimeSeconds)")
  objRSCalldata.Update
  objRSDN.close
  objRSConnect.close
  objRSCall.movenext
loop
objRSCall.close

'End Test
end if
