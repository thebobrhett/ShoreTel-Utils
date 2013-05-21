<%@ LANGUAGE="VBSCRIPT" %>
<%
'*************************************************************************************
' Modification history:
'     Engineer          Date         Modification
'     Keith Brooks  12/23/2008   Added call to HitCounters function to track number
'				     of page hits.
'
'     Bob Rhett         7/28/2010    Added forced entries to be included in list.
'
'     Bob Rhett         9/22/2010    Added more forced entries.
'     Bob Rhett         1/17/2011    Added forced entry for KP.
'     Keith Brooks      10/3/2011    Updated Shoreware database connection string to MySQL.
'     Bob Rhett         4/23/2012    Added forced entry for BHFCU.
'*************************************************************************************
%>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<%
'<!--#include virtual="Functions/HitCounter.asp"-->
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<hta:Application
  Border = 'Thick'
  BorderStyle = 'Complex'
  ShowInTaskBar = 'No'
  MaximizeButton = 'No'
  MinimizeButton = 'No'
>
<title>AKSA Phone System Directory</title>
<link rel=STYLESHEET href='http://mogsa4/aksastyle.css' type='text/css'>
<%
'Dim	HitCounts

'Set/get hit counts.
'HitCounts = HitCounter("assetmgt_directory")
%>
</head>
<body link='black' vLink='black'>
<h1><p align='center'>Asahi Kasei Spandex America</p>
<h2><p align='center'>Note: Dial 82 + last 4 digits when calling from Bushy Park

<%
'on error resume next
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3

dim objCfg
dim objRS
dim objRSdo
dim strSQL
dim force
dim forcename
dim forcenumber
dim done
dim previous
dim CountofUsers
dim LineCount
dim TotalCount
dim list
dim listexternal
dim listemail
dim listcell
dim listpager

force = False
forcename = ""
done = False
list = True
'Seed the user count with the number of forced entries
CountofUsers = 11
LineCount = 1
TotalCount = 1

set objCfg = CreateObject("adodb.connection")
'objCfg.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Shoreline Data\Database\ShoreWare.mdb"
objCfg.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=omaha;database=shoreware;port=4308;user=st_configread; pwd=passwordconfigread;"
set objRS = CreateObject("adodb.recordset")
set objRSdo = CreateObject("adodb.recordset")
objRS.CursorLocation = adUseClient
objRSdo.CursorLocation = adUseClient

objRS.Open "SELECT Count(UserDN) AS CountofUsers FROM Users " & _
  "WHERE UserDN < '2000' OR UserDN > '6000'" _
, objCfg, adOpenStatic, adLockOptimistic
CountofUsers = CountofUsers + objRS("CountofUsers")
objRS.Close

strSQL = "select DN from DN where DNTypeID='1' or DNTypeID='8' or DNTypeID='22' and Hidden='0'"

objRS.Open "SELECT TABAddresses.AddressID, Users.AddressID, TABAddresses.FirstName, TABAddresses.LastName, Users.UserDN, TABAddresses.EmailAddress, TABAddresses.CellPhone, TABAddresses.PagerPhone FROM TABAddresses " & _
  "INNER JOIN (SELECT Users.AddressID, Users.UserDN FROM Users) AS Users ON TABAddresses.AddressID = Users.AddressID " & _
  "ORDER BY TABAddresses.LastName, TABAddresses.FirstName", objCfg, adOpenStatic, adLockOptimistic

Response.Write "<table border='1'><tr><td valign='top'>"
Response.Write "<table>"
'Response.Write "<tr bgcolor='powderblue'><th><h2>Name</th><th><h2>Extension</th><th><h2>Direct In Dial</th><th><h2>Email Address</th><th><h2>Cell Phone</th><th><h2>Pager</th></tr>"
Response.Write "<tr bgcolor='powderblue'><th><h2>Name</th><th><h2>Number</th><th><h2>Cell Phone</th></tr>"

previous = objRS("LastName") & ", " & objRS("FirstName")

Do While not objRS.EOF

  'Exceptions (do not list)

  'Last name = Data
  if objRS("LastName") = "Data" then
    list = False
    CountofUsers = CountofUsers - 1
  end if

  'Last name = Fax
'  if objRS("LastName") = "Fax" then
'    list = False
'    CountofUsers = CountofUsers - 1
'  end if

  '5000 series numbers
  if objRS("UserDN") > 2000 and objRS("UserDN") < 6000 then
    list = false
  end if

  'set the exchange to 725 or 820 depending on the series
  listexternal = objRS("UserDN")
  if listexternal > 1699 and listexternal < 1921 then
'    listexternal = "+1 (843) 725-" & listexternal
    listexternal = "725-" & listexternal
  elseif listexternal > 6499 and listexternal < 6600 then
'    listexternal = "+1 (843) 820-" & listexternal
    listexternal = "820-" & listexternal
  end if

  'aksa & numeric email addresses
  listemail = objRS("EmailAddress")
  if listemail = "aksa@dorlastan.com" then listemail = "&nbsp"

  on error resume next
  if left(listemail, 4) < 9999 then
    if err.number = 0 then
      listemail = "&nbsp"
    end if
  end if
  err.clear
  on error goto 0

  listcell = objRS("CellPhone")
  if isnull(listcell) then
    listcell = "&nbsp"
  else
    listcell = right(objRS("CellPhone"), 8)
  end if

  listpager = objRS("PagerPhone")
  if isnull(listpager) then listpager = "&nbsp"

  if list = True then

    'Exceptions (force to list)
    'This looks weird but do them in alphabetical order such that one entry sets up for the next.
    'Move done = True to the last entry.
    'Don't forget to modify the CountofUsers seed at the top.
    do while strComp(forcename, objRS("LastName") & ", " & objRS("FirstName")) = -1
      if force = False then
        select case forcename
          case ""
            force = True
            forcename = "C11-1, KP Unit"
            forcenumber = "820-6554"
          case "C11-1, KP Unit"
            force = True
            forcename = "C11-1, Tank Farm"
            forcenumber = "725-1701"
          case "C11-1, Tank Farm"
            force = True
            forcename = "C11-1, Winder Shop"
            forcenumber = "725-1860"
          case "C11-1, Winder Shop"
            force = True
            forcename = "C11-2, Inspection"
            forcenumber = "820-6559"
          case "C11-2, Inspection"
            force = True
            forcename = "C11-2, Shipping Dock"
            forcenumber = "820-6556"
          case "C11-2, Shipping Dock"
            force = True
            forcename = "C11-2, Warping"
            forcenumber = "820-6550"
          case "C11-2, Warping"
            force = True
            forcename = "Credit Union, BHFCU"
            forcenumber = "725-5070"
          case "Credit Union, BHFCU"
            force = True
            forcename = "Dirks, Mike"
            forcenumber = "820-6513"
          case "Dirks, Mike"
            force = True
            forcename = "IT Help Desk"
            forcenumber = "725-1900"
          case "IT Help Desk"
            force = True
            forcename = "Johnson, Patrick"
            forcenumber = "820-6513"
          case "Johnson, Patrick"
            force = True
            forcename = "Stello, Chuck"
            forcenumber = "820-6513"
            done = True
        end select
      end if
      if strComp(forcename, previous) = 1 and strComp(forcename, objRS("LastName") & ", " & objRS("FirstName")) = -1 then
        Response.Write "<tr><td><h4>" & forcename & "</td><td align='center'><h4>" & forcenumber & "</td><td align='center'><h4>&nbsp;</td></tr>"
        force = False
        previous = forcename
        LineCount = LineCount + 1
        TotalCount = TotalCount + 1
      end if
      if done = True then exit do
    loop

'    Response.Write "<tr><td><h5><font color='black'>" &  objRS("LastName") & ", " & objRS("FirstName") & "</font></td><td><h5><font color='black'>" & objRS("UserDN") & "</font></td><td><h5><font color='black'>" & listexternal & "</font></td><td><h5><font color='black'>" & listemail & "</font></td><td><h5><font color='black'>" & listcell & "</font></td><td><h5><font color='black'>" & listpager & "</font></td></tr>"
    Response.Write "<tr><td><h4>" &  objRS("LastName") & ", " & objRS("FirstName") & "</td><td align='center'><h4>" & listexternal & "</td><td align='center'><h4>" & listcell & "</td></tr>"
'    Response.Write "<tr><td><h4>(" & LineCount & ") " &  objRS("LastName") & ", " & objRS("FirstName") & " (" & TotalCount & ")</td><td align='center'><h4>" & listexternal & "</td><td align='center'><h4>" & listcell & "</td></tr>"
    previous = objRS("LastName") & ", " & objRS("FirstName")
    LineCount = LineCount + 1
    TotalCount = TotalCount + 1
  end if

  if TotalCount <= CountofUsers then
    if LineCount > ((CountofUsers + 1) / 2) then
      LineCount = 1
      Response.Write "</table></td><td valign='top'>"
      Response.Write "<table>"
'      Response.Write "<tr bgcolor='powderblue'><th><h2>Name</th><th><h2>Extension</th><th><h2>Direct In Dial</th><th><h2>Email Address</th><th><h2>Cell Phone</th><th><h2>Pager</th></tr>"
      Response.Write "<tr bgcolor='powderblue'><th><h2>Name</th><th><h2>Number</th><th><h2>Cell Phone</th></tr>"
'      Response.Write "<tr bgcolor='powderblue'><th><h2>Name (" & CountofUsers & ")</th><th><h2>Number</th><th><h2>Cell Phone</th></tr>"
    end if
  else
    Response.Write "</table></td>"
  end if

  list = True
  objRS.MoveNext
Loop

objRS.Close
Response.Write "</table></p>"

%>

</font>
</p>
</body>
</html>
