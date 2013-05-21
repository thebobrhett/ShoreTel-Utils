 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>AKSA Phone System Cost Detail</title>
<link href='phone.css' rel='stylesheet' type='text/css'>
<style>
th        { font-size:12; }
td        { font-size:12; }
</style>
</head>
<body>
<h1>AKSA Phone System Reports - Cost Detail</h1>
<%
'on error resume next
dim objCalldata
dim objRS

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3

set objCalldata = CreateObject("adodb.connection")
objCalldata.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\AssetMgt\Phone.mdb"
set objRS = CreateObject("adodb.recordset")
objRS.CursorLocation = adUseClient

'****************
'Cost Detail
'****************
objRS.Open "SELECT * FROM VoiceCost ORDER BY SortOrder", objCalldata, adOpenStatic, adLockOptimistic
Response.Write "<hr/>"
Response.Write "<p>"
Response.Write "<table border='1'>"
Response.Write "<caption>Cost Detail (yellow highlighted numbers are budgetary)</caption>"
Response.Write "<tr><th>Service</th><th>Jan</th><th>Feb</th><th>Mar</th><th>Apr</th><th>May</th><th>Jun</th><th>Jul</th><th>Aug</th><th>Sep</th><th>Oct</th><th>Nov</th><th>Dec</th><th>Total</th></tr>"
Do While not objRS.EOF
  ServiceTotal = 0
  Response.Write "<tr><td>" & objRS("Service") & "</td>"
  'January
  if objRS("Mth1") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth1Total = Mth1Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth1")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth1")
    Mth1Total = Mth1Total + objRS("Mth1")
  end if
  'February
  if objRS("Mth2") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth2Total = Mth2Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth2")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth2")
    Mth2Total = Mth2Total + objRS("Mth2")
  end if
  'March
  if objRS("Mth3") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth3Total = Mth3Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth3")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth3")
    Mth3Total = Mth3Total + objRS("Mth3")
  end if
  'April
  if objRS("Mth4") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth4Total = Mth4Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth4")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth4")
    Mth4Total = Mth4Total + objRS("Mth4")
  end if
  'May
  if objRS("Mth5") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth5Total = Mth5Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth5")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth5")
    Mth5Total = Mth5Total + objRS("Mth5")
  end if
  'June
  if objRS("Mth6") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth6Total = Mth6Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth6")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth6")
    Mth6Total = Mth6Total + objRS("Mth6")
  end if
  'July
  if objRS("Mth7") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth7Total = Mth7Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth7")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth7")
    Mth7Total = Mth7Total + objRS("Mth7")
  end if
  'August
  if objRS("Mth8") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth8Total = Mth8Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth8")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth8")
    Mth8Total = Mth8Total + objRS("Mth8")
  end if
  'September
  if objRS("Mth9") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth9Total = Mth9Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth9")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth9")
    Mth9Total = Mth9Total + objRS("Mth9")
  end if
  'October
  if objRS("Mth10") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth10Total = Mth10Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth10")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth10")
    Mth10Total = Mth10Total + objRS("Mth10")
  end if
  'November
  if objRS("Mth11") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth11Total = Mth11Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth11")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth11")
    Mth11Total = Mth11Total + objRS("Mth11")
  end if
  'December
  if objRS("Mth12") = 0 then
    Response.Write "<td class='hilite'>" & FormatCurrency(objRS("Budget") / 12) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Budget") / 12
    Mth12Total = Mth12Total + objRS("Budget") / 12
  else
    Response.Write "<td class='right'>" & FormatCurrency(objRS("Mth12")) & "</td>"
    ServiceTotal = ServiceTotal + objRS("Mth12")
    Mth12Total = Mth12Total + objRS("Mth12")
  end if
'<td>" & FormatCurrency(objRS("Mth2")) & "</td><td>" & FormatCurrency(objRS("Mth3")) & "</td><td>" & FormatCurrency(objRS("Mth4")) & "</td><td>" & FormatCurrency(objRS("Mth5")) & "</td><td>" & FormatCurrency(objRS("Mth6")) & "</td><td>" & FormatCurrency(objRS("Mth7")) & "</td><td>" & FormatCurrency(objRS("Mth8")) & "</td><td>" & FormatCurrency(objRS("Mth9")) & "</td><td>" & FormatCurrency(objRS("Mth10")) & "</td><td>" & FormatCurrency(objRS("Mth11")) & "</td><td>" & FormatCurrency(objRS("Mth12")) & "</td>"
'  CostOfService = objRS("Mth1") + objRS("Mth2") + objRS("Mth3") + objRS("Mth4") + objRS("Mth5") + objRS("Mth6") + objRS("Mth7") + objRS("Mth8") + objRS("Mth9") + objRS("Mth10") + objRS("Mth11") + objRS("Mth12")
  Response.Write "<td class='right'>" & FormatCurrency(ServiceTotal) & "</td></tr>"
  objRS.MoveNext
Loop
objRS.Close
Response.Write "<tr class='bold'>"
Response.Write "<td class='right'>Total</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth1Total) & "</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth2Total) & "</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth3Total) & "</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth4Total) & "</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth5Total) & "</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth6Total) & "</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth7Total) & "</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth8Total) & "</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth9Total) & "</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth10Total) & "</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth11Total) & "</td>"
Response.Write "<td class='right'>" & FormatCurrency(Mth12Total) & "</td>"
GrandTotal = Mth1Total + Mth2Total + Mth3Total + Mth4Total + Mth5Total + Mth6Total + Mth7Total + Mth8Total + Mth9Total + Mth10Total + Mth11Total + Mth12Total
Response.Write "<td class='right'>" & FormatCurrency(GrandTotal) & "</td></tr>"
Response.Write "<tr class='bold'><td colspan='13'align='right'>Cost of the Phone System per Month</td><td class='right'>" & FormatCurrency(GrandTotal / 12) & "</td></tr>"
Response.Write "</table></p>"

%>

</font>
</p>
</body>
</html>
