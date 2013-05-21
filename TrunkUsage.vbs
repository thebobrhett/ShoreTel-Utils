'****************
'Bob Rhett - Wednesday, August 11, 2010
'  Created to monitor trunk usage
'Keith Brooks - Thursday, October 6, 2011
'  Added check for last record more recent than StartDate.
'****************
'on error resume next

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3

dim PastDays
dim StartDate
dim EndDate
dim LastDate
dim TestTime
dim EndTime
dim dbCall
dim dbPhone
dim rs
dim strSQL
dim iCnt
dim iMax
dim iDur
dim iPnt
dim iAvg
dim bCont
dim tPeak
dim StartTimeStamp
dim strMode

strMode = "s"
'strMode = "n"
'strMode = "h"
StartTimeStamp = now()
PastDays = 30
EndDate = date()
StartDate = DateAdd("d", 0 - Pastdays, EndDate)

set dbCall = CreateObject("adodb.connection")
set dbPhone = CreateObject("adodb.connection")
dbCall.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=Omaha;database=shorewarecdr;port=4309;user=st_cdrreport;password=passwordcdrreport;"
dbPhone.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\Assetmgt\Phone.mdb"
set rs = CreateObject("adodb.recordset")
dbPhone.CursorLocation = adUseClient

strSQL = "delete from Trunkdata where RecDate<#" & StartDate & "#"
'rs.open strSQL, dbPhone, adOpenStatic, adLockOptimistic

do until StartDate => EndDate
  iCnt = 0
  iMax = 0
  iDur = 0
  iPnt = 0
  iAvg = 0
  bCont = True
  strSQL = "select * FROM TrunkUsage order by RecDate desc"
  rs.open strSQL, dbPhone, adOpenStatic, adLockOptimistic
  if not rs.eof then
    If CDate(rs("RecDate")) > CDate(StartDate) Then
      StartDate = DateAdd("d", 1, rs("RecDate"))
    End If
  end if
  rs.close
  EndTime = DateAdd("d", 1, StartDate)
  TestTime = StartDate

wscript.echo "StartDate: " & StartDate
wscript.echo "EndDate: " & EndDate
wscript.echo "TestTime: " & TestTime
wscript.echo "EndTime: " & EndTime
wscript.echo
  if StartDate => EndDate then exit do

  do until cdate(TestTime) => cdate(EndTime)
    wscript.echo "Test Time:" & TestTime
    strSQL = "select count(*) as iCnt from Calldata where StartTime<=#" & TestTime & "# and EndTime>=#" & TestTime & "#"
    rs.open strSQL, dbPhone, adOpenStatic, adLockOptimistic
    iCnt = rs("iCnt")
    if not rs.eof then
      iPnt = iPnt + 1
      if iCnt > iMax then
        iMax = iCnt
        iDur = 1
        tPeak = cdate(TestTime)
      elseif iCnt < iMax then
        iCont = False
      else
        iDur = iDur + 1
      end if
      iAvg = ((iAvg * iPnt) + iCnt) / (iPnt + 1)
    else
      iAvg = (iAvg * iPnt) / (iPnt + 1)
    end if
    wscript.echo "Datapoint:" & iPnt
    wscript.echo "Current Usage:" & iCnt
    wscript.echo "Average:" & iAvg
    wscript.echo "Max In Use:" & iMax
    wscript.echo "Duration:" & iDur
    rs.close
    TestTime = dateadd(strMode, 1, TestTime)
    TestTime = cstr(year(TestTime)) & "-" & cstr(month(TestTime)) & "-" & cstr(day(TestTime)) & " " & cstr(hour(TestTime)) & ":" & cstr(minute(TestTime)) & ":" & cstr(second(TestTime))
    wscript.echo "End Time:" & EndTime
    wscript.echo "Next Test Time:" & TestTime
    wscript.echo
  loop

  wscript.echo "For " & StartDate
  wscript.echo "On average, " & formatnumber(iAvg, 0) & " channels were in use."
  select case strMode
    case "s"
      if iMax = 1 and iDur = 1 then
        wscript.echo "Peak usage: " & iMax & " channel for " & iDur & " second."
      elseif iMax > 1 and iDur = 1 then
        wscript.echo "Peak usage: " & iMax & " channels for " & iDur & " second."
      elseif iMax = 1 and iDur > 1 then
        wscript.echo "Peak usage: " & iMax & " channel for " & iDur & " seconds."
      elseif iMax > 1 and iDur > 1 then
        wscript.echo "Peak usage: " & iMax & " channels for " & iDur & " seconds."
      else
        wscript.echo "No usage recorded."
      end if
    case "n"
      if iMax = 1 and duration = 1 then
        wscript.echo "Peak usage: " & iMax & " channel for " & iDur & " minute."
      elseif iMax > 1 and iDur = 1 then
        wscript.echo "Peak usage: " & iMax & " channels for " & iDur & " minute."
      elseif iMax = 1 and iDur > 1 then
        wscript.echo "Peak usage: " & iMax & " channel for " & iDur & " minutes."
      elseif iMax > 1 and iDur > 1 then
        wscript.echo "Peak usage: " & iMax & " channels for " & iDur & " minutes."
      else
        wscript.echo "No usage recorded."
      end if
    case "h"
      if iMax = 1 and duration = 1 then
        wscript.echo "Peak usage: " & iMax & " channel for " & iDur & " hour."
      elseif iMax > 1 and iDur = 1 then
        wscript.echo "Peak usage: " & iMax & " channels for " & iDur & " hour."
      elseif iMax = 1 and iDur > 1 then
        wscript.echo "Peak usage: " & iMax & " channel for " & iDur & " hours."
      elseif iMax > 1 and iDur > 1 then
        wscript.echo "Peak usage: " & iMax & " channels for " & iDur & " hours."
      else
        wscript.echo "No usage recorded."
      end if
  end select
  if bCont = True then
    wscript.echo "This peak usage occurred at " & tPeak & " and was contiguous."
  else
    wscript.echo "This peak usage occurred at " & tPeak & " but was not contiguous."
  end if
  strSQL = "insert into TrunkUsage values (#" & StartDate & "#, " & iMax & ", " & iDur & ", " & int(formatnumber(iAvg, 0)) & ", #" & tPeak & "#, " & bCont & ")"
  wscript.echo strSQL
  if tPeak <> "" then
    rs.open strSQL, dbPhone, adOpenStatic, adLockOptimistic
  end if
  StartDate = DateAdd("d", 1, StartDate)
loop

wscript.echo
wscript.echo "The program started running at " & StartTimeStamp
wscript.echo "the program finished running at " & now()

