Option Explicit

'Dim strCon: strCon = "Driver={Microsoft ODBC for Oracle}; " & _
Dim strCon: strCon = "Driver={Oracle in instantclient11_1};dbq=PVPROD; uid=proview_adm;pwd=proview_adm_123;"


Dim strDvcCnt:
Dim strSuccess:
Dim strFailure:
Dim FlrDvs:
Dim strDate:



strDvcCnt = "SELECT count(distinct deviceid) FROM distributionentry WHERE distributionid in (53,54,55,57,58,59,61,62,63,70,71,72,83)"
strSuccess = "SELECT count(distinct deviceid) FROM distributionentry WHERE distributionid in (53,54,55,57,58,59,61,62,63,70,71,72,83) AND deviceid IN (SELECT resourceid FROM jobresult WHERE trunc(timestamp)=trunc(sysdate)-1 AND jobid like 'EJ%' AND result=0)"
strFailure = "SELECT count(distinct deviceid) FROM distributionentry WHERE distributionid in (53,54,55,57,58,59,61,62,63,70,71,72,83) AND deviceid NOT IN (SELECT resourceid FROM jobresult WHERE trunc(timestamp)=trunc(sysdate)-1 AND jobid like 'EJ%' AND result=0)" 					 
FlrDvs = "SELECT distinct deviceid FROM distributionentry WHERE distributionid in (53,54,55,57,58,59,61,62,63,70,71,72,83) AND deviceid IN (SELECT resourceid FROM jobresult WHERE trunc(timestamp)=trunc(sysdate)-1 AND jobid like 'EJ%' AND result=0)"
strDate = "select to_char(sysdate-2,'DD-MON-YYYY') from dual"


Dim fso:
Set fso = CreateObject("Scripting.FileSystemObject")



'* Create log file
Dim LogFile
Set LogFile = fso.CreateTextFile("E:\Proview\Logs\WincorEJSuccess.txt", true)
LogFile.WriteLine("================================================================================")
LogFile.WriteLine("|                    ProView EJ Pull Summary                                   |")
LogFile.WriteLine("================================================================================")
LogFile.WriteBlankLines 1
LogFile.WriteBlankLines 1
			 

Dim oCon: 
Set oCon = WScript.CreateObject("ADODB.Connection")
Dim oRs: 
Set oRs = WScript.CreateObject("ADODB.Recordset")

Dim SuccsVal
Dim HSuccsVal
Dim DSuccsVal
Dim FailVal
Dim HFailVal
Dim DFailVal
Dim DvcCnt
Dim HDvcCnt
Dim DDvcCnt
Dim devid
'Dim Hdevid
'Dim Ddevid
Dim errspec
'Dim Herrspec
'Dim Derrspec
Dim spcnt
Dim fpcnt
Dim FrDte

oCon.Open strCon

Set oRs = oCon.Execute(strDvcCnt)
DvcCnt=oRs.Fields(0).Value

LogFile.WriteLine("EJ Pull Scheduled:" )
LogFile.WriteLine("Wincor:"& DvcCnt)
LogFile.WriteBlankLines 1
LogFile.WriteLine("Sucess Count:")


Set oRs = oCon.Execute(strSuccess)
SuccsVal=oRs.Fields(0).Value

LogFile.WriteLine("Wincor:" & SuccsVal)
LogFile.WriteBlankLines 1
LogFile.WriteLine("Failure Count:")


Set oRs = oCon.Execute(strFailure)
FailVal =oRs.Fields(0).Value

LogFile.WriteLine("Wincor:" & FailVal)



Set oRs = oCon.Execute(strDate)
FrDte=oRs.Fields(0).Value



LogFile.WriteBlankLines 1
LogFile.WriteLine("================================================================================")
LogFile.WriteLine("|                     Success Devices -" & FrDte & "                            |")
LogFile.WriteLine("================================================================================")
LogFile.WriteBlankLines 1
LogFile.WriteBlankLines 1

LogFile.WriteLine("ATMID")

Set oRs = oCon.Execute(FlrDvs)

While Not oRs.EOF
	
	
	devid = oRs.Fields(0).Value
		
	LogFile.WriteLine(devid)
	
	oRs.MoveNext
Wend


oCon.Close
LogFile.WriteLine("================================================================================")
LogFile.WriteLine("|                                End of List                                   |")
LogFile.WriteLine("================================================================================")

Set oRs = Nothing
Set oCon = Nothing

Dim objMessage:
Dim strMessage:
Dim strSgntr:

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "/ProView/Total EJ Pull-Summary Wincor ATMs"
objMessage.From = "proview.support@"


'objMessage.To = ""
'objMessage.Cc = ""

objMessage.To = ""
objMessage.Cc = ""
objMessage.Bcc = "

strMessage =  "<font face='Verdana' size='2'><table border=1 cellpadding=0 cellspacing=0 width=45% style='border-collapse:collapse;font-size:12px'><tr><td colspan=4 align=center>ICICI Daily EJ Pulling Report Through ProView - " & FrDte & "</tr>" & _
	      "<tr><th>Machine Type</th><th>Wincor</th></tr>" & _
	      "<tr><td>Total ATMs</td><td>" & DvcCnt & "</td></tr>" & _
	      "<tr><td>Success Count</td><td>" & SuccsVal & "</td></tr>" & _
	      "<tr><td>Failure Count</td><td>" & FailVal & "</td></tr>" & _
	      "<tr><td>Success %</td><td>&nbsp;</td></tr>" & _
	      "<tr><td>Failure %</td><td>&nbsp;</td></tr>" & _
	      "</table></font>"

'The line below shows how to send using HTML included directly in your script
objMessage.HTMLBody = "<p align = left style='font-size:12px;font-family:Verdana'> Dear Sir / Madam" & _
		      "<br>" & _
		      "<br>" & _
		      "<br>" & _
		      "ProView EJ Pull Summary</p>" & _
		      "<br>" & _
		       strMessage & _
		       "<br>"& _
		       "<p style='font-size:12px;font-family:Verdana'>Attached herewith the Success devices for your action" & _
		       "<br>"& _
		       "<br>"& _
		       "Regards" & _
		       "<br>" & _
		       "<b>ProView Support Team</b></p>"


'The line below shows how to send a webpage from a remote site
''objMessage.CreateMHTMLBody "http://www.paulsadowski.com/wsh/"

'The line below shows how to send a webpage from a file on your machine
'objMessage.CreateMHTMLBody "file://c|/temp/test.htm"


objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="10.16.0.39"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
objMessage.Configuration.Fields.Update

LogFile.Close
objMessage.AddAttachment "E:\Proview\Logs\WincorEJSuccess.txt"


objMessage.Send
