Sub procBegin()
  If TimeValue(Now()) < TimeValue("11:30:00") Then
    'MsgBox ("Starting")
    Call iterBRD("Start")
  End If
'call the birthday

End Sub

Sub iterBRD(yes As String)
Set mybk = ThisWorkbook
mysheet = "BRD"
mybk.Sheets(mysheet).Activate
Dim lastRow As Integer, dayDiff As Integer, i As Integer
Range("A1").End(xlDown).Select
lastRow = ActiveCell.Row
   
'Create active mail ID list
Dim idlist As String
idlist = ";"
For i = 2 To lastRow
    If Range("E" & i).Value = 1 Then
        idlist = idlist + Range("C" & i).Value + ";"
    End If
Next
   
'iterate over multiple rows
   For i = 2 To lastRow
     'MsgBox (Cells(i, 1).Value - DateValue(Now))
     On Error Resume Next
     dayDiff = Range("B" & i).Value - DateValue(Now)
     'MsgBox (dayDiff)
     'send for recently past
     If Range("E" & i).Value = 1 Then
        If dayDiff > -4 Then
            If dayDiff <= 0 Then
                If LCase(Range("D" & i).Value) = "haveto" Then
                    Call brdMail(i, idlist)
                    Range("B" & i).Value = DateAdd("yyyy", 1, Range("B" & i).Value)
                    Range("D" & i).Value = "Done"
                    Range("F" & i).Value = " "
                End If
            End If
        End If
        If dayDiff < 7 Then
            If dayDiff > 0 Then
                Range("D" & i).Value = LCase("haveto")
            End If
        End If
     End If
    Next
End Sub
Sub brdMail(rowNum As Integer, cclist As String)
Dim wk As Worksheet
Dim ol As Object
Dim em As Object
Set ol = CreateObject("Outlook.Application")
Set em = ol.CreateItem(olMailItem)
Dim name As String
name = Range("A" & rowNum).Text
msg = Range("F" & rowNum).Text
bdayday = Day(Range("B" & rowNum).Value)
bdaymonth = MonthName(Month(Range("B" & rowNum).Value), True)
bday = CStr(bdayday) + "-" + bdaymonth

msgstart = "<p><span style='font-family: comic\ sans\ ms, sans-serif; font-size: 14pt; color: #800000;'>Dear <span style='font-size: 18pt;'>&nbsp;" + UCase(name) + " !!!!!!</span></span></p>"
msgstart2 = "<p>&nbsp;</p><p><span style='font-family: comic\ sans\ ms, sans-serif; font-size: 13pt; color: #000080;'>Many congratulations on your birthday (" + bday + ") !!<span style='color: #339966;'> BEST WISHES FOR MANY MORE YEARS TO COME!!</span></span></p>"

msgmid = "<p>&nbsp;</p><p><span style='font-family: comic\ sans\ ms, sans-serif; font-size: 16pt; color: #800000;'>'" + msg + "'</span></p>"
msgend = "<p>&nbsp;</p><p>&nbsp;</p><hr /><span style='color: #000080;font-size: 16pt;font-family: comic\ sans\ ms, sans-serif;'><strong>&nbsp;&nbsp;Cheers!!! </strong><br/><em>'Totally Unofficial'</em></span>"

em.To = Range("C" & rowNum).Text
em.CC = cclist
em.Subject = "Wishing You Happy Birthday" + " " + UCase(name) + " :) :)"
em.HTMLBody = msgstart + msgstart2 + " " + msgmid + " " + msgend

'Creating the waiting time
'rndMins = Int((90 - 1 + 1) * Rnd + 1)
'newHour = Hour(Now())
'newMinute = Minute(Now()) + rndMins
'newSecond = Second(Now())
'waitTime = TimeSerial(newHour, newMinute, newSecond)
'Application.Wait waitTime
em.Send
'em.Display
Set ol = Nothing
Set em = Nothing
Set wk = Nothing
End Sub
