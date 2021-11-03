'This script displays a maintenance popup window.
'
'Run this script from the Task Scheduler using WScript.
'
'wscript "c:\path_to_script\maintMsg.vbs"
'
'I set this to run at logon for any user. They must have read/execute access to the script to run.
'
'This script will check today's date.
'If the day is Mon-Sat, it will show the popup to users logging in.


Sub Main()
	'Get today's date.
	Dim today
	today = Date
	
	'Get today's weekday (int)
	Dim weekdaynum
	weekdaynum = Weekday(today)
	
	'Check which weekday it is
	'Adds days to make the maint_day Saturday
	'Remove any days you don't want to show message.
	Select Case weekdaynum
		Case vbMonday
			Call MaintMsg(DateAdd("d",5, today))
		Case vbTuesday
			Call MaintMsg(DateAdd("d",4, today))
		Case vbWednesday
			Call MaintMsg(DateAdd("d",3, today))
		Case vbThursday
			Call MaintMsg(DateAdd("d",2, today))
		Case vbFriday
			Call MaintMsg(DateAdd("d",1, today))
		Case vbSaturday
			Call MaintMsg(today)
	End Select
End Sub

Sub MaintMsg(maint_day)
	dim showMsg 
	showMsg = CheckFor3rdTuesday(DateAdd("d",-4, maint_day))
	
	if showMsg = true then
		Call MsgBox("It is time for our planned maintenance window to ensure the safety and security of your hosted experience with {company}"& vbCrLf & vbCrLf & "On " & WeekdayName(Weekday(maint_day)) & " (" & MonthName(Month(maint_day)) &" " & Day(maint_day) & ") from 18:00 - 22:00 Pacific, hosted services may be unavailable. Please plan your activities with this maintenance window in mind.",vbSystemModal,"Maintenance Reminder") ' Display message on computer screen.
	end if
End Sub

Function CheckFor3rdTuesday(maint_day)
	if Int((Day(maint_day) - 1) / 7) = 2 Then
		CheckFor3rdTuesday = true
	Else
		CheckFor3rdTuesday = false
	end if
end function

'Begin program
call Main
WScript.Quit

' vbSunday (1)  
' vbMonday (2)  
' vbTuesday (3)  
' vbWednesday (4)  
' vbThursday (5)  
' vbFriday (6)  
' vbSaturday (7)  
