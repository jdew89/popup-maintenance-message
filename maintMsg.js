/*This script displays a maintenance popup window.

Run this script from the Task Scheduler using WScript. 

wscript "c:\path_to_script\maintMsg.vbs"

I set this to run at logon for any user. They must have read/execute access to the script to run.

This script will check today's date.
If the day is Mon-Sat, it will show the popup to users logging in.
*/

var today = new Date();

var weekday = today.getUTCDay();

var maintday = today;

//First weekday is Sunday as 0
//Change this to limit which days to prompt message
if(weekday > 0) { //not sunday
	maintday.setDate( maintday.getUTCDate() + 6 - weekday); //sets maint day to saturday
	
	var prev_tuesday_date = maintday.getUTCDate() - 4; 
	
	//if tuesday was the 3rd tuesday, show msg
	if(Math.floor(prev_tuesday_date / 7 ) == 3){
		//builds options for the window and then executes it.
		var mbOK = 0;
		var mbOKCancel = 1;
		var mbCancel = 2;
		var mbInformation = 64; // Information icon
		var text  = "It is time for our planned maintenance window to ensure the safety and security of your hosted experience with {business}.\n\nOn " + maintday.toDateString() + " from 18:00 - 22:00 Pacific, hosted services may be unavailable. Please plan your activities with this maintenance window in mind.";
		var title = "Maintenance Reminder";
		var wshShell = WScript.CreateObject("WScript.Shell");
		var popup_window =  wshShell.Popup(text,
						0, //if > 0, seconds to show popup. With 0 it is without timeout.
						title,
						mbOK + mbInformation);
						
		WScript.Quit(); //End of WScript (every for & while loops end too)
	}
}