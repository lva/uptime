; EventLog Tool
; Version history:
; Version 1.0 - Initial version showing "Uptime 7 hours 30 minutes 48 seconds"
; Version 1.1 - Extended tool to show boot timestamp and to show uptime in days as well.
;               Display is now "Up since 2016/01/25 08:24:13 AM - Uptime 3 days 7 hours 36 minutes 9 seconds"
; Version 1.2 - Uptime is wrong when system is booted in the afternoon since time in event log is displayed in AM/PM. Convert time to 24-hour clock first.
;               Display is now "Up since 2016/01/25 08:24:13 - Uptime 3 days 7 hours 36 minutes 9 seconds"
;               Now also searches for the event log that displays the system uptime in seconds. This speeds up the searching process for systems that have an uptime of more that 1 day.

#pragma compile(Console, true)
#pragma compile(LegalCopyright, Lode Vanstechelman)
#pragma compile(CompanyName, Lode Vanstechelman)
#pragma compile(ProductName, Uptime)
#pragma compile(ProductVersion, 1.2)
#pragma compile(FileVersion, 1.2.0.0)
#pragma compile(Compatibility, win7)
#pragma compile(Out, uptime.exe)

Opt("MustDeclareVars", 1)

#include <Array.au3>
#include <EventLog.au3>
#include <GUIConstantsEx.au3>
#include <StringConstants.au3>

Uptime()

Func DisplayUptime($sBootTimestamp)
   Local $sCurrentTime = _NowCalc()

   Local $sOutput = "Up since " & $sBootTimestamp & " - Uptime "
   Local $iDiff = _DateDiff("d", $sBootTimestamp, $sCurrentTime)
   If $iDiff > 0 Then
	  If $iDiff = 1 Then
		 $sOutput &= $iDiff & " day "
	  Else
		 $sOutput &= $iDiff & " days "
	  EndIf
	  $sBootTimestamp = _DateAdd("d", $iDiff, $sBootTimestamp)
   EndIf

   $iDiff = _DateDiff("h", $sBootTimestamp, $sCurrentTime)
   $sOutput &= $iDiff & " hours "
   $sBootTimestamp = _DateAdd("h", $iDiff, $sBootTimestamp)

   $iDiff = _DateDiff("n", $sBootTimestamp, $sCurrentTime)
   $sOutput &= $iDiff & " minutes "
   $sBootTimestamp = _DateAdd("n", $iDiff, $sBootTimestamp)

   $iDiff = _DateDiff("s", $sBootTimestamp, $sCurrentTime)
   $sOutput &= $iDiff & " seconds "
   $sBootTimestamp = _DateAdd("s", $iDiff, $sBootTimestamp)

   ConsoleWrite($sOutput & @CRLF)
EndFunc   ;==>DisplayUptime

Func GetEventTimestamp($sDate, $sTime)
   Local $aDate = StringSplit($sDate, "/")
   Local $aTime = StringSplit($sTime, ": ")
   ; Transform AM/PM to 24-hour time.
   If $aTime[4] = "PM" Then
	  $aTime[1] = $aTime[1] + 12
   EndIf

   Return $aDate[3] & "/" & $aDate[1] & "/" & $aDate[2] & " " & $aTime[1] & ":" & $aTime[2] & ":" & $aTime[3]
EndFunc   ;==>GetEventTimestamp

Func Uptime()
   Local $hEventLog, $aEvent

   ; Uptime events are logged in the Windows System event log.
   $hEventLog = _EventLog__Open("", "System")
   ; Read the event log sequentially starting from the most recent event.
   $aEvent = _EventLog__Read($hEventLog, True, False)

   While ($aEvent[0] <> False)
	  ; [ 4] - Date at which this entry was received to be written to the log
	  ; [ 5] - Time at which this entry was received to be written to the log
	  ; [ 6] - Event identifier / Event ID / Event Code
	  ; [10] - Event source
	  ; [13] - Event description
	  If ((Int($aEvent[6]) = 12) And ($aEvent[10] = "Microsoft-Windows-Kernel-General")) Then
		 Local $sBootTimestamp = GetEventTimestamp($aEvent[4], $aEvent[5])
		 DisplayUptime($sBootTimestamp)
		 ExitLoop
	  ElseIf ((Int($aEvent[6]) = 6013) And ($aEvent[10] = "EventLog")) Then
		 ; This event logs "The system uptime is 1255308 seconds."
		 Local $sEventTimestamp = GetEventTimestamp($aEvent[4], $aEvent[5])

		 ; Extract the uptime in seconds from the event body.
		 Local $aUptime = StringRegExp($aEvent[13], "(\d+)", $STR_REGEXPARRAYMATCH)
		 If IsArray($aUptime) Then
			; Subtract the seconds from the event timestamp.
			Local $sBootTimestamp = _DateAdd("s", -$aUptime[0], $sEventTimestamp)
			DisplayUptime($sBootTimestamp)
			ExitLoop
		 EndIf
	  EndIf

	  $aEvent = _EventLog__Read($hEventLog, True, False) ; read last event
   WEnd

   _EventLog__Close($hEventLog)
EndFunc   ;==>Uptime

