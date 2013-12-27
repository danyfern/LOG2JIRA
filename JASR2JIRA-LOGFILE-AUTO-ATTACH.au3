; AutoIT Script to automate the upload of JASR log files to a JIRA while reporting an ASR/SCI
; Preconditions:
; (only) the JIRA in scope needs to be opened in IE
; one single asr/sr and one single voicepad/sr log is opened in JASR
; have the device connected and MyMobiler running with the Screen to upload
; Heinrich Krupp, Nuance, December 2013

Global $path = "C:\Temp\", $filename ="", $action="", $hWnd = WinWait("Device Screen Capture", "", 10)

#include <Clipboard.au3>
;#include <MsgBoxConstants.au3>
ClipPut("")

Func changefilepath()
    ClipPut("")
	Send ( "^a") ;highlight filename
	Send ( "^c") ; copy filename to clipboard
	Send ( "^c") ; copy filename to clipboard
	$filename = ClipGet()
	;copy the filename and concatenate with new path, then put it back into the  field
	ClipPut($path & $filename)
	;MsgBox($MB_SYSTEMMODAL, "Title", "Clipboard Content: " & ClipGet() & ".", 10) ;show pop up with desired text.
	Send ( "^v") ; paste clipboard (new pathname/filename)
EndFunc

Func openuploaddialog(); JIRA Upload dialog
	if Not WinExists( "Choose File to Upload")Then
		;WinActivate ( "Choose File to Upload") ; go to file save dialog
		;WinWait ( "Choose File to Upload")
		Send("{.}")
		Sleep(1000)
		Send("attach files")
		Sleep(1000)
		Send("{ENTER}")
		Sleep(1000)
	Endif
EndFunc


; main procedure

if WinExists ("[CLASS:IEFrame]", "[#") Then  ; JIRA Window exists in IE
	WinActivate("[CLASS:IEFrame]", "[#")
	;WinActivate("[CLASS:IEFrame]", "Attach Files:")

	;Exit
	;Send ( "!d") ; higlight url in IE
	;Send("^c") ;Copy selected text.
	;$url = ClipGet()
	;MsgBox(0, "Url is:", $url)
	;WinActivate ( "[#" )
	;Local $title = WinGetTitle("[#", "")
	;MsgBox(0, "Full title read was:", $title)

	ASR2JIRA()

	VOICEPAD2JIRA()

	MYMOBILERSCREENSHOT2JIRA()

	DDMSCREENSHOT2JIRA()

	if WinExists( "Choose File to Upload")Then
		Send ( "!s") ; attach files
		WinWait("[CLASS:IEFrame]", "Attach Files:")
	EndIf

	WinActivate ( "[#" )
	;WinActivate ( $title )

EndIf ; of JIRA Window exists  in IE

Func attachlog(); in JIRA Upload dialog
	Send("{TAB}");Tab
	Sleep(500)
	Send("{SPACE}");Space
	WinWait ( "Choose File to Upload")
	Sleep(500)
EndFunc

Func FILE2JIRA()
	Sleep(500)
	if 	$action = "JASR" then
		Send ( "!fs") ;save JASR log file`
		changefilepath()
		Send ( "!o") ; save file
	EndIf
	;attach log file to JIRA
	WinActivate ( "Choose File to Upload") ; go to file save dialog
	WinWait ( "Choose File to Upload")
	Send ( "!d") ;go to address bar
	Send($path)
	;Send ( "^v") ; paste file name out of clipboard
	Send("{ENTER}") ; open the path
	Send ( "!n") ;go file name
	ClipPut($filename)
;	Send ( $filename)
	Send ( "^v") ; paste filename out of clipboard
	Send ( "!o") ; open (upload) file
	Sleep(2000)
	WinWait("[CLASS:IEFrame]", "Attach Files:")
EndFunc


Func MYMOBILERSCREENSHOT2JIRA()
	$action = "MYMOBILER"
	_ClipBoard_Empty()
	if WinExists ("MyMobiler") Then  ; MyMobiler Window exists
		openuploaddialog()
		attachlog()
		WinActivate ( "MyMobiler" )
		WinWait ( "MyMobiler" )
		;Send ( "^s"); screenshot to clipboard
		Send ( "^f"); screenshot to file
		Send ( "^a") ;highlight filename
		Send ( "^c") ; copy filename to clipboard
		Send ( "^c") ; copy filename to clipboard
		$filename = ClipGet()
		MsgBox(0, "Clipboard contains:", $filename)
		WinWait ( "MyMobiler" )
		Send ( "!d"); go to path
		Send ( "^a")
		Send ( $path) ;new pathname
		;changefilepath()
		Send ( "!s") ; save file
		FILE2JIRA()
	EndIf
EndFunc

Func DDMSCREENSHOT2JIRA()
	$action = "DDM"
	ClipPut("")
	if WinExists ("Device Screen Capture") Then  ; Device Screen Capture Window exists having the desired screen
		openuploaddialog()
		attachlog()
		doddmscreenshot()
	Else
		IF WinExists("Dalvik Debug Monitor") Then ; open the screenshot captureing window and refresh the screen
			openuploaddialog()
			attachlog()
			WinActivate ( "Dalvik Debug Monitor" )
			WinWait ( "Dalvik Debug Monitor" )
			Send ( "^s") ;screen capture
			WinActivate ( "Device Screen Capture" )
			ControlClick($hWnd, "", "Refresh"); refresh screen capture before saving it
			Sleep(5000)
			doddmscreenshot()
		EndIf
	EndIf
	
	
EndFunc


func doddmscreenshot()
		_ClipBoard_Empty()
		WinActivate ( "Device Screen Capture" )
		WinWait ( "Device Screen Capture" )
		ControlClick($hWnd, "", "Save"); files save dialog
		Send ( "^a") ;highlight filename
		Send ( "^c") ; copy filename to clipboard
		$filename = ClipGet()
		changefilepath()
		Send ( "!s") ; save file
		FILE2JIRA()
EndFunc

Func VOICEPAD2JIRA()
	$action = "JASR"
	if WinExists ("voicepad/sr") Then  ; "voicepad/sr" Window exists
		openuploaddialog()
		attachlog()
		WinActivate ( "voicepad/sr" )
		WinWait ( "voicepad/sr" )
		Send ( "!fs") ;save JASR log file`
		changefilepath()
		Send ( "!o") ; save file
		FILE2JIRA()
		;WinClose( "voicepad/sr" )
	Endif
EndFunc

Func ASR2JIRA()
	$action = "JASR"
	if WinExists ("asr/sr") Then  ; "asr/sr" Window exists
		openuploaddialog()
		attachlog()
		WinActivate ( "asr/sr" )
		WinWait ( "asr/sr" )
		Send ( "!fs") ;save JASR log file`
		changefilepath()
		Send ( "!o") ; save file
		FILE2JIRA()
		;WinClose( "asr/sr" )
	Endif
EndFunc