; Eric Ginn Removal

;;;Error Check;;;
;ERRORS
HotKeySet("{ESC}","Terminate")
Func Terminate()
   Exit
EndFunc

If @error = True Then
   MsgBox(0,"Error","General Error")
   Exit
EndIf

;;;;FUNCTIONS;;;;

;Calls User Input for a name and returns the length of this name
func name_length(); Not used
$sname = InputBox("Name Removal","What is the string to remove from folder?", "Insert Removal here ")

$inameLength = StringLen($sname)
Return $inameLength

EndFunc

;Inputs are key to simulate press and number of times to simulate pressing it
func sendKey($sKey,$inumKeyPress)
		 for $k = 0 to $inumKeyPress step 1
		 Send("{" & $sKey &"}") ; press backspace that many times
		 sleep(100)
		 Next
EndFunc

;Inputs are the length of text string and it returns a copy of text that is that long - used in error checking
func getText($ilengthOfText) ; Gets text highlighted a set number of letters from current cursor position and returns it
   Send("{SHIFTDOWN}")
   For $j = 0 to $ilengthOfText-1 step 1
	  Send("{RIGHT}")
   Next
   Send("{SHIFTUP}")
	  Sleep(100)
   Send("^c")
	  Sleep(100)
   $sclip = ClipGet()
   Return $sclip
EndFunc

;Same as getText but only highlights it
Func highlightText($ilengthOfText) ; same as getText but only highlights it
   Send("{SHIFTDOWN}")
   For $j = 0 to $ilengthOfText-1 step 1
	  Send("{RIGHT}")
   Next
   Send("{SHIFTUP}")
	  Sleep(100)

	  Sleep(100)
EndFunc

;Compares two text strings and returns "True" if they are the same, "False" if they are not
Func compareText($stext1, $stext2);compares two strings and returns "true" if they are equal and "False" if they are not. not case sensitive
   if $stext1 = $stext2 Then ; checks that the highlighted text is equal
		 Return True;
	  Else
		 Return False;
   EndIf

EndFunc

;Activates renaming of a folder and ensures in the leftmost quadrant. Input is how many milliseconds to wait before executing next line (after function)
Func activateRenamingLeft($isleepTime); activates renaming by pushing F2 and going all the way left.
   Sleep($isleepTime);1second delay
   Send("{F2}") ;activate renaming
   Sleep($isleepTime)
   Send("{LEFT 500}") ; 500 should be the most necessary
   Sleep($isleepTime)
EndFunc

;;;;MAIN FUNCTION

;Inputs are number of loops (numberToDelete), the window handle that the folder applies to, and the prefix to be deleted. Some issues with getting it not restart at the top of the folders if error found
func delete_prefix($sprefixToDelete,$inumberToDelete,$hWnd)
   $ilengthPrefix = StringLen($sprefixToDelete);

   $fmsgbox = MsgBox(1, "Prepare to Run delete_prefix function", "Prepared to run with " & $ilengthPrefix & " characters deleted from prefix and " & $inumberToDelete & " folders to delete this prefix?") ;verification

   ;If user clicks cancel, exit program
   if $fmsgbox = 2 Then
	  MsgBox(0,"Error", "Select first folder to rename")
	  Exit
   EndIf

   $CounterFail = 0;checks the number of failures

   ;Loops through all folders and deletes prefix if it exists
   for $i = 0 to $inumberToDelete-1 Step 1

	  WinActivate(HWnd($hWnd));activate window by handle
	  Sleep(500);1second delay


	  activateRenamingLeft(50);starts renaming

	  $stextCheck = getText($ilengthPrefix) ;goes from current cursor position and hightlights $ilengthPrefix characters right of it
	  $ftextCheck = compareText($stextCheck, $sprefixToDelete)

	  ;Checks if the text identified as a suffix is the correct string
	  if $ftextCheck = False Then ; checks that the highlighted text is equal
		 MsgBox(0,"Error","" & $stextCheck & " does not equal " & $sprefixToDelete); Gives warning that the current folder doesn't have the suffix
		 send("{ENTER}");
		 WinActivate(HWnd($hWnd));activate window by handle
		 Sleep(100)
		 Send("{Down}")

			$CounterFail = $CounterFail + 1 ; adds one to counter fail
			;Checks to see if there have been multiple failures in a row.
			If $CounterFail > 3 Then ; checks to see if in a bad loop
			   MsgBox(0,"In wrong location","Sending 'n'");TEST
			EndIf
		 ContinueLoop;
	  EndIf

	  activateRenamingLeft(100);starts renaming
	  highlightText($ilengthPrefix) ; highlights the desired text (assuming it is a prefix and all the way in the left)
	  Sleep(100)
	  Send("{DEL}");Deletes text highlighted
	  Sleep(100)
	  Send("{Enter}")
	  Sleep(100)
	  $CounterFail = 0
	  Send("n") ; goes back to the next "n"

   Next

EndFunc

;;;;SCRIPT;;;;

;;USER INPUTS <---- CHANGE HERE, especially handle window (Use AutoIT Window Information tool to find Handle);;
$inumFolders = InputBox("Folder Information", "How many folders in a row?", "7");
$sprefixToDelete1 = InputBox("Prefix Information", "What is prefix to remove?","NEWCO - ");
;Change handle to match example below - error here
MsgBox(0,"Edit Code", "Please go into source code (line 142) and see user inputs to set the window handle and edit out this warning message") ; Warning to edit
$hWnd1 = "0x00090512"; hopefully still the window of others PLACEHOLDER for clicking to select the handle to use


;Codes
delete_prefix($sprefixToDelete1,$inumFolders,$hWnd1)

MsgBox(0,"Completed AutoIT Script 'Folder Renaming'.","DONE")
Exit