Startup:         ;MS done
{
validhelpers=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890
letters=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz
numbers=1234567890
mildCodes=,cnmild,cnmildr,cnmi,jnmild,jnmildr,jnmi,jnfs,cnfs,lcnmi,ljnmi,nljn,nlcn,jnami,cnami,

#SingleInstance force
#WinActivateForce
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetTitleMatchMode 1
CoordMode, Mouse, Relative
ComObjError(false)

IfExist, %A_MyDocuments%\CarisCodeRocket.ini
	ReadIniValues()
Else
	FirstTimeSetup()

SetTimer, checkForMelanoma, 3600000

Loop Files, %programfiles%\Microsoft Office\*.* , D
	{
	IfInString, A_LoopFileName, Office
		IfExist, %programfiles%\Microsoft Office\%A_LoopFileName%\OUTLOOK.EXE
			OutlookPath= %programfiles%\Microsoft Office\%A_LoopFileName%\OUTLOOK.EXE
		else 
		{
			np := RegExReplace(programfiles, " \(x86\)","")
			IfExist, %np%\Microsoft Office\%A_LoopFileName%\OUTLOOK.EXE
				OutlookPath= %np%\Microsoft Office\%A_LoopFileName%\OUTLOOK.EXE
			else
				Msgbox, Your path to OUTLOOK.EXE could not be determined.  You will not have automatic emailing capabilities.

		}
    }

Gosub, BuildMainGui

Gui, 2:Font, S12, Verdana
Gui, 2:Add, Text, x18 vDocPreferenceLabel, TestNameXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Gui, 2:Font, S8, Verdana
Gui, 2:Add, Checkbox, vPhotoSelect, Photos Required
Gui, 2:Add, Checkbox, vMicroSelect, Micros Required
;Gui, 2:Add, Checkbox, vICD9Select, ICD9s Required
Gui, 2:Add, Text, , Margin Preferences
Gui, 2:Add, Edit, vMarginSelect w500, 
Gui, 2:Add, Button, gSavePreferences vGo, Save Preferences

Gui, 4:Font, S12, Verdana
Gui, 4:Add, Text, vDisplayCaseNumber , Case Number: XXXXXXXXXXX
Gui, 4:Add, DropDownList, vLoc w500, Boston|Irving
Gui, 4:Add, DropDownList, AltSubmit vEmailType w500, Patient Double Blind Error|Need Previous Biopsy Report|Clinical Note and Photos|Critical Result Call|Pull Bottles and Blocks
GuiControl, 4:Choose, EmailType, 1
Gui, 4:Add, Text, ,Comments
Gui, 4:Add, Edit, vEmailComments w500, 
Gui, 4:Add, Button, Default gSendEmail vSendEmail, Send Email

; This Gui Window (5) is only to hold the ListView.  It is not to be shown.
Gui, 5:Add, ListView, Hidden w600 r5 gMyListView, ID|Code|Category|Subcategory|Dx Line|Comment|Micro|CPT Code|ICD9|ICD10|SNOMED|Premalignant|Malignant|Dysplastic|Melanocytic|Inflammatory|MarginIncluded|Log


Progress, 0 x400 y1 h130, Preparing for first time use..., Written by Matthew Muenster M.D.`n`nInitializing..., Caris CodeRocket 
Progress, 40, Reading personalized values...
Progress, 60, Getting the diagnosis codes from the database...
Gui, 5:Default
ReadDXCodes()
Gui, 1:Default

Progress, 80, Getting the helper codes from the database...
ReadHelpers()
Progress, 100, Initialization complete!
Progress, Off
{
	df =FRONT OF DIAGNOSIS HELPERS`n
	df=%df%------------------------------------------------------`n
	Loop, 10
		{
			x := A_index -1 
			ph1 := FrontofDiagnosisHelper%x%
			df = %df%%x% -- %ph1%`n
	}
	Loop, 26
		{
			x := Chr(64 + A_Index)
			ph1 := FrontofDiagnosisHelper%x%
			df = %df%%x% -- %ph1%`n
		}
	df = %df%----------------------------------------------------`n

	dm =MARGIN HELPERS`n---------------------------------------------------------------------------------`n
	Loop, 10
		{
			x := A_index -1 
			ph1 := BackofDiagnosisHelper%x%
			dm = %dm%%x% -- %ph1%`n
	}
	Loop, 26
		{
			x := Chr(64 + A_Index)
			ph1 := BackofDiagnosisHelper%x%
			dm = %dm%%x% -- %ph1%`n
		}
	dm = %dm%---------------------------------------------------------------------------------`n

dc =COMMENT HELPERS`n---------------------------------------------------------------------------------`n
	Loop, 10
		{
			x := A_index -1 
			ph1 := CommentHelper%x%
			dc = %dc%%x% - %ph1%`n
	}
	Loop, 26
		{
			x := Chr(64 + A_Index)
			ph1 := CommentHelper%x%
			dc = %dc%%x% - %ph1%`n
		}

	dc = %dc%---------------------------------------------------------------------------------`n
Gui, 3:Font, S8, Verdana
Gui, 3:Add, Text, w150, %df%
Gui, 3:Add, Text, w100 ym, %dm%
Gui, 3:Add, Text, w200 ym, %dc%
}
if OpMode=T
	{
		SplashTextOn, 100,100,,Entering Windowless Operation Mode
		Gosub, ModeTranscription
		Sleep, 1200
		SplashTextOff
	}
Else
	{	
		SetTimer, WinSURGECaseDataUpdater, 2000
		Gosub, WinSurgeCaseDataUpdater
		If OpMode=D
			Gosub, ModeDeluxeAutomatic
		else if OpMode=B
			Gosub, ModeBasicAutomatic
		Else
			Gosub, ModeManual
	}
If SpeakEnabled
	Menu, SettingsMenu, Check, Speak Patient Name
Else
	Menu, SettingsMenu, UnCheck, Speak Patient Name

If UseSendMethod
	Menu, SettingsMenu, Check, Use Send Method
Else
	Menu, SettingsMenu, UnCheck, Use Send Method

If BeepOnShiftEnter
	Menu, SettingsMenu, Check, Beep On Shift-Enter
Else
	Menu, SettingsMenu, UnCheck, Beep On Shift-Enter

StringLeft, x, A_ScriptDir, 1
if x=C
	x=S
Run, %x%:\CodeRocket\bin\EP\Autohotkey.exe %x%:\CodeRocket\bin\EP\DermpathExtendedPhrases.ahk, ,UseErrorLevel , epid
If ErrorLevel
	MsgBox, 4112, Connectivity Error, Your connection to the S:\CodeRocket directory is not present.  Usually`, restarting your computer will correct this.  You can use the CodeRocket program but will have no extended phrase capabilities.


IfExist, %A_MyDocuments%\PersonalExtendedPhrases.ahk
	{
	Run, %x%:\CodeRocket\bin\EP\Autohotkey.exe "%A_MyDocuments%\PersonalExtendedPhrases.ahk", ,UseErrorLevel , ppid	
	If ErrorLevel
		Msgbox, There was an error loading your personal extended phrases file.
	}
return
}

BuildMainGui:
{
	preferences := ""
	if (z1="")
		z1 := "WinSURGE Not Open"
 	
 		x := get_filled_case_number(z1)
			
		s := "select s.dx, s.gross, s.numberofspecimenparts, s.custom03, s.clin, p.name, s.clindata, pt.name, s.Computed_PATIENTAGE, p.proficiencylog, p.comment, s.custom04, s.patient from specimen s, physician p, patient pt where s.patient = pt.id and s.clin=p.id and computed_numberfilled='" . x . "'"
		WinSurgeQuery(s)

		CurrentCaseNumber := z1
		finaldiagnosistext := Result_1
		grossdescriptiontext := Result_2
		numberofvials := Result_3
		OrderedCPTCodes := Result_4
		ClientWinSurgeId := Result_5
		ClientName := Result_6		
		ClinicalData := Result_7
		PatientName := Result_8
		StringSplit, j, Result_9, .
		PatientAge := j1
		preferences := Result_10
		ClientOfficeName := Result_11
		ClientID := Result_12
		StringSplit, PatientID, Result_13, .    ; The patient ID is in a variable called PatientID1
		StringReplace, ClientID, ClientID, `n,,All
		
		StringReplace, preferences, preferences, ¥, ,All
		StringReplace, preferences, preferences, ***,`n, All
		StringReplace, preferences, preferences, `r,,All

		preferences := CleanText(preferences)		

		StringReplace, ClientOfficeName, ClientOfficeName, -Att, ,
		
		STringReplace, finaldiagnosistext, finaldiagnosistext, `%`%P`%`%%A_Space%, `%`%P`%`%`n, All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%B., `n`nB., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%C., `nC., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%D., `nD., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%E., `nE., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%F., `nF., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%G., `nG., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%H., `nH., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%I., `nI., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%Comment:, `n`nComment:, All
		finaldiagnosistext := RegExReplace(finaldiagnosistext,"%%\w+%%","")
		

		STringReplace, ClinicalData, ClinicalData, %A_Space%B., `nB., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%C., `nC., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%D., `nD., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%E., `nE., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%F., `nF., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%G., `nG., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%H., `nH., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%I., `nI., All

		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%B., `nB., All
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%C., `nC., All
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%D., `nD., All
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%E., `nE., All
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%F., `nF., All
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%G., `nG., All
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%H., `nH., All
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%I., `nI., All

		helpText =%PatientName%`n%PatientAge% years`n`n%ClientName%`n%ClientOfficeName%`n`nPREFERENCES%additionalPreferences%%preferences%`n`nCLINICAL INDICATIONS/HISTORY`n%ClinicalData%`n`nFINAL DIAGNOSIS`n%finaldiagnosistext%`n`nGROSS DESCRIPTION`n%grossdescriptiontext%

		ReDrawGui()
		
		additionalPreferences := ""
		s=select p.proficiencylog from physician p where p.number='%ClientID%'
		Msgbox, %s%
		WinSUrgeQuery(s)

		StringReplace, msg, msg, ¥, ,All
		StringReplace, msg, msg, ***,`n, All
		StringReplace, msg, msg, `r,,All

		msg := CleanText(msg)		
		;Msgbox, msg=`n%msg%`npreferences=`n%preferences%
		If msg
			{
			FoundPos1 := InStr(msg, preferences)
			FoundPos2 := InStr(preferences, msg)
			;Msgbox, %FoundPos1%,%FoundPos2%
			if(FoundPos1 OR FoundPos2)
				additionalPreferences := ""
			else
				{
				StringReplace, msg, msg,¥,,All
				additionalPreferences := msg
				}	
			}

	if (additionalPreferences)
		{
		preferences=%additionalPreferences%`n%preferences%
		preferences := CleanText(preferences)
		RedrawGui()
		}
		
		
	priorCaseInfo:=""
	perc := "%%"
	msg := ""
	if(PatientID1)
	{
		s := "select s.number, s.acdate, s.sodate, p.name, s.dx from specimen s, physician p where s.patient=" . PatientID1 . " and s.path=p.id"
		WinSurgeQuery(s)
	}

	if (msg)
	{
			Loop, Parse, msg, `n
			{
					If A_LoopField
					{
						FoundPos := RegExMatch(A_LoopField, "¥[A-Z][A-Z]\d\d-\d+¥")
						If FoundPos>0
							priorCaseLine=`n%A_LoopField%
						else
							priorCaseLine=%A_LoopField%
						
						StringReplace, priorCaseLine, priorCaseLine, %perc%P%perc%,`n, All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88305%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88304%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88312%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88346%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88350%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88342%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, Comment:, `n`nComment: , All
						StringReplace, priorCaseLine, priorCaseLine, ¥A., `nA., All
						StringReplace, priorCaseLine, priorCaseLine, ¥, %A_Space%, All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%B., `n`nB., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%C., `n`nC., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%D., `n`nD., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%E., `n`nE., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%F., `n`nF., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%G., `n`nG., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%H., `n`nH., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%I., `n`nI., All
						
						priorCaseInfo = %priorCaseInfo%`n%priorCaseLine%
					}
			}
		}
		else
		{
			GuiControl, Disable, PriorCases
		}


		
return
}

WM_MOUSEMOVE()
{
    static CurrControl, PrevControl, _TT  ; _TT is kept blank for use by the ToolTip command below.
    CurrControl := A_GuiControl
    If (CurrControl <> PrevControl and not InStr(CurrControl, " "))
    {
        ToolTip  ; Turn off any previous tooltip.
        SetTimer, DisplayToolTip, 1000
        PrevControl := CurrControl
    }
    return

    DisplayToolTip:
    SetTimer, DisplayToolTip, Off
    ToolTip % %CurrControl%_TT  ; The leading percent sign tell it to use an expression.
    SetTimer, RemoveToolTip, 3000
    return

    RemoveToolTip:
    SetTimer, RemoveToolTip, Off
    ToolTip
    return
}

CleanText(txt)
{
	newText := ""
	txt := Trim(txt)
	StringReplace, txt, txt, `r,,All
	Loop, Parse, txt, `n
	{
		j := Trim(A_LoopField)
		If (j)
			newText=%newText%%j%`n
	}
	
	return newText
}
	
ReDrawGui()
{
	global

	if(z1="No Current Case" OR z1="")
	{
		
		Gui, Destroy

		Gui, Add, Text, vCaseLoaderLbl,Case Loader:
		Gui, Add, Edit, vCaseScanBox ys,
		Gui, Font, S14, FixedSys
		Gui, Add, Text, cBlue r1 w400 xs vCaseNumberLabel, No Current Case Open
		Gui, Add, Button, w0 h0 Default, OK
	}
	else if (z1="WinSURGE Not Open")
	{
		
		Gui, Destroy

		Gui, Add, Text, vCaseLoaderLbl,Case Loader:
		Gui, Add, Edit, vCaseScanBox ys,
		Gui, Font, S14, FixedSys
		Gui, Add, Text, cBlue r1 w400 xs vCaseNumberLabel, WinSURGE Not Open
		Gui, Add, Button, w0 h0 Default, OK
		GuiControl, Disable, CaseScanBox
		GuiControl, Disable, CaseLoaderLbl
		GuiControl, 1:Hide, UseMicros
		GuiControl, 1:Hide, UsePhotos

	}
	else
	{
		Gui, Destroy
		Gui, Font, S12, Arial

		Gui, Add, Text, vCaseLoaderLbl,Case Loader:
		Gui, Add, Edit, vCaseScanBox ys,
		Gui, Font, S16, Arial
		Gui, Add, Button, x350 y10 vPriorCases ,Prior Cases
		
		Gui, Font, S20, Arial
		Gui, Add, Text, x500 y0 r1 cBlue w450 vUsePhotos, <PHOTOS REQUIRED>
		Gui, Add, Text, r1 cRED vUseMicros, <MICROS REQUIRED>
		;Gui, Add, Text, r1 cBlack w500 vUseMargins, <Margin Preferences appear here>

		Gui, Font, S18, Arial
		Gui, Add, Text, cBlue y40 r1 w400 xs vCaseNumberLabel, Case Number: %CurrentCaseNumber%
		Gui, Font, S14, FixedSys
		Gui, Add, Text, r1 w400 vPatientLabel, Patient: %PatientName% --- Age:%PatientAge%
		Gui, Add, Text, r1 w400 vDoctorLabel, Doctor: %ClientName%
		Gui, Add, Text, r1 w400 vClientLabel, Client:  %ClientOfficeName%
		gui, add, text, x10 y+10 w800 h1 0x7  ;Horizontal Line > Black
		Gui, Font, S14, FixedSys
		Gui, Add, Text, w800 vPreferences, %preferences%
		gui, add, text, x10 y+10 w800 h1 0x7  ;Horizontal Line > Black
		Gui, Add, Text, r1 w400 cBlue vJarSiteLabel, Clinical information:
		Gui, Add, Text, w800 cBlue vJarSiteInformation, %ClinicalData%
		gui, add, text, x10 y+10 w800 h1 0x7  ;Horizontal Line > Black
		Gui, Add, text, w800 vFinalDiagnosisText, Final Diagnosis:`n`n%finaldiagnosistext%
		gui, add, text, x10 y+10 w800 h1 0x7  ;Horizontal Line > Black
		Gui, Add, Text, w800 vgrossDescription gGrossDescription, %grossdescriptiontext%
		grossDescription_TT = %grossdescriptiontext%
		gui, add, text, x10 y+10 w800 h1 0x7  ;Horizontal Line > Black
		;gui, add, text, x10 y+10 w800 h1 0x7  ;Horizontal Line > Black
		Gui, Add, Button, w0 h0 Default, OK
		OnMessage(0x200, "WM_MOUSEMOVE")
	}
	
Menu, FileMenu, Add, E&xit, GuiClose
Menu, HelpMenu, Add, Search Diagnosis Codes  (F7), F7
Menu, HelpMenu, Add, Search Extended Phrases  (Shift-F7), +F7
Menu, HelpMenu, Add, Display All Helpers  (F9), F9

Menu, ModeMenu, Add, Manual, ModeManual
Menu, ModeMenu, Add, Basic Automatic, ModeBasicAutomatic
Menu, ModeMenu, Add, Deluxe Automatic, ModeDeluxeAutomatic
Menu, SettingsMenu, Add, Speak Patient Name, SpeakPatientName
Menu, SettingsMenu, Add, Use Send Method, UseSendMethod
Menu, SettingsMenu, Add, Beep On Shift-Enter, BeepOnShiftEnter


Menu, EditMenu, Add, Edit Client Preferences, EditDoctorPreferences

Menu, MyMenuBar, Add, &File, :FileMenu  
Menu, MyMenuBar, Add, &Edit, :EditMenu
Menu, MyMenuBar, Add, &Mode, :ModeMenu
Menu, MyMenuBar, Add, &Settings, :SettingsMenu
Menu, MyMenuBar, Add, &Help, :HelpMenu
Gui, Menu, MyMenuBar

Gui, 1:+Resize +MinSize400x300 +MaxSize900x800
Gui, Show, x%CarisRocketWindowX% y%CarisRocketWindowY% w%CarisRocketWindowW% h%CarisRocketWindowH%

return
}

SpeakPatientName:
{
	Menu, SettingsMenu, ToggleCheck, Speak Patient Name
	If SpeakEnabled
		SpeakEnabled := 0
	Else
		SpeakEnabled := 1
	IniWrite, %SpeakEnabled%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, SpeakEnabled
Return
}

UseSendMethod:
{
	Menu, SettingsMenu, ToggleCheck, Use Send Method
	If UseSendMethod
		UseSendMethod := 0
	Else
		UseSendMethod := 1
	IniWrite, %UseSendMethod%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, UseSendMethod

Return
}

BeepOnShiftEnter:
{
	Menu, SettingsMenu, ToggleCheck, Beep On Shift-Enter
	If BeepOnShiftEnter
		BeepOnShiftEnter := 0
	else
		BeepOnShiftEnter := 1
	
	IniWrite, %BeepOnShiftEnter%,  %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, BeepOnShiftEnter
	
	return
}
	
ModeTranscription:
{
	Hotkey, F7, On
	Hotkey, +F7, On
	Hotkey, F8, Off
	Hotkey, F9, Off
	Hotkey, +F8, Off
	Hotkey, F12, Off
	Hotkey, ^!s, Off
	Hotkey, ^k, Off
	Hotkey, ^!p, Off
	Return
}

ModeManual:
{
	Menu, ModeMenu, Check, Manual
	Menu, ModeMenu, UnCheck, Basic Automatic	
	Menu, ModeMenu, UnCheck, Deluxe Automatic	
	Menu, EditMenu, Disable, Edit Client Preferences
	Menu, SettingsMenu, Disable, Speak Patient Name
	SpeakEnabled := 0
	
	GuiControl, Text, UsePhotos, MANUAL PREFERENCES
	GuiControl, 1:Show, UsePhotos
	GuiControl, 1:Hide, UseMicros
	;GuiControl, 1:Hide, UseICD9s
	GuiControl, Disable, CaseScanBox
	GuiControl, Disable, CaseLoaderLbl
	GuiControl, Text, CaseScanBox, Manual Mode
	GuiControl, 1:Text, PatientLabel, 
	GuiControl, 1:Text, DoctorLabel, 
	Hotkey, F8, Off
	Hotkey, +F8, Off
	Hotkey, F12, On
	Hotkey, ^!s, Off
	Hotkey, ^k, Off
	Hotkey, ^!p, Off
	OpMode=M
	IniWrite, %OpMode%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, OpMode
	Return
}

ModeBasicAutomatic:
{
	if (!QueueIntoBatchBox OR QueueIntoBatchBox=Error)
		{
			Msgbox, 4, , You have not setup Caris CodeRocket for automated operation.  Do you want to go through setup to enable automation?
			IfMsgBox, Yes
				{
					FileDelete, %A_MyDocuments%\CarisCodeRocket.ini
					Reload
					Return
				}
			Else
				Return
		}
	SpeakEnabled := 0
	Menu, ModeMenu, UnCheck, Manual
	Menu, ModeMenu, Check, Basic Automatic	
	Menu, ModeMenu, UnCheck, Deluxe Automatic	
	Menu, SettingsMenu, Disable, Speak Patient Name
	Menu, EditMenu, Disable, Edit Client Preferences
	Hotkey, F8, On
	Hotkey, +F8, On
	Hotkey, F12, On
	Hotkey, ^!s, On
	Hotkey, ^k, On	
	GuiControl, Text, UsePhotos, MANUAL PREFERENCES
	GuiControl, 1:Text, PatientLabel, 
	GuiControl, 1:Text, DoctorLabel, 
	GuiControl, 1:Show, UsePhotos
	GuiControl, 1:Hide, UseMicros
	;GuiControl, 1:Hide, UseICD9s
	GuiControl, Enable, CaseScanBox
	GuiControl, Enable, CaseLoaderLbl
	GuiControl, Text, CaseScanBox, 
	OpMode=B
	IniWrite, %OpMode%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, OpMode
Return
}

ModeDeluxeAutomatic:
{
	if (!QueueIntoBatchBox OR QueueIntoBatchBox=Error)
		{
			Msgbox, 4, , You have not setup Caris CodeRocket for automated operation.  Do you want to go through setup to enable automation?
			IfMsgBox, Yes
				{
					FileDelete, %A_MyDocuments%\CarisCodeRocket.ini
					Reload
					Return
				}
			Else
				Return
		}
		
	Menu, ModeMenu, UnCheck, Manual
	Menu, ModeMenu, UnCheck, Basic Automatic	
	Menu, ModeMenu, Check, Deluxe Automatic	
	OpMode=D
	Hotkey, F8, On
	Hotkey, +F8, On
	Hotkey, F12, On
	Hotkey, ^!s, On
	Hotkey, ^k, On	
	Hotkey, ^!p, On
	GuiControl, Text, UsePhotos, <PHOTOS REQUIRED>
	GuiControl, Enable, CaseScanBox
	GuiControl, Enable, CaseLoaderLbl
	GuiControl, Text, CaseScanBox, 
	Menu, EditMenu, Enable, Edit Client Preferences
	Menu, SettingsMenu, Enable, Speak Patient Name


	IniWrite, %OpMode%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, OpMode
	lastWinSURGEtitle=  ;Forces Case update to occur.
	Gosub, WinSurgeCaseDataUpdater
Return
}

MyListview:       ;MS done
return	

EditDoctorPreferences:
{
	if OpMode=D
	{
		GuiControl, 2:, PhotoSelect, %UsePhotos%
		GuiControl, 2:, MicroSelect, %UseMicros%
		;GuiControl, 2:, ICD9Select, %UseICD9s%	
		GuiControl, 2:, MarginSelect, %UseMargins%
		GuiControl, 2:Text, PhotoSelect, Photos Required
		GuiControl, 2:Text, MicroSelect, Micros Required
		;GuiControl, 2:Text, ICD9Select, ICD9s Required
		SetTimer, WinSURGECaseDataUpdater, Off
		Gui, 1:Hide
		Gui, 1:+Disabled
		GuiControl, 2:Text, DocPreferenceLabel, Enter preferences for %ClientName%
		;Gosub, GlobalPreferenceSelect
		Gui, 2:Show  
		Gui, 2:-Disabled
		Sleep, 500
		WinActivate, Caris CodeRocket
	}
	Else
		Msgbox, You must be in Deluxe Automatic Mode to use this feature.
	Return
}
	
SavePreferences:   ;MS done
{
Gui, 2:Submit
	
	StringReplace, y, WinSurgeFullName,','', All
	StringReplace, z, ClientName,','',All

If ClinicianFound
	s =	Update CliniPref SET Photo_pref='%PhotoSelect%', Micro_pref ='%MicroSelect%', Icd9_pref='0', Margin_pref='%MarginSelect%', log='%CurrentLog%;%y% %A_mm%-%A_DD%-%A_YYYY%' WHERE Winsurge_id = %ClientWinSurgeId%
Else
	s =	Insert into CliniPref (Winsurge_id,Name,Photo_pref,Micro_pref,Margin_pref,Icd9_pref,log) VALUES (%ClientWinSurgeId%,'%z%','%PhotoSelect%', '%MicroSelect%', '%MarginSelect%', '0', '%y% %A_mm%-%A_DD%-%A_YYYY%')

CodeDatabaseQuery(s) 

lastWinSURGEtitle = 
Gui, 2:Hide
Gui, 2:+Disabled
Gui, 1:-Disabled
Gosub, WinSURGECaseDataUpdater
Gui, 1:Show
WinActivate, WinSURGE - 
WinActivate, Caris CodeRocket
lastWinSURGEtitle=
SetTimer, WinSURGECaseDataUpdater, 2000

Return
}

ButtonPriorCases:
{
	;~ x := get_filled_case_number(CurrentCaseNumber)
	;~ priorCaseInfo:=""
	;~ perc := "%%"
	;~ s := "select s.patient from specimen s where computed_numberfilled='" . x . "'"
	;~ WinSurgeQuery(s)

	;~ StringSplit, patientID, Result_1, .
	
	;~ s := "select s.number, s.acdate, s.sodate, p.name, s.dx from specimen s, physician p where s.patient=" . patientId1 . " and s.path=p.id"
	;~ ;. ", physician p, patient pt where s.patient = pt.id and s.clin=p.id and computed_numberfilled='" . x . "'"
	;~ WinSurgeQuery(s)

	;~ Loop, Parse, msg, `n
	;~ {
			;~ If A_LoopField
			;~ {
				;~ FoundPos := RegExMatch(A_LoopField, "¥[A-Z][A-Z]\d\d-\d+¥")
				;~ If FoundPos>0
					;~ priorCaseLine=`n%A_LoopField%
				;~ else
					;~ priorCaseLine=%A_LoopField%
				
				;~ StringReplace, priorCaseLine, priorCaseLine, %perc%P%perc%,`n, All
				;~ StringReplace, priorCaseLine, priorCaseLine, %perc%88305%perc%, , All
				;~ StringReplace, priorCaseLine, priorCaseLine, %perc%88304%perc%, , All
				;~ StringReplace, priorCaseLine, priorCaseLine, %perc%88312%perc%, , All
				;~ StringReplace, priorCaseLine, priorCaseLine, %perc%88346%perc%, , All
				;~ StringReplace, priorCaseLine, priorCaseLine, %perc%88350%perc%, , All
				;~ StringReplace, priorCaseLine, priorCaseLine, %perc%88342%perc%, , All
				;~ StringReplace, priorCaseLine, priorCaseLine, Comment:, `n`nComment: , All
				;~ StringReplace, priorCaseLine, priorCaseLine, ¥A., `nA., All
				;~ StringReplace, priorCaseLine, priorCaseLine, ¥, %A_Space%, All
				;~ StringReplace, priorCaseLine, priorCaseLine, %A_Space%B., `n`nB., All
				;~ StringReplace, priorCaseLine, priorCaseLine, %A_Space%C., `n`nC., All
				;~ StringReplace, priorCaseLine, priorCaseLine, %A_Space%D., `n`nD., All
				;~ StringReplace, priorCaseLine, priorCaseLine, %A_Space%E., `n`nE., All
				;~ StringReplace, priorCaseLine, priorCaseLine, %A_Space%F., `n`nF., All
				;~ StringReplace, priorCaseLine, priorCaseLine, %A_Space%G., `n`nG., All
				;~ StringReplace, priorCaseLine, priorCaseLine, %A_Space%H., `n`nH., All
				;~ StringReplace, priorCaseLine, priorCaseLine, %A_Space%I., `n`nI., All
				
				;~ priorCaseInfo = %priorCaseInfo%`n%priorCaseLine%
			;~ }
	;~ }
	
	Msgbox, %priorCaseInfo%
	return
}

ButtonOK:
{
	Gui, Submit, NoHide
	If (UndoEnabled AND !DataEntered)
		{
			Msgbox, You must first save or undo the changes to the current case!
			Gosub, F12
			Return
		}
	CaseNumberProblem=0
	StringSplit, x, CaseScanBox, %A_Space%
	StringSplit, y, x1, -
	Stringlen, csbLength, y1
	if (csbLength<>4)
		CaseNumberProblem=1
	StringLeft, csbPrefix, y1, 2
	StringRight, j, csbPrefix, 1
	IfNotInString, letters, %j%
		CaseNumberProblem=1
	StringLeft, j, csbPrefix, 1
	IfNotInString, letters, %j%
		CaseNumberProblem=1
	StringRight, csbYear, y1, 2
	csbCaseNum := y2
	StringRight, j, csbYear, 1
	IfNotInString, numbers, %j%
		CaseNumberProblem=1
	StringLeft, j, csbYear, 1
	IfNotInString, numbers, %j%
		CaseNumberProblem=1
	StringLen, casenumlength, csbCaseNum
	Loop, %casenumlength%
	{
		StringMid, j, csbCaseNum, A_Index, 1
		IfNotInString, numbers, %j%
			CaseNumberProblem=1
	}

if CaseNumberProblem
		{
			Msgbox, You did not enter a valid case number!
			Gosub, F12
			Return
		}

NewCaseNum=%csbPrefix%%csbYear%-%csbCaseNum%

z1 := NewCaseNum
additionalPreferences := ""

Gosub, BuildMainGui

	If (DataEntered AND UndoEnabled)
		Gosub, F8
	
	If SaveError
		{
			Gosub, F12
			Return
		}
	

	CloseWinSURGEModalWindow("WinSURGE - Final Diagnosis:","","Close")

	SetTimer, UnblockInput, 5000
	BlockInput, On
	OpenCase(NewCaseNum)
	OpenFinalDiagnosisModal()
	ActivateNextTripleAsterisk() 
	BlockInput, Off
	SetTimer, UnblockInput, Off
	
	ControlSetText, ThunderRT6TextBox19, %NewCaseNum%, Dictation Data, 	
	
	Gosub, WinSURGECaseDataUpdater
	StringSplit, name, PatientName, `,
	p = %name2%%A_Space%%name1%
	If SpeakEnabled
		{
		Run, %A_ScriptDir%\wscript.exe "%A_ScriptDir%\speak.vbs" "%p%"	
		Sleep, 400
		Run, %A_ScriptDir%\wscript.exe "%A_ScriptDir%\speak.vbs" "Age %PatientAge%"
		}
	DataEntered := 0
	Sleep, 100
	gosub, F12
	WinActivate, WinSURGE , 

	Return
}

UnblockInput:
{
	BlockInput, Off
	return
}

GuiClose:
{
	Process, Close, %epid%, 
	Process, Close, %ppid%,
	Process, Close, Autohotkey.exe,
	
	If ErrorLevel
		Process, Close, %ErrorLevel%, 
	ExitApp	
	Return
}

WinSURGECaseDataUpdater:  
{
	;This section gets and saves the position of the CodeRocket Window if it has moved
		SetTitleMatchMode, 2
		WinGetPos, x, y, w, h, CodeRocket, , SciTE4AutoHotkey
		SetTitleMatchMode, 1
		
		If(w>400 AND h>400)
			{
			if((x<>CarisRocketWindowX OR y<>CarisRocketWindowY OR w<>CarisRocketWindowW OR h<>CarisRocketWindowH) AND x<>-32000)
				{
					CarisRocketWindowX := x	
					CarisRocketWindowY := y
					CarisRocketWindowW := w
					CarisRocketWindowH := h
			
					IniWrite, %CarisRocketWindowX%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowX
					IniWrite, %CarisRocketWindowY%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowY
					IniWrite, %CarisRocketWindowW%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowW
					IniWrite, %CarisRocketWindowH%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowH
				}
			}

	
	ifWinNotExist, WinSURGE
		{
		lastWinSurgeTitle=""
		z1 := "WinSURGE Not Open"
		GuiControl, Disable, CaseScanBox
		GuiControl, Disable, CaseLoaderLbl
		GuiControl, Disable, ButtonOK
		GuiControl, 1:Hide, UseMicros
		GuiControl, 1:Hide, UsePhotos
		GuiControlGet, CaseNumberLabel
		If (CaseNumberLabel<>"WinSURGE Not Open")
			Gosub, BuildMainGui
			
		Return	
		}
	else
	{
			
		;Block for getting UNDO button status
		{  
		SetTitleMatchMode, 2
		ControlGet, UndoEnabled1, Enabled, ,&3 Undo, WinSURGE [, &2 Open Case
		ControlGet, UndoEnabled2, Enabled, ,&Undo, WinSURGE - Final Diagnosis:
		SetTitleMatchMode, 1
		If(UndoEnabled1 = 1 OR UndoEnabled2=1)
			UndoEnabled := 1
		Else
			UndoEnabled := 0

		if(UndoEnabled AND !DataEntered)
			{
			GuiControl, Disable, CaseScanBox
			GuiControl, Disable, CaseLoaderLbl
			GuiControl, Text, CaseScanBox, Data in Case
		}
		Else if (OpMode="B" OR OpMode="D")
		{
			Gui, Submit, NoHide
			GuiControl, Enable, CaseScanBox
			GuiControl, Enable, CaseLoaderLbl
			if CaseScanBox=Data in Case
				GuiControl, Text, CaseScanBox, 
		}
		}
			
		if (OpMode="M" OR OpMode="B")
			Return



		WinGetTitle, x, WinSURGE, &2 Open Case
		if (x=lastWinSURGEtitle)
			Return
		StringGetPos, y, x, No Current Case
		StringGetPos, z, x, New
		z := y + z

		If (z>0 OR x="WinSURGE")
			{
			lastWinSURGEtitle := x
			z1 := "No Current Case"
			Gosub, BuildMainGui	
			GuiControl, 1:Hide, UsePhotos
			GuiControl, 1:Hide, UseMicros
			Return
			}
		lastWinSURGEtitle := x
		DataEntered = 0
		StringReplace, x, x, Case, |, All
		StringSplit, y, x, |, %A_Space%
		StringSplit, z, y2, %A_Space%, %A_space%

		additionalPreferences := ""
		Gosub, BuildMainGui

		if OpMode=D
		{
			GuiControl, 1:Hide, UsePhotos
			if (ClientWinSurgeId="")
				Return
				
			s = Select top 1 c.name,c.photo_pref,c.micro_pref,c.margin_pref, c.icd9_pref, c.log from clinipref c where c.WinSurge_id=%ClientWinSurgeId%
			CodeDatabaseQuery(s)
			If Result_1  ;If a database entry was found
				{
					ClinicianFound = 1
					CurrentLog := Result_6
					if Result_2  ;Photos Selected\
						{
						GuiControl, 1:Show, UsePhotos
						UsePhotos := 1
					}
					Else	
						{
						GuiControl, 1:Hide, UsePhotos
						UsePhotos := 0
					}

					If Result_3
						{
						GuiControl, 1:Show, UseMicros	
						UseMicros := 1
					}	
					Else	
					{
						GuiControl, 1:Hide, UseMicros
						UseMicros := 0
					}

					If Result_4
						{
						GuiControl, Text, UseMargins ,% Result_4
						UseMargins := Result_4
						}
					Else	
					{
						GuiControl, Text, UseMargins ,
						UseMargins := ""
					}

					If Result_5
						{
						;GuiControl, 1:Show, UseICD9s	
						;UseICD9s := 1
						}
					Else	
					{
						;GuiControl, 1:Hide, UseICD9s
						;UseICD9s := 0
					}
					
				}
			Else    ;Else doctor was not found in the clinipref database
				{
					ClinicianFound = 0
					Gui, 1:Hide
					Gui, 1:+Disabled
					GuiControl, 2:, PhotoSelect, 0
					GuiControl, 2:, MicroSelect, 0
					;GuiControl, 2:, ICD9Select, 0
					GuiControl, 2:Text, DocPreferenceLabel, Enter preferences for %ClientName%
					Gui, 2:Show, , CarisDemo			;x%CarisRocketWindowX% y%CarisRocketWindowY% w%CarisRocketWindowW% h%CarisRocketWindowH%
					Gui, 2:-Disabled +AlwaysOnTop
					SoundBeep
					SoundBeep

					Return
				}
		}

		;~ s := " SELECT p.number as clientID  FROM specimen s left join physician p on s.client=p.id where computed_numberfilled = '" . x . "'"
		;~ WinSurgeQuery(s)
		;~ ClientID := Result_1
		;~ StringReplace, ClientID, ClientID, `r,,All
		;~ StringReplace, ClientID, ClientID, `n,,All

		Return
}
}

F7::            ;MS done
{
	InputBox, SearchWord, Diagnostic Code Search, Enter the single word or phrase you want to search for:
	if SearchWord
		{
	
	dt =Diagosis Code Search TABLE`n---Searched for   "%SearchWord%"  ---`n
	counter := 0
	pages := 1
Gui, 5:Default
	Loop % LV_GetCount()
		{
		LV_GetText(RetrievedText1, A_Index, 2)
		LV_GetText(RetrievedText2, A_Index, 5)
		LV_GetText(RetrievedText3, A_Index, 6)
		LV_GetText(RetrievedText4, A_Index, 7)
		StringGetPos, t1, RetrievedText1, %SearchWord%
		StringGetPos, t2, RetrievedText2, %SearchWord%
		StringGetPos, t3, RetrievedText3, %SearchWord%
		StringGetPos, t4, RetrievedText4, %SearchWord%

		if (t1>-1 or t2>-1 or t3>-1 or t4>-1)
			{
			counter := counter + 1
			dt = %dt%%RetrievedText1% `n%RetrievedText2%`n  	
			if (RetrievedText3 OR RetrievedText4)
				dt = %dt%Comment:  %RetrievedText3%  %RetrievedText4%`n`n
			if (Floor(counter/9) =  counter/9)
			{
				dt := dt . "¶"
				pages := pages + 1
			}

			}
		}

	Loop, Parse, dt, ¶
		{
			Gui, Font, Courier
			Msgbox, 4097, Diagnostic Code Search Results, % A_LoopField . "`nPage " . A_Index . " of " . pages	
				IfMsgbox, Cancel
					Break
			Gui, Font, Arial
		}

	}
	Gui, 1:Default
Return
}

+F7::
{
		InputBox, SearchWord, Extended Phrase Search, Enter the single word or phrase you want to search for:
	if SearchWord
		{
	
	dt =Extended Phrase Search`n---Searched for   "%SearchWord%"  ---`n
	dt = %dt%---------------------------------------------------------------------------------`n
	counter := 0
	pages := 1
	perc :="%"
	j='%perc%%SearchWord%%perc%'
	s = Select code,text from extendedphrases where text LIKE %j% OR code LIKE %j%
	
	connectstring := "DRIVER={SQL Server};SERVER=s-irv-sql02;DATABASE=winsurgehotkeys;uid=wshotkeys;pwd=hotkeys10;"
	adodb := ComObjCreate("ADODB.Connection")
	rs := ComObjCreate("ADODB.Recordset")
	rs.CursorType := "0"
	strRequest := s
	adodb.open(connectstring)
	rs := adodb.Execute(strRequest)
	If A_LastError
		{
			Msgbox,  There was an error accessing the Code Database.`n Try again later. `ns=%s%`n  A_LastError = %A_LastError%
			Return
		}

If (rs.EOF=0)
		{
		rs.MoveFirst()
		while rs.EOF = 0{
			DXCodeCount := A_Index
			j := 0
			for field in rs.fields
				{
				j := j + 1
				y := Field.Value
				DxCode%j%=%y%
				}
			dt = %dt%%DxCode1% -- %DxCode2%`n`n
			counter := counter + 1
			if (Floor(counter/12) =  counter/12)
			{
				dt := dt . "¶"
				pages := pages + 1
			}

			rs.MoveNext()
			}
	rs.close()   
	adodb.close()
		}
	dt = %dt%---------------------------------------------------------------------------------`n

	Loop, Parse, dt, ¶
		{
			Gui, Font, Courier
			Msgbox, 4097, Diagnostic Code Search Results, % A_LoopField . "`nPage " . A_Index . " of " . pages	
				IfMsgbox, Cancel
					Break
			Gui, Font, Arial
		}
		}
Return
}

F8::           
{

	SaveError = 0	
	ifWinExist, WinSURGE Case Lookup	
		Return
		
;CPT Code Checking Block
	{
		finaldiag := WinSURGEFinalDiagnosisContents()
		
		lowerCaseProblem := RegExMatch(finaldiag, "%%P%%\r\n[a-z]")
		if lowerCaseProblem>0
			{
			SoundBeep
			Msgbox,  STOP! There appears to be a lower case letter where the TOP LINE Diagnosis should be!
			Return
			}
			
		StringGetPos, i, finaldiag, ***
		if i>0
			{
				SaveError = 1
				Msgbox, The final diagnosis is missing critical information (there is a '***' in the box)!
				Return
			}
		
		StringGetPos, i, finaldiag, `%`%88305`%`%
		StringGetPos, j, finaldiag, `%`%88304`%`%
		StringGetPos, k, finaldiag, `%`%88321`%`%
		StringGetPos, l, finaldiag, `%`%88323`%`%
		StringGetPos, m, finaldiag, `%`%88300`%`%
		StringGetPos, n, finaldiag, `%`%NOCHG`%`%
		
		if (i<0 AND j<0 AND k<0 AND l<0 AND m<0 AND n<0)
			{
				SaveError = 1
				SoundBeep
				Msgbox,4,, The final diagnosis is missing a proper billing code!  Do you wish to continue without further editing?
				IfMsgbox, No
					Return
			}

		StringGetPos, i1, finaldiag, PAS
		StringGetPos, i2, finaldiag, schiff
		StringGetPos, i3, finaldiag, GMS
		StringGetPos, i4, finaldiag, gomori 
		if (i1>0 OR i2>0 OR i3>0 OR i4>0)
			{
				perc := "%%"
				k = %perc%88312%perc%
				StringGetPos, j, finaldiag, %k%
				If j=-1
					{
					SoundBeep
					Msgbox,4,,PAS or GMS stain mentioned without adding an 88312 billing code.  Do you wish to continue without code editing?
					IfMsgbox, No
						{
						SaveError=1					
						Return	
						}
					}
			}

		
		CloseWinSURGEModalWindow("WinSURGE - Final","","&Close")
	}

;Photo presence checking block
	if (OpMode="D" AND UsePhotos)
		{
			x := WinSURGEOpenCasePhoto1()
			y := WinSURGEOpenCasePhoto2()
			if (!x AND !y)
				{
				SoundBeep
				Msgbox, 4, NO PHOTOS WARNING, Client has requested photos and there are none. Do you want to continue without photos?	
				IfMsgBox No
					{
					SaveError = 1	
					Return
					}
				}
			}
	QueueandAssign()
	CloseandSaveCase()
	DataEntered = 0
	Gosub, F12
	Return
}

+F8::
{
	if OpMode<>M
	{
	CloseandSaveCase()	
	DataEntered = 0
	GuiControl, Hide, StatusLabel
	Gui, Show, NoActivate
	Gosub, F12
	}
	Return
}

F9::
{
Gui, 3:Font, 
Gui, 3:show, ,List of Available Helpers
Return
}

F11::
{
	IfWinExist, WinSURGE - 
	{
		Send, %LastCodeUsed%	
		Gosub, Shift & Enter
	}
	Return
}

F12::
{
	WinActivate, WinSURGE , 	
	WinActivate, CodeRocket
	if OpMode<>M
	{
		GuiControl, Text, CaseScanBox,  ;Blanks the data entry textbox 	
		GuiControl, Focus, CaseScanBox,
	}
	
return
}

^k::   ;Special Stain order
{
	
	CloseWinSURGEModalWindow("WinSURGE - Final","","&Close")
	CloseWinSURGEModalWindow("WinSURGE - Gross Description","","&Close")
	CloseWinSURGEModalWindow("WinSURGE Case Lookup","","Cancel")
	
	IfWinNotActive, WinSURGE [, &2 Open , WinActivate, WinSURGE [, &2 Open
	WinWaitActive, WinSURGE [, &2 Open
	Sleep, 400
	DataEntered = 0
	WinMenuSelectItem, WinSURGE [, &2 Open, Tools, Enter Special Stains via Checklist
	Return
}

^!s::  ;Batch Signout
{
	SetTimer, WinSURGECaseDataUpdater, Off

	Progress, x10 y10 h150, Preparing to signout, Obtaining Routine Cases for signout`n Press Ctrl-Alt-R to stop the signout, Working....,

s =	select s.number, s.zaudittraillast, s.yesno07 from specimen s, physician p where  s.sodate < '1950-01-01' and s.path = p.id and p.name ='%WinSurgeFullName%' and s.calculatedslidecountdate>'1950-01-01' order by 2 desc
	WinSurgeQuery(s)
	if !msg
		{
		Msgbox, There are no cases (not including amendments/addendums) to signout!	
		Progress, Off   ;Internal only
		SetTimer, WinSURGECaseDataUpdater, 1000
		Return
		}
	SignOutFileList = 
	SignOutCount := 0

Loop, Parse, msg, `n
		{
			LoopCaseNum = 
			Loop, Parse, A_LoopField, ¥
				{
				if A_Index =2	
					LoopCaseNum = %A_LoopField%
				if A_Index=3
					LoopYesNo07 = %A_LoopField%
				}
				if (LoopCaseNum<>"")
					{
					StringLeft, x, LoopCaseNum, 1
					StringRight, y, LoopCaseNum, 2
					;if ((x="C" OR y="MG") AND LoopYesNo07<>"Y")
						;Continue    ;Skip the C and MG cases where LoopYesNo07 is not set to Y
					SignoutFileList = %SignOutFileList%%LoopCaseNum%`n	
					SignOutCount := SignOutCount + 1
					}
		}

if (SignOutCount>0)
{
	;Progress, , Preparing to signout, Signing out the cases`n, Working....,

	Loop, parse, SignOutFileList, `n
	{
		y := SignOutCount - A_Index + 1
		x := 100 * (A_Index / SignOutCount)
		;Progress, X100 Y100 %x%,  %y% of %SignOutCount% cases remaining...

		if A_LoopField =  ; Omit the last linefeed (blank item) at the end of the list.
			break
		CloseWinSURGEModalWindow("WinSURGE Case Lookup","","Cancel")
		Open4SignOut(A_LoopField)
		
		z1 := A_LoopField
		Gosub, BuildMainGui
		
		If JumptoCaseNotSet
			{
				Msgbox, You must set the "Jump to Case Entered by User" in the WinSURGE User Preferences menu in order to use the automated signout functions!
				Progress, Off
				Return
			}

		If (!ApprovalPasswordControl OR ApprovalPasswordControl="ERROR")  ;Setup function to get the name of the Approval Password control the first time through	
		{
			BlockInput, On
			Sleep, 1000
			Loop,
			{
				Sleep, 500
				WinGetTitle, active_title, A
				ControlGetFocus, xxx, %active_title%
				StringGetPos, yyy, xxx, TextBox
			if (yyy>0)
				{
					ApprovalPasswordControl := xxx	
					IniWrite, %ApprovalPasswordControl%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, ApprovalPasswordControl
					break
				}
			Else
				{
					continue
				}
			}
			BlockInput, Off
			Sleep, 500
		}
			
			
		Progress, X100 Y100,  %y% of %SignOutCount% cases remaining...`n  Press F3 to approve this case.`nPress F4 to skip and go to the next one.`nPress 'End' to End signout loop.

	If (ExpressSignout=1)
		EnterSignoutPasswordandApprove()	
	Else
		{
			Loop,
				{
					Sleep, 40
					GetKeyState, F3state, F3
					GetKeyState, F4state, F4
					GetKeyState, EndState, End
					
					If F3state=D
						{
							EnterSignoutPasswordandApprove()	
							Break
						}
					If F4state=D
						{
							ControlGet, SkipEnabled, Enabled, ,&Skip, WinSURGE E-signout [, &Approve
							if SkipEnabled
								{
								SkipCaseSignout()
								Break
								}
							Else
								{
								Progress, Off   ;Internal only	
								SetTimer, WinSURGECaseDataUpdater, 2000
								Return
								}
						}
					If Endstate=D
						{
						Progress, Off   ;Internal only	
						SetTimer, WinSURGECaseDataUpdater, 2000
						Return
						}
				}
		}
}
Progress, Off   
SetTimer, WinSURGECaseDataUpdater, 2000
Return
}
Return
}

^!e::
{
GuiControl, 4:Text, DisplayCaseNumber, Case Number: %CurrentCaseNumber%
Gui, 4:Show
StringLeft, firstLetter, CurrentCaseNumber, 1
If (firstLetter="C")
	GuiControl, 4:Choose, Loc, Boston
else if (firstLetter="D")
	GuiControl, 4:Choose, Loc, Irving

GuiControl, 4:Focus, EmailType
return
}

SendEmail:
{
		StringSplit, parts, CurrentCaseNumber, -
		np2 := LTrim(parts2, "0")
		ccn=%parts1%-%np2%
	
	Gui, 4:Submit
	EmailComments := RegExReplace(EmailComments, " ", "`%20")
	
	if (Loc="Boston")
	{
		Email1 := "g_adminstaff@cohenderm.com"
		Email2 := "g_clientservices@cohenderm.com"
		Email3 := "g_clientservices@cohenderm.com"
		Email4 := "g_clientservices@cohenderm.com"
		Email5 := "BostonHistoSupsLeads@miracals.com"
	}
	else if (Loc="Irving")
	{
		Email1 := "IRVING-SPCDEPT@MiracaLS.com"
		Email2 := "PathologySupportClientServicesIRVPHX@MiracaLS.com"
		Email3 := "PathologySupportClientServicesIRVPHX@MiracaLS.com"
		Email4 := "PathologySupportClientServicesIRVPHX@MiracaLS.com"
		Email5 := "IrvingPA-Leads@MiracaLS.com"
	}
	else
		return
	
	if (EmailType=1)
		{
		SetTitleMatchMode, 2
		Run, "%OutlookPath%" /c ipm.note  /m %Email1%&subject=Patient`%20Double`%20Blind`%20Error`%20on`%20%ccn%&body=(%EmailComments%)
		WinWaitActive, (HTML)
		SetTitleMatchMode, 1
		Send, !s
		}

	else if (EmailType=2)
		{
		SetTitleMatchMode, 2
		Run, "%OutlookPath%" /c ipm.note  /m %Email2%&subject=Need`%20Previous`%20Biopsy`%20Report`%20on`%20%ccn%(%EmailComments%)&body=
		WinWaitActive, (HTML)
		SetTitleMatchMode, 1
		Send, !s
		}
	else if (EmailType=3)
		{
		SetTitleMatchMode, 2
		Run, "%OutlookPath%" /c ipm.note  /m %Email3%&subject=Clinical`%20Note`%20and`%20Photos`%20on`%20%ccn%&body=Please`%20obtain`%20from`%20client`%20clinical`%20note`%20and`%20photos.`%20(%EmailComments%)
		WinWaitActive, (HTML)
		SetTitleMatchMode, 1
		Send, !s
		}
	else if (EmailType=4)
		{
		SetTitleMatchMode, 2
		Run, "%OutlookPath%" /c ipm.note  /m %Email4%&subject=Critical`%20Result`%20Call`%20%ccn%&body=Please`%20fax`%20and`%20call`%20confirm`%20only!
		WinWaitActive, (HTML)
		SetTitleMatchMode, 1
		Send, !s
		}
	else if (EmailType=5)
				{
		SetTitleMatchMode, 2
		Run, "%OutlookPath%" /c ipm.note  /m %Email5%&subject=Please`%20Pull`%20The`%20Bottles`%20And`%20Blocks`%20on`%20%ccn%&body=Please`%20bring`%20them`%20to`%20my`%20office`%20for`%20review`%20with`%20the`%20slides!
		WinWaitActive, (HTML)
		SetTitleMatchMode, 1
		Send, !s
		}

	
	;Msgbox, %EmailComments%
	return
}

^!u::
{
checkForMelanoma:
	SplashTextOn, 100, 100, Melanoma Call Checker, Initializing...

	StringLeft, y, A_UserName, 5
	StringUpper, y, y
	
	s = Select p.name, p.id, p.state, p.suid, p.wsid from pathologists p where p.abbr='%y%'
	CodeDataBaseQuery(s)
	pathWinSurgeId := Result_5
	

		
		searchdate = %a_now%
		searchdate += -5, days
		FormatTime, searchdate, %searchdate%, yyyy-MM-dd
		
/* 		lasttime = 12:00:00 AM
 * 
 * 		finishTime := A_Now - 1500
 * 		FormatTime, endtime, %finishTime%, h:mm:ss tt
 * 		StringSplit, j, endtime, %A_Space%:
 */

	s := "select s.number, s.sotime, s.text01, s.sodate, s.dx from specimen s where s.path =" . pathWinSurgeId . " and s.dx LIKE '%MELANOMA%' and s.sodate >= '" . searchdate . "'" ;and s.sotime >= '" . lasttime . "' and s.sotime <= '" . endtime . "'"

	SplashTextOn, 100, 100, Melanoma Call Checker, Searching for Melanomas...

	WinSurgeQuery(s)
	If A_LastError
		{
			Msgbox, There was an error accessing the WinSURGE Database to check your Melanoma cases.`n`ns=%s%`n  A_LastError = %A_LastError%`n
			Return
		}	
		
	SplashTextOn, 100, 100, Melanoma Call Checker, Filtering for Client Service Calls...

	caselist := ""
	Loop, parse, msg, ¥ 
	{
	if (A_Index = 1)
		Continue
	Else
		{
		x := InStr(A_LoopField,"%%mbn%%")
		x := x + Instr(A_LoopField, "%%mis%%")
		x := x + Instr(A_LoopField, "%%misl%%")
		x := x + Instr(A_LoopField, "%%miss%%")
		x := x + Instr(A_LoopField, "%%mm%%")
		x := x + Instr(A_LoopField, "%%mmm%%")
		x := x + Instr(A_LoopField, "%%mis1%%")
		x := x + Instr(A_LoopField, "%%mmaz%%")
		x := x + Instr(A_LoopField, "%%misaz%%")
		x := x + Instr(A_LoopField, "%%lmaz%%")
		x := x + Instr(A_LoopField, "%%aimm1%%")
		x := x + Instr(A_LoopField, "%%aimm2%%")
		x := x + Instr(A_LoopField, "%%asnmm%%")
		x := x + Instr(A_LoopField, "%%pmis%%")
		x := x + Instr(A_LoopField, "%%nmis%%")
		x := x + Instr(A_LoopField, "%%nmm%%")
		x := x + Instr(A_LoopField, "%%omm%%")
		x := x + Instr(A_LoopField, "%%omis%%")
		x := x + Instr(A_LoopField, "%%pamis%%")
		x := x + Instr(A_LoopField, "%%cmm%%")
		x := x + Instr(A_LoopField, "%%misdn%%")
		x := x + Instr(A_LoopField, "%%mmbap%%")
		x := x + Instr(A_LoopField, "%%mmpn%%")
		x := x + Instr(A_LoopField, "%%idmm%%")

		if x>0
			{
				isFaxed := Instr(intcomment, "fax")
				isFaxed := isFaxed + Instr(intcomment, "notified")
				isFaxed := isFaxed + Instr(intcomment, "called")
				isFaxed := isFaxed + Instr(intcomment, "discussed")
				
				if (isFaxed<=0)
					caselist = %caselist%[ %casenumber% ];  %sodate%`n`nInternal Comments:`n%intcomment%`n`n
			}
		casenumber := sotime
		sotime := intcomment
		intcomment := sodate
		sodate := A_LoopField
	}

	}

	if(caselist)
	{
		SplashTextOff
		SoundBeep
		Msgbox, MELANOMA CALL CHECKER`nIt appears you have one or more cases that was/were melanoma but for which client services has not "fax"ed, "called", or "notified" the client.`nPlease check the below information and take appropriate action!`n`n%caselist%
	}
	else
	{
		SplashTextOn, 100, 100, Melanoma Call Checker, All Cases Notified
		Sleep, 500
	}
		

	caselist := ""
	
	SplashTextOff

	return
}

^!c::
{
	StringLeft, y, A_UserName, 5
	StringUpper, y, y
	ComObjError(True)
	s = Select p.name, p.id, p.state, p.suid, p.wsid from pathologists p where p.abbr='%y%'
	CodeDataBaseQuery(s)

	jarcount:=casecount:=mxcount:=0
    InputBox, searchdate, Search Date "MM-DD", Please enter the date to search.  Leave blank for today!, , 640, 480
	if (searchdate)
		todaydate = %A_YYYY%-%searchdate%
	else
		todaydate = %A_YYYY%-%A_MM%-%A_DD%
	
	s := "select s.number, s.numberofspecimenparts from specimen s where s.path =" . Result_5 . " and s.sodate = '" . todaydate . "'"
	WinSurgeQuery(s)
	Loop, parse, msg, `n 
		{
			If(A_LoopField)
			{
				StringSplit, res, A_LoopField, ¥
				casecount := casecount + 1
				jarcount := jarcount + res3
				IfInString, res2, MX
					mxcount := mxcount + 1
			}
		}
	Msgbox, % todaydate . "`n" . "Cases: " . casecount . ", Jars: " .  jarcount + mxcount * 2 . "`nNote that counts are only approximate and may not account for all consults and two-tray cases."
	return	
}

^!y::
{
	names := []
	nameCounts := {}
	
	InputBox, clientID, ,Enter the Client ID you want to search...,
	
	s := "select s.number, p.name, s.numberofspecimenparts from specimen s, physician p where s.custom04='" . clientID . "' and s.path=p.id and s.sodate>'2017-11-01'"
	;. ", physician p, patient pt where s.patient = pt.id and s.clin=p.id and computed_numberfilled='" . x . "'"
	WinSurgeQuery(s)
	;Msgbox, %msg%
	
	Loop, Parse, msg, `n
		{
				;Msgbox, %A_Index% is %A_LoopField%
				Stringsplit, oput, A_LoopField, ¥
				StringSplit, lastName, oput3,`,
				;Msgbox, %lastName1%
				if (A_Index=1)
					{
						names.Push(lastName1)
						nameCounts[lastName1] := 1
						;Msgbox, % nameCounts[lastName1]
						continue
					}
					
				;Msgbox, % names.Length()
				Loop % names.Length()
					{
					;Msgbox, A_Index=%A_Index%
					if (names[A_Index]=lastName1)
						break
					else if (names.Length()=A_Index)
						{
							names.Push(lastName1)
							nameCounts[lastName1] := 0
						}
					
					nameCounts[lastName1]+=1
					}
		}
			

	total := 0
	theMsg := ""
	perc := "%"
	
	Loop % names.Length()
		{
			theName := names[A_Index]
			theCount := nameCounts[theName]
			total := total + theCount
		}

		;Msgbox, %total%
		Loop % names.Length()
		{
			theName := names[A_Index]
			theCount := nameCounts[theName]
			thePercent := theCount/total * 100
			thePercent := Round(thePercent, 1)
			theMsg = %theMsg%`n%theName% - %thePercent%%perc% - %theCount%
		}
			
		
			
	Msgbox, %theMsg%
			
	Return
}

^!z::
{
		s := "select s.custom04,p.proficiencylog from specimen s, physician p where s.clin=p.id and computed_numberfilled='DD18-027561'"
		WinSurgeQuery(s)
		Msgbox, %msg%
		
		s := "select p.proficiencylog from physician p where p.number='TX6189D'"
		WinSUrgeQuery(s)
		Msgbox, %msg%

	StringSplit, oput, msg, ¥
	Msgbox, %msg%
	Loop, %oput0%
		{
			p := oput%A_Index%
			if p
				Msgbox, %A_Index% is %p%
		}
			
			Msgbox, Done!

x := "Here is some text %%88305%%%%cnmod%%`nHere is some more text. %%88305%%"
		y := RegExReplace(x,"%%\w+%%","")
		MsgBox, %y%

return
}

^!v::
{
	ListVars
	Return
}

^!l::  ;Hotkey for testing the program
{
	ListLines
	Return
}

^!p::
{

	FileDelete, %A_MyDocument%\%CurrentCaseNumber%.txt
	
	FileAppend, %helpText%, %A_MyDocument%\%CurrentCaseNumber%.txt
	Run, Notepad.exe %A_MyDocument%\%CurrentCaseNumber%.txt
	
	Pause
	Send, ^p
	WinWaitActive, Print
	Send, {Enter}
	WinWaitClose, Print
	WinClose, %CurrentCaseNumber%
	WinWaitClose, %CurrentCaseNumber%
	FileDelete, %A_MyDocument%\%CurrentCaseNumber%.txt

	
Return
}

^!1::
{
	If (A_Username<>"mmuenster")
		Return
		
	FileDelete, S:\CodeRocket\bin\EP\_ermpathExtendedPhrases.ahk
	FileMove, S:\CodeRocket\bin\EP\DermpathExtendedPhrases.ahk, S:\CodeRocket\bin\EP\_ermpathExtendedPhrases.ahk
	FileAppend, #NoTrayIcon`n#SingleInstance force`n#Hotstring EndChars  ``t`n#IfWinActive`, WinSURGE - `n, S:\CodeRocket\bin\EP\DermpathExtendedPhrases.ahk

	connectstring := "DRIVER={SQL Server};SERVER=s-irv-sql02;DATABASE=winsurgehotkeys;uid=wshotkeys;pwd=hotkeys10;"
	adodb := ComObjCreate("ADODB.Connection")
	rs := ComObjCreate("ADODB.Recordset")
	rs.CursorType := "0"
	strRequest := "SELECT * FROM extendedphrases"
	adodb.open(connectstring)
	rs := adodb.Execute(strRequest)
	rs.MoveFirst()
	while rs.EOF = 0{
		DXCodeCount := A_Index
		j := 0
		for field in rs.fields
			{
			j := j + 1
			y := Field.Value
			DxCode%j%=%y%
			}
			SplashTextOn, 100, 100, EP convertor, Doing %DXCodeCount%
			FileAppend, ::%DxCode2%::%DxCode3%`n, S:\CodeRocket\bin\EP\DermpathExtendedPhrases.ahk
			rs.MoveNext()
	}
	rs.close()   
	adodb.close()

	SplashTextOff
	Return
}

^!4::
{
	GuiControlGet, ppp, , CaseNumberLabel
	Msgbox, %ppp%
	return
}

Pause::Pause
^!r::
{
	Run, C:\Documents and Settings\All Users\Desktop\Launcher - Caris CodeRocket.exe
	ExitApp
	Return
}

^!q::Reload
^+!r::
{
	FileDelete, %A_MyDocuments%\CarisCodeRocket.ini
	Run, C:\Documents and Settings\All Users\Desktop\Launcher - Caris CodeRocket.exe
	ExitApp
	Return
}

#IfWinActive, WinSURGE - 
Shift & Enter::
{
		If (BeepOnShiftEnter)
			SoundBeep
		
		PercP := "%%P%%"
		Send, {#}{#}
		Sleep, 200
		ControlGetText, x, TX202, WinSURGE - 	
		finaldxcontents := x
		
		;DetermineWhichJar
		
		Loop, parse, finaldxcontents, `n
		{
			;Msgbox, %A_LoopField%
			IfInString, A_LoopField, ##
				break
			else
				{
				IfInString, A_LoopField, %PercP%
					StringLeft, jar, A_LoopField, 1
				}
		}
		
		;Get the Clinical History for that Jar
		y := Asc(%jar%)
		z := y +1 + 65
		nextJar := Chr(z)
		Transform, y, Asc, %jar%
		z := y + 1 

		nextJar := Chr(z)
		StringGetPos, jarPos, ClinicalData, %jar%.
		StringGetPos, nextJarPos, ClinicalData, %nextJar%.
		jarPos := jarPos + 1
		If (nextJarPos>0)
			thisJarLength := nextJarPos - jarPos + 1
		else
			thisJarLength := StrLen(ClinicalData) -  jarPos +1

		thisJarClinicalData := Substr(ClinicalData, jarPos, thisJarLength)


		Send, {Backspace}{Backspace}
		StringGetPos, y, x, ##	
		Stringlen, totallen, x
		firsthalflen:= y
		secondhalflen := totallen-y-2

			TempMicros = 0	
			;TempICD9s = 0	

			Loop,parse, x, `n
			{
				CurrentLine := A_LoopField
				IfInString, CurrentLine, ##
					{
					Loop, parse, CurrentLine, %A_Space%, `n `r	
						{
						CurrentWord := A_LoopField
						IfInstring, CurrentWord, ##
							{
								StringSplit, wordpart, CurrentWord, ##
								StringLen, len, wordpart1
								WinActivate, WinSURGE - 
								Loop, %len%
									{
									Send, {Backspace}	
									firsthalflen := firsthalflen -1
									}
								DxCode := wordpart1
								LastCodeUsed := DxCode
								StringLen, len, DxCode
								StringRight, y, DxCode, 2
								IfInstring, y, /
									{
									TempMicros = 1	
									len := len -1
									}
								IfInstring, y, *
									{
									;TempICD9s = 1	
									len := len -1
									}
								StringLeft, DxCode, DxCode, len
							}
						}
					}
			}	


				Gui, 5:Default
				ErrorLevel := ParseandCheckDXCode()				
				Gui, 1:Default
				
				If ErrorLevel
					return

				ErrorLevel := MildDysplasticWarningCheck()
				If ErrorLevel
						return

				;Loop to add front helper codes to the diagnosis
				Stringlen, i, fronttophelp
				Loop, %i%
				{
					x := i + 1 - A_Index
					StringMid, j, fronttophelp, x, 1
					p := FrontofDiagnosisHelper%j%
					SelDiagnosis = %p% %SelDiagnosis%
				}

				;Loop to add Back helper codes to the diagnosis
				Stringlen, i, backtophelp
				
				
				

				if (i=0)
				{
					ErrorLevel := JerriJohnsonMarginCheck()
					If ErrorLevel
						return
 				}


				if (i>1)
					{
					Msgbox, You can only enter one margin code!	
					GuiControl, Text, DXCode,  ;Blanks the data entry textbox 
					GuiControl, Focus, DXCode
					return
					}
				
				if i=1
					{
					IfInString, UseMargins, none
						{
							SoundBeep
							Msgbox,4,,This client has requested the following margin preferences (%UseMargins%) and you have used one!  Do you wish to continue?
							IfMsgbox, No
								Return	
						}
						
					
						
						p := BackofDiagnosisHelper%backtophelp%
						IfInString, SelDiagnosis, /-/
						{
							StringGetPos, x, SelDiagnosis, /-/
							StringLen, l, SelDiagnosis
							StringLeft, DiagFirstLine, SelDiagnosis, %x%
							y := l - x
							StringRight, RestofDiag, SelDiagnosis, %y%
							SelDiagnosis = %DiagFirstLine%; %p%%RestofDiag%	
						}
						Else
							SelDiagnosis = %SelDiagnosis%; %p%
					}	

				;Loop to add comment helper codes to the comment
				Stringlen, i, comhelp
				Loop, %i%
				{
					StringMid, j, comhelp, %A_Index%, 1
					p := CommentHelper%j%
					SelComment = %SelComment%  %p%  `
				}


				x3 := SelDiagnosis ;Diagnosis
					StringRight, lastletter, x3, 1
						if (lastletter<>"." AND x3<>"")
							x3 = %x3%.
				x4 := SelComment ;Comment
				x5 := SelMicro ;Micro
				x6 := SelCPTCode ;CPT Code
				cr = `n
				StringReplace, x3, x3, /-/, %cr%, All
				StringReplace, x4, x4, /-/, %cr%, All
				perc := "%%"
				dxtext = %x3%
				If ((TempMicros OR UseMicros) AND !x5)
					Msgbox, Client has requested microscopic descriptions and there is not one for this diagnostic code!  Please enter manually.
				If (x4 or (x5 and (TempMicros OR UseMicros)))
					{
					If (TempMicros OR UseMicros)
						dxtext = %dxtext%`n`nComment:%A_Space%%x4%%A_Space%%A_Space%%x5%
					Else
						dxtext = %dxtext%`n`nComment:%A_Space%%x4%
					}
				
				;if (TempICD9s OR UseICD9s)
					;{
					;if SelICD9	
						;dxtext = %dxtext% (%SelICD9%)
					;Else
						;dxtext = %dxtext% (***)
					;}

				dxtext = %dxtext%`n
				If x6
					{
					Loop, parse, x6,`;
						dxtext = %dxtext%%perc%%A_LoopField%%perc%	
					}
				dxtext = %dxtext%%perc%%SelDXCode%%perc%
					
				SetCapsLockState, Off
			
		StringLeft, firsthalf, finaldxcontents, %firsthalflen%
		StringRight, secondhalf, finaldxcontents, %secondhalflen%
		newtext =%firsthalf%%dxtext%%secondhalf%
		
		if UseSendMethod
			{
			Send, %dxtext%	
			Sleep, 50
			Send, {Shift Down}   ;these lines are to correct the but where the shift key is locked down after a send.
			Sleep, 50
			Send, {Shift Up}
			}
		Else
			ControlSetText, TX202, %newtext%, WinSURGE -  

			DataEntered = 1
			

			ifWinActive,  WinSURGE - Final Diagnosis:
			{
			if OpMode<>M
				{
				Gui, Submit, NoHide	
				GuiControl, Enable, CaseScanBox
				GuiControl, Enable, CaseLoaderLbl
				if CaseScanBox=Data in Case
					GuiControl, Text, CaseScanBox, 
				}

			
			finaldiag := WinSURGEFinalDiagnosisContents()
			StringGetPos, i, finaldiag, ***
			if i>0
				ActivateNextTripleAsterisk()
			Else if OpMode<>M
				Gosub, F12
			}

Return
}

#IfWinActive, Special Stains Checklist
WheelDown::
{
	Send, !s
	WinWaitClose, Special Stains Checklist, , 4
	if ErrorLevel
		return
	
	Loop,
		if (A_Cursor="Arrow" OR A_Cursor="IBeam")
			break
	WinWaitActive, WinSURGE , 
	Sleep, 300
	WinActivate, CodeRocket
return
}

#IfWinActive     ;Resets #IfWin directive so that hotkeys can be turned off


;ExternalFunctions These are the functions of the former "External Funtions File"

;DATA ENTRY FUNTIONS
ActivateNextTripleAsterisk() 
{
		WinActivate, WinSURGE -
		Send, ^g
		Return
}
	
AppendOtherLabComment() 
{
	global
	SetControlDelay, 100
	OpenGrossDescriptionModal()
	ControlFocus, TX202, WinSURGE - Gross Description:
	ControlGetPos, x,y,w,h,TX202, WinSURGE - Gross Description:

	mx := x+w-50
	my := y+h-50
	ControlGetText, grossdescriptiontext, TX202, WinSURGE - Gross Description:
	StringGetPos, z, grossdescriptiontext, Microscopic Examination performed
	if z=-1
		{
		grossdescriptiontext = %grossdescriptiontext%  %HomeLabGrossAddendum%`n
		ControlSetText, TX202, %grossdescriptiontext%, WinSURGE - Gross Description:
		ControlFocus, TX202, WinSURGE - Gross Description:
		MouseClick, left, %mx%, %my%
		Send, {End}
		Sleep, 100
		Send, {Space}
		}
	Else
		{
		StringLeft, grossdescriptiontext, grossdescriptiontext, %z%
		grossdescriptiontext = %grossdescriptiontext%  %HomeLabGrossAddendum%`n
		ControlSetText, TX202, %grossdescriptiontext%, WinSURGE - Gross Description:
		ControlFocus, TX202, WinSURGE - Gross Description:
		MouseClick, left, %mx%, %my%
		Send, {End}
		Sleep, 100
		Send, {Space}
		}
	CloseWinSURGEModalWindow("WinSURGE - Gross Description","","&Close")
	Return 0
}

CloseandSaveCase() 
{
		global
		SetTitleMatchMode, 2
		CloseWinSURGEModalWindow("WinSURGE - Final","","&Close")
		IfWinNotActive, WinSURGE [, &7 Save && Open , WinActivate, WinSURGE [, &7 Save && Open			
		WinWaitActive, WinSURGE [, &7 Save && Open
		LabeledButtonPress("WinSURGE [","&7 Save && Open","&7 Save && Open")

		Loop,
		{
			Sleep, 200
			WinActivate, WinSURGE Case Lookup
			IfWinActive, WinSURGE Case Lookup
				break
		}


Return 
}

CloseWinSURGEModalWindow(WinTitle,WinText,CloseButton)
{
	FirstTimeBeep := 0

	Loop,
	{
		IfWinExist, %WinTitle%,%WinText%
			WinClose, %WinTitle%, %WinText%
	
		Sleep, 300
		
		;IfWinExist, %WinTitle%,%WinText%
			;LabeledButtonPress(WinTitle, WinText, CloseButton)
			
		IfWinNotExist, %WinTitle%,%WinText%
			Break
		else
		{			
			IfWinExist, Word Not Found In Dictionary	
			{
				If !FirstTimeBeep
				{
					SoundBeep
					FirstTimeBeep := 1
				}
				Sleep 1000
				Continue
			}
		}

	}
	Return
}
	
LabeledButtonPress(WinTitle, WinText, ButtonLabel)
{
	IfWinExist, %WinTitle%, %WinText%
		{
			ControlGetPos, x,y,w,h,%ButtonLabel%, %WinTitle%, %WinText%	
			x := Round(w/2 + x)
			y := Round(h/2 + y)
			WinActivate, %WinTitle%, %WinText%
			MouseClick, left, %x%, %y%
		}				
	Return
}

OpenGrossDescriptionModal() 
{
		global
		Loop,
		{
		IfWinNotExist, WinSURGE - Gross Description:
			{
			WinWait, WinSURGE [, &7 Save && Open 
			IfWinNotActive, WinSURGE [, &7 Save && Open , WinActivate, WinSURGE [, &7 Save && Open 
			WinWaitActive, WinSURGE [, &7 Save && Open 
			MouseClick, left, %GrossDescriptionButtonX%, %GrossDescriptionButtonY%	
			}
			
			WinWait, WinSURGE - Gross Description: , ,2
			If ErrorLevel
				Continue
			Else
				Break
		}

		IfWinNotActive, WinSURGE - Gross Description:, , WinActivate, WinSURGE - Gross Description:, 
		WinWaitActive, WinSURGE - Gross Description:, 
		WinMove, WinSURGE - Gross Description:,, %WinSURGEModalWindowX%,%WinSURGEModalWindowY%,%WinSURGEModalWindowW%,%WinSURGEModalWindowH%

		Return 0
	}

OpenFinalDiagnosisModal() 
{
		global
		Loop,
		{
		IfWinNotExist, WinSURGE - Final Diagnosis:
			{
			CloseWinSURGEModalWindow("WinSURGE Case Lookup","","Cancel")
			WinWait, WinSURGE [, &7 Save && Open 
			IfWinNotActive, WinSURGE [, &7 Save && Open , WinActivate, WinSURGE [, &7 Save && Open 
			WinWaitActive, WinSURGE [, &7 Save && Open 
			MouseClick, left, %FinalDiagnosisButtonX%, %FinalDiagnosisButtonY%	
			}
			
			WinWait, WinSURGE - Final , ,2
			If ErrorLevel
				Continue
			Else
				{
				Break
				}
		}

		IfWinNotActive, WinSURGE - Final Diagnosis:, , WinActivate, WinSURGE - Final Diagnosis:, 
		WinWaitActive, WinSURGE - Final Diagnosis:, 
		WinMove, WinSURGE - Final Diagnosis:,, %WinSURGEModalWindowX%,%WinSURGEModalWindowY%,%WinSURGEModalWindowW%,%WinSURGEModalWindowH%
		Return 0
	}

OpenCase(obj) 
{
/* 	global
 * 	Loop, 
 * 		{
 * 			IfWinNotExist, WinSURGE Case Lookup
 * 				LabeledButtonPress("WinSURGE","&7 Save && Open","&2 Open Case")
 * 			Else
 * 				Break
 * 				
 * 			Sleep, 100
 * 		}
 * 
 * 	WinActivate, WinSURGE Case Lookup
 * 	If !CaseLookupCaseNumberTextBox   ;Setup function to get the name of the Case Lookup Textbox the first time through
 * 		{
 * 		Loop,
 * 			{
 * 				WinGetTitle, active_title, A
 * 				If (active_title="WinSURGE Case Lookup")
 * 					{
 * 						Sleep, 300
 * 						ControlGetFocus, x,	%active_title%
 * 						CaseLookupCaseNumberTextBox := x
 * 						Break
 * 					}
 * 					
 * 				Sleep, 100
 * 			}
 * 		}
 * 		
 * 	ControlSetText, %CaseLookupCaseNumberTextBox%, %obj%, WinSURGE Case Lookup	
 * 	ControlFocus, %CaseLookupCaseNumberTextBox%, WinSURGE Case Lookup
 * 	Send, {Enter}{Enter}     ;Need two enters to properly get to QueueIntoBatchbox
 * 
 * if (QueueIntoBatchBox AND QueueIntoBatchBox<>"ERROR")
 * 	{
 * 	Loop,	
 * 		{
 * 			CloseWinSURGEModalWindow("WinSURGE Case Lookup","","Cancel") ;To fix bug in WinSurge where the Window just pops up!
 * 			ControlFocus, %QueueIntoBatchBox%, WinSURGE
 * 			ControlGetFocus, x, WinSURGE,
 * 			if x=%QueueIntoBatchBox%
 * 				Break
 * 		}
 * }
 * 
 */
 
 	global
	WinActivate, WinSURGE
	Loop, 
		{
			IfWinNotExist, WinSURGE Case Lookup
				Send, !2 ;LabeledButtonPress("WinSURGE","&7 Save && Open","&2 Open Case")
			Else
				Break
				
			Sleep, 50
		}

	WinActivate, WinSURGE Case Lookup
	If !CaseLookupCaseNumberTextBox   ;Setup function to get the name of the Case Lookup Textbox the first time through
		{
		Loop,
			{
				WinGetTitle, active_title, A
				If (active_title="WinSURGE Case Lookup")
					{
						Sleep, 300
						ControlGetFocus, x,	%active_title%
						CaseLookupCaseNumberTextBox := x
						Break
					}
					
				Sleep, 500
			}
		}
		
	ControlSetText, %CaseLookupCaseNumberTextBox%, %obj%, WinSURGE Case Lookup	
	ControlFocus, %CaseLookupCaseNumberTextBox%, WinSURGE Case Lookup
	Send, {Enter}     ;Need two enters to properly get to QueueIntoBatchbox

	Loop,
	{
		MouseMove, 500, 500
		Sleep, 400
		If (A_Cursor="Wait")
			continue
		else
			break
	}

Return
}

QueueandAssign() 
{
			global
			Loop, 
			{
				SetControlDelay, 200
				ControlSetText, %QueueIntoBatchBox%, Final reports, WinSURGE [
				ControlClick, %PathologistTextBox%, WinSURGE [
				ControlSetText, %PathologistTextBox%, , WinSURGE [
				len := StrLen(WinSurgeFullName)
				len := len -1
				StringLeft, PartialName, WinSurgeFullName, %len%
				ControlSetText, %PathologistTextBox%, %PartialName%, WinSURGE [
				ControlSend, %PathologistTextBox%, {TAB}, WinSURGE [
				ControlClick, %QueueIntoBatchBox%, WinSURGE [

				Loop,20 
					{
					ControlGetText, t1, %PathologistTextBox%, WinSURGE [
					Sleep, 100
					if t1=%WinSurgeFullName%
						Break
					}
				
				if t1=%WinSurgeFullName%
						Break
			}
return

}

SkipCaseSignout()
{
	global
	WinWaitClose, WinSURGE E-signout, abcdefgABCDEFG 1234567890, , [
	WinActivate, WinSURGE E-signout [, &Approve
	WinWaitActive, WinSURGE E-signout [, &Approve

	Loop,
	{
	LabeledButtonPress("WinSURGE E-signout [","&Approve","&Skip")
	Sleep, 1000

	If WinExist("WinSURGE", "abcdefgABCDEFG 1234567890","E-signout","")    
		Break
	if WinExist("WinSURGE E-signout","abcdefgABCDEFG 1234567890","[","")
		Break
	}
Return
}

WinSURGEFinalDiagnosisContents() 
{
	ifWinNotExist, WinSURGE - Final Diagnosis:
		OpenFinalDiagnosisModal()	
	ControlGetText, t, TX202, WinSURGE - Final Diagnosis:	
	Return t
}

WinSURGEOpenCasePhoto1() 
{
	global
	ControlGetText, t1, %Photo1TextBox%, WinSURGE [
	Return t1
}

WinSURGEOpenCasePhoto2() 
{
	global
	ControlGetText, t1, %Photo2TextBox%, WinSURGE [
	Return t1
}

;CASE SIGNOUT FUNCTIONS
EnterSignoutPasswordandApprove() 
{
	global
	WinWaitClose, WinSURGE E-signout, abcdefgABCDEFG 1234567890, , [
	WinActivate, WinSURGE E-signout [, &Approve
	WinWaitActive, WinSURGE E-signout [, &Approve
	

	Loop,    ;Makes sure the signout password textbox has the focus.
		{
			Sleep, 200
			ControlFocus, %ApprovalPasswordControl%, WinSURGE E-signout [
			ControlGetFocus, ctrlfoc, WinSURGE E-signout [
			if ctrlfoc=%ApprovalPasswordControl%
				Break
			Else ifWinExist, WinSURGE E-signout, abcdefgABCDEFG 1234567890, [ ;In case the window for signout just pops up!
				Return	

		}

	
	Loop,    ;Enters the sopw
	{
		ControlSetText, %ApprovalPasswordControl%, , WinSURGE E-signout [	
		ControlSetText, %ApprovalPasswordControl%, %WinSurgeSignoutPassword%, WinSURGE E-signout [	
		ControlGetText, t1, %ApprovalPasswordControl%, WinSURGE E-signout [
		if t1=%WinSurgeSignoutPassword%
			Break
	}
	SetTitleMatchMode, 2
	Loop,
	{
	LabeledButtonPress("WinSURGE E-signout [","&Approve","&Approve")
	Sleep, 1000

	If WinExist("WinSURGE", "OK","E-signout")
		{
			Sleep, 500
			ControlClick, OK, WinSURGE, Ok, left
			SplashTextOn, 200,100, Setup..., Close any error windows.  Then type your signout password into the box and press F3 on your keyboard
			ControlClick, OK, WinSURGE, Ok, left
			KeyWait, F3, d
			KeyWait, F3, u
			SplashTextOff
			ControlGetText, t1, %ApprovalPasswordControl%, WinSURGE E-signout [
			WinSurgeSignoutPassword=%t1%
			IniWrite, %WinSurgeSignoutPassword%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSurgeSignoutPassword
			Continue
		}
	If WinExist("WinSURGE", "abcdefgABCDEFG 1234567890","E-signout","")    
		Break
	if WinExist("WinSURGE E-signout","abcdefgABCDEFG 1234567890","[","")
		Break
	}
SetTitleMatchMode, 1
Return
}

Open4SignOut(obj) 
{
	global
	JumptoCaseNotSet := 0
	IfWinNotExist, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
				{
				WinMenuSelectItem, WinSURGE , &7 Save && Open , Tools, E-Signout	
				Loop, 10
					{
							WinWait, WinSURGE E-signout,abcdefgABCDEFG 1234567890,0.3,[
							If ErrorLevel
								Continue
							Else
								Break
					}
						
				ifWinExist, WinSURGE E-signout [, R&eturn	
					{
					JumptoCaseNotSet := 1	
					Return
					}
				}
	
	WinWait, WinSURGE E-signout,abcdefgABCDEFG 1234567890,,[
	IfWinNotActive, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[, , WinActivate, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[,
	WinWaitActive, WinSURGE E-signout,abcdefgABCDEFG 1234567890,,[

	If (!EsignoutTextBox OR EsignoutTextBox="ERROR")  ;Setup function to get the name of the EsignoutTextBox the first time through
		{
		ControlFocus, ThunderRT6TextBox, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
		Sleep, 300
		ControlGetFocus, x,	WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
		StringGetPos, t, x, TextBox
		if t>0
			EsignoutTextBox := x
		Else
			{
				ToolTip, Click in the "Enter case to jump to:" box and press F3!
				KeyWait, F3, d
				KeyWait, F3, u
				ControlGetFocus, x,	WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
				EsignoutTextBox := x
				ToolTip   ;Removes the tooltip
			}

		IniWrite, %EsignoutTextBox%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, EsignoutTextBox
		}

	Loop,   ;Enters the case number to the E-signout window
	{
		ControlSetText, %EsignoutTextBox%, %obj%, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[	
		Sleep, 30
		ControlGetText, t1, %EsignoutTextBox%, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
		if t1=%obj%
			Break
	}

	Loop,   ;Clicks the OK button to the E-signout window
	{
		IfWinExist, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
			{
			LabeledButtonPress("WinSURGE E-signout", "abcdefgABCDEFG 1234567890", "OK")
			WinWaitClose, WinSURGE E-signout,abcdefgABCDEFG 1234567890, 2, [
			If ErrorLevel
				Continue
			Else
				Break
			
			}
	}
Return
}

ReleaseApprovedCases()
{
	global
	if WinExist("WinSURGE E-signout","abcdefgABCDEFG 1234567890","[","")
		{
			Loop,
			{
				LabeledButtonPress("WinSURGE E-signout","abcdefgABCDEFG 1234567890","Cancel") 
				WinWaitClose, WinSURGE E-signout,abcdefgABCDEFG 1234567890,1,[
				If ErrorLevel
					Continue
				Else
					Break
			}
			
			WinWait, WinSURGE, Retry, 1
			If ErrorLevel
				LabeledButtonPress("WinSURGE E-signout","abcdefgABCDEFG 1234567890","&Release")
			Else
				LabeledButtonPress("WinSURGE","Retry","Yes")
			
			
			WinWait, E-Signout Approval & Release
			WinActivate, E-Signout Approval & Release
			LabeledButtonPress("E-Signout Approval & Release","&Release","&Release")	
			Return
		}
		
	If WinExist("WinSURGE", "abcdefgABCDEFG 1234567890","E-signout","")    
		{
			LabeledButtonPress("WinSURGE", "&Yes","&Yes")
			WinWait, E-Signout Approval & Release
			WinActivate, E-Signout Approval & Release
			LabeledButtonPress("E-Signout Approval & Release","&Release","&Release")	
			Return
		}
}

;DATABASE QUERY FUNCTIONS
{
WinSurgeQuery(s)
{
	global
	Loop, 15          ;Blank the results array
		Result_%A_Index% =

	;connectstring := "DRIVER={InterSystems ODBC};SERVER=s-irv-wsg01;DATABASE=wins1csql;uid=_system;pwd=sys;"
	connectstring := "DRIVER={InterSystems ODBC};SERVER=dfw-mpwsg01;DATABASE=wins1csql;uid=sqluser;pwd=sqluserpw;"
	adodb := ComObjCreate("ADODB.Connection")
	rs := ComObjCreate("ADODB.Recordset")
	rs.CursorType := "0"
	adodb.open(connectstring)
RetryWinSURGEDatabase:
	rs := adodb.Execute(s)
	If A_LastError
		{
			Msgbox, 4101, Error Message, There was an error accessing the WinSURGE Database.`n`ns=%s%`n  A_LastError = %A_LastError%
			IfMsgbox, Retry
				Goto, RetryWinSURGEDatabase
			else ifMsgBox, Cancel
				ExitApp
		}	
	msg := ""
	txt := rs.state
	If !txt
		return
	while rs.EOF = 0{
		for field in rs.fields
			msg := msg . "¥" . Field.Value
		msg = %msg%`n
		rs.MoveNext()
	}
	
	Loop, parse, msg, ¥ 
		{
		if (A_Index = 1)
			Continue
		Else
			{
			 i := A_Index -1 
			 Result_%i% := A_LoopField		
			}
		}
	
	rs.close()   
	adodb.close()
	return 	
}

CodeDatabaseQuery(s)  ;((((((((((((((((
{
global
RetryCodeDatabase:
	Loop, 15          ;Blank the results array
		Result_%A_Index% =
		
	connectstring := "DRIVER={SQL Server};SERVER=s-irv-sql02;DATABASE=winsurgehotkeys;uid=wshotkeys;pwd=hotkeys10;"
	adodb := ComObjCreate("ADODB.Connection")
	adodb.open(connectstring)

	rs := adodb.Execute(s)
	If A_LastError
		{
			Msgbox, 4101, Error Message, There was an error accessing the Code Database.`n  `ns=%s%`n  A_LastError = %A_LastError%
			IfMsgbox, Retry
				Goto, RetryCodeDatabase	
			else ifMsgBox, Cancel
				ExitApp
		}	
		
	msg := ""
	txt := rs.state
	If !txt
		return
		
	while rs.EOF = 0{
		for field in rs.fields
			msg := msg . "¥" . Field.Value
		rs.MoveNext()
	}
	
	Loop, parse, msg, ¥ 
		{
		if (A_Index = 1)
			Continue
		Else
			{
			 i := A_Index -1 
			 Result_%i% := A_LoopField		
			}
		}

	rs.close()   
	adodb.close()
	return msg	
}
}
   ;End of the External Functions

;Internal Functions - These are the functions of the former "Internal Functions File"

performCPTCodeCheck(cptCodesFrom61,finalDiagnosisText)
{

	;Everything in these arrays must be a single "word" without a space, comma or period.

#Include CPTCheckSupport.ahk

	v305 := "%%88305%%"
	v304 := "%%88304%%"
	v312 := "%%88312%%"
	v313 := "%%88313%%"
	v342 := "%%88342%%"
	v365 := "%%88365%%"
	
	dperc := "%%"

	finalDiagnosisCptCodes := []
		
	StringReplace, finalDiagnosisText, finalDiagnosisText, %v305%, %A_Space%88305%A_Space%, All
	StringReplace, finalDiagnosisText, finalDiagnosisText, %v304%, %A_Space%88304%A_Space%, All
	StringReplace, finalDiagnosisText, finalDiagnosisText, %v312%, %A_Space%88312%A_Space%, All
	StringReplace, finalDiagnosisText, finalDiagnosisText, %v342%, %A_Space%88342%A_Space%, All
	StringReplace, finalDiagnosisText, finalDiagnosisText, %v313%, %A_Space%88313%A_Space%, All

	StringSplit, arrayFinalDiagnosis, finalDiagnosisText, %A_Space%.`,:

	Loop, %arrayFinalDiagnosis0%
	{
		p := arrayFinalDiagnosis%A_Index%
		if p=88305
			finalDiagnosisCptCodes.push(88305)
		else if p=88304
			finalDiagnosisCptCodes.push(88304)
		else if p=88312
			finalDiagnosisCptCodes.push(88312)
		else if p=88342
			finalDiagnosisCptCodes.push(88342)
		else if p=88313
			finalDiagnosisCptCodes.push(88313)
		else if p=88365
			finalDiagnosisCptCodes.push(88365)
	}

	allCptCodes := cptCodesFrom61

	For i,j in finalDiagnosisCptCodes
		allCptCodes.push(j)

	documentedFinalDiagnosisBillCodes := []

	Loop, %arrayFinalDiagnosis0%
	{
		p := arrayFinalDiagnosis%A_Index%

		For index, value in array312
			{
				if (p=value)
					{
						documentedFinalDiagnosisBillCodes.Insert("88312")
						fdText=%fdText%'%p%' - 88312`n
					}
			}
		For index, value in array313
			{
				if (p=value)
					{
						documentedFinalDiagnosisBillCodes.Insert("88313")
						fdText=%fdText%'%p%' - 88313`n
					}
			}
		For index, value in array342
			{
				if (p=value)
					{
						documentedFinalDiagnosisBillCodes.Insert("88342")
						fdText=%fdText%'%p%' - 88342`n
					}
			}
		For index, value in array365
			{
				if (p=value)
					{
						documentedFinalDiagnosisBillCodes.Insert("88365")
						fdText=%fdText%'%p%' - 88365`n
					}
			}

		if (p=88305)
			{
				documentedFinalDiagnosisBillCodes.Insert("88305")
				fdText=%fdText%'%dperc%88305%dperc%' - 88305`n
			}

		if (p=88304)
			{
				documentedFinalDiagnosisBillCodes.Insert("88304")
				fdText=%fdText%'%dperc%88305%dperc%' - 88304`n
			}

	}
		
		ErrorLevel=0
		
		For index, value in allCptCodes
		{
			matchPos := -1   ; returns -1 if there is no match, otherwise returns array position of the match

			For i,v in documentedFinalDiagnosisBillCodes
				{
					if (value=v)
					{
						matchPos := i
						break
					}
				}

			if (matchPos>0)
				documentedFinalDiagnosisBillCodes.removeAt(matchPos)
			else
			{
				If(value!="88")  ;Some cases have an "88" code.  I don't know why...So we ignore it.
				{
				MsgBox, 4,, There is no documentation for a %value% cpt code.`n%fdText%`nDo you wish to ignore and continue? (press Yes or No)
				IfMsgBox No
					ErrorLevel=1
				}
			}
		}

		return ErrorLevel
}

get_filled_case_number(c)
{
    StringUpper c,c
	ret := ""
    StringSplit,arr,c,-
	if (arr0 = 2) {
        stringlen,len,arr2
		repeatnum := 6-len
		zeros:=""
		cnt := 0
		While (cnt < repeatnum){
			cnt := cnt +1
			zeros:=zeros . "0"
		}
		ret := arr1 . "-" . zeros . arr2
	}		
	return ret
} 

CalculateWindowSizes()
{
	global
	If A_ScreenWidth>1270
	{
	CarisRocketWindowX := 415	
	CarisRocketWindowY := 2
	CarisRocketWindowW = 400
	CarisRocketWindowH = 260
	WinSURGEModalWindowX := 415
	WinSURGEModalWindowY := 313
	WinSURGEModalWindowW := 500
	WinSURGEModalWindowH := 500
	}
	Else if A_ScreenWidth > 1024
	{
	CarisRocketWindowX := 415	
	CarisRocketWindowY := 2
	CarisRocketWindowW = 400
	CarisRocketWindowH = 260
	WinSURGEModalWindowX := 415
	WinSURGEModalWindowY := 313
	WinSURGEModalWindowW := 500
	WinSURGEModalWindowH := 500
	}

	Return
	}
	
ParseQAQCData()
{
	global
	msg = 
	
OrderedCPTCodes:   ;OrderedCPTCodeX  OrderedCPTCount
{
OrderedCPTCount = 0
Loop, Parse, OrderedCPTCodes, CSV	
			{
			OrderedCPTCode%A_Index% = %A_LoopField%	
			OrderedCPTCount := A_Index
			}
		msg = %msg%OrderedCPTCount = %OrderedCPTCount%, %OrderedCPTCode1%, %OrderedCPTCode2%`n
}

GrossDescription:   ;For up to ten vials, grosstext_X, GrossVialCount
{

	StringCaseSense, On
	PositionNotFound := 0
	GrossVialCount := 0
	pos_0 := 0	

	Loop, 26
		{
		grosstext_%A_Index% = 	
		pos_%A_Index% =
		ascii_code := Chr(64+A_Index)
			
		StringGetPos, pos_%A_Index%, grossdescriptiontext, %ascii_code%.%A_Space%		
		If (Errorlevel AND A_Index=1)
			{
			StringCaseSense, Off
			StringGetPos, pos_1, grossdescriptiontext, received
			if Errorlevel
				msg = %msg%The gross description could contain an error because no information for "A. " could be found. `n	
			Else
				{
				GrossVialCount := 1	
				grosstext_1 = %grossdescriptiontext%
				Break
				}
			}

	GrossVialCount := A_Index - 1
	if pos_%A_Index% > 0
		l := pos_%A_Index% - pos_%GrossVialCount% - 1
	Else
		l := StrLen(grossdescriptiontext)
		
	StringMid, grosstext_%GrossVialCount%, grossdescriptiontext, pos_%GrossVialCount%, %l%
	If ErrorLevel
		Break
		}
	StringCaseSense, Off
}

FinalDiagnosis:    ;finaltext_x, biopsytype_X, FinalVialCount
{
		msg=%msg%GrossVialCount=%GrossVialCount%|| %grosstext_1%|| %grosstext_2%`n
		Needle := "%%P%%"
		i := 0
		h := 0
		k := 1

		Loop, 26
		{
			biopsytype_%A_Index% =
			finaltext_%A_Index% = 
		}
		
		Loop, 26
		{
			StringGetPos, i, finaldiagnosistext, %Needle%, , %k%
			if i>0
				{
				j := i - k
				h := h + 1
				StringMid, finaltext_%A_Index%, finaldiagnosistext, %k%+1, %j%
				k := i + 6
				StringGetPOs, y, finaltext_%A_Index%, shave
					if y>0
						biopsytype_%A_Index% = shave
				StringGetPOs, y, finaltext_%A_Index%, punch
					if y>0
						biopsytype_%A_Index% = punch
				StringGetPOs, y, finaltext_%A_Index%, excision
					if y>0
						biopsytype_%A_Index% = excision
				StringGetPOs, y, finaltext_%A_Index%, curettage
					if y>0
						biopsytype_%A_Index% = curettage
				StringReplace, finaltext_%A_Index%, finaltext_%A_Index%,*** ,,All
				}
			Else
				Break
		}
	
	FinalVialCount := h
	msg = %msg%FinalVialCount =%FinalVialCount%|| %finaltext_1%|| %finaltext_2%`n
/*
Loop, %FinalVialCount%
		{
			StringcaseSense, Off
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%R%A_space%, %A_Space%Right%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%R.%A_space%, %A_Space%Right%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%Rt%A_space%, %A_Space%Right%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%Rt.%A_space%, %A_Space%Right%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%L%A_space%, %A_Space%Left%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%L.%A_space%, %A_Space%Left%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%Lt%A_space%, %A_Space%Left%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%Lt.%A_space%, %A_Space%Left%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%abd%A_space%, %A_Space%abdomen%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%lat%A_space%, %A_Space%lateral%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%med%A_space%, %A_Space%medial%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%prox%A_space%, %A_Space%proximal%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%dist%A_space%, %A_Space%distal%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%bx%A_space%, %A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%bx., %A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%bx:, %A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%bx.:, %A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%post%A_space%, %A_Space%posterior%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%ant%A_space%, %A_Space%anterior%A_Space%
		}
*/

		}

ClinicalData:
{
	StringCaseSense, On
	PositionNotFound := 0
	ClinDataVialCount := 0
	pos_0 := 0	

	Loop, 26
		{
		clindata_%A_Index% = 	
		pos_%A_Index% =
		ascii_code := Chr(64+A_Index)
			
		StringGetPos, pos_%A_Index%, ClinicalData, %ascii_code%.%A_Space%		
		If (Errorlevel AND A_Index=1)
			{
				StringCaseSense, Off
				ClinDataVialCount := 0	
				clindata_1 = %ClinicalData%
				Break				
			}

	ClinDataVialCount := A_Index - 1
	if (pos_%A_Index% > 0 AND pos_%A_Index% > pos_%ClinDataVialCount%)
		l := pos_%A_Index% - pos_%ClinDataVialCount% - 1
	Else
		l := StrLen(ClinicalData)
		
	StringMid, clindata_%ClinDataVialCount%, ClinicalData, pos_%ClinDataVialCount%, %l%
	If ErrorLevel
		Break
	}
	StringCaseSense, Off
	If ClinDataVialCount =0
		ClinDataVialCount = 1
	msg = %msg%%ClinDataVialCount%||%clindata_1%||%clindata_2%`n	
}
		
Return
}

ParseandCheckDXCode()
{
	global
				SelectedCodeIndex := 0
				LV_Modify(0, "-Select") 
				
				Stringlen, ltot, DxCode
				StringGetPos, j, DxCode,.
				StringGetPos, k, DxCode,;
				StringGetPos, i, DxCode,:
					
				if k=-1
					comhelp = 
				Else
					{
						x := ltot - k - 1
						StringRight, comhelp, DxCode, x
					}
				
				Stringlen, comlength, comhelp
				If (comlength>0)
					comlength := comlength + 1
					
				if j=-1
					backtophelp = 
				Else
					{
					if k=-1
						x :=  ltot - j -1
					Else
						x := k - j - 1
						
					xstart := j + 2
					StringMid, backtophelp, DxCode, xstart, x
					}

				Stringlen, backlength, backtophelp
				If (backlength>0)
					backlength := backlength + 1

				if i=-1
					{
					fronttophelp =	
					baselength := ltot - comlength - backlength
					StringLeft, basediag, DxCode, baselength
					}					
				Else
					{
					StringLeft, fronttophelp, DxCode, i	
					Stringlen, frontlength, fronttophelp
					frontlength := frontlength + 1
					xstart := i + 2
					x := ltot - comlength - backlength - frontlength
					StringMid, basediag, DxCode, xstart, x
					}
									
				Loop % LV_GetCount()
					{
					LV_GetText(RetrievedText, A_Index, 2)
					if basediag=%RetrievedText%
						{
						LV_Modify(A_Index, "Select")  ; Select each row whose first field contains the filter-text.
						SelectedCodeIndex = %A_Index%
						Break
						}
					}

				If (SelectedCodeIndex = 0 OR ltot=0)
					{
					Msgbox, That is not a valid diagnosis code!
					return, 1
					}
				
				
				LV_GetText(SelDXCode,SelectedCodeIndex,2)  
				LV_GetText(SelDiagnosis,SelectedCodeIndex,5)  
				LV_GetText(SelComment,SelectedCodeIndex,6)
				LV_GetText(SelMicro,SelectedCodeIndex,7)
				LV_GetText(SelCPTCode,SelectedCodeIndex,8)
				LV_GetText(SelICD9,SelectedCodeIndex,9)
				LV_GetText(SelICD10,SelectedCodeIndex,10)
				LV_GetText(SelSnomed,SelectedCodeIndex,11)
				LV_GetText(SelPre,SelectedCodeIndex,12)
				LV_GetText(SelMal,SelectedCodeIndex,13)
				LV_GetText(SelDys,SelectedCodeIndex,14)
				LV_GetText(SelMel,SelectedCodeIndex,15)				
				LV_GetText(SelInf,SelectedCodeIndex,16)
				LV_GetText(SelMargInc,SelectedCodeIndex,17)
				LV_GetText(SelLog,SelectedCodeIndex,18)
				
/*
Stringlen, x, backtophelp
				If (x > 1)    ;MULTIPLE MARGIN CODES WERE ENTERED
					{
					Msgbox, You may only enter one margin code!
					return, 1
					}
				else if (x=1)  ;MARGIN CODE WAS ENTERED
					{
						if SelInf
						{
							Msgbox, 4, ,The diagnosis code you selected is "inflammatory" and you gave a margin.  Are you sure you wish to continue?
							IfMsgbox No
								Return 1
						}
						else if SelMal
						{
							if (MarginNoPreference OR MarginMalignant OR MarginAll)
								Return 0
							Else
							{
								Msgbox, 4, , The client has not requested to have margins on 
							}
						}
						else if SelDys
						{
						}
						else if SelMel
						{
						}
						else if SelPre
						{
						}
							
					}
				else if (x=0)   ;MARGIN CODE WAS NOT ENTERED
					{
						;Margin Checking to ensure lack of margin information is ok goes here.
					}
*/
return 0
}

ZeroMarginFlags()
{
	global
	MarginAll = 0
	MarginOnRequestOnly = 0
	MarginNoPreference = 0
	MarginExcision = 0
	MarginMelanocytic = 0
	MarginDysplastic = 0
	MarginMalignant = 0
	MarginPremalignant = 0
	MarginShave = 0
	Return
}

ReadDXCodes()
{
	global
	IfExist, %A_MyDocuments%\dxcodes.csv
	{
		LV_Delete()
		Loop, read, %A_MyDocuments%\dxcodes.csv
		{
			Loop, parse, A_LoopReadLine, CSV
				DxCode%A_Index%=%A_LoopField%
		
		LV_Add("",DXCode1,DXCode2,DXCode3,DXCode4,DXCode5,DXCode6,DXCode7,DXCode8,DXCode9,DXCode10,DXCode11,DXCode12,DXCode13,DXCode14,DXCode15,DXCode16,DXCode17,DXCode18)
		}
		LV_ModifyCol()  ; Auto-size each column to fit its contents. 
		LV_Modify(1, "Sort")
}
	else
	{
	LV_Delete()
	connectstring := "DRIVER={SQL Server};SERVER=s-irv-sql02;DATABASE=winsurgehotkeys;uid=wshotkeys;pwd=hotkeys10;"
	adodb := ComObjCreate("ADODB.Connection")
	rs := ComObjCreate("ADODB.Recordset")
	rs.CursorType := "0"
	strRequest := "SELECT * FROM dxcodes"
	adodb.open(connectstring)
	rs := adodb.Execute(strRequest)
	rs.MoveFirst()
	while rs.EOF = 0{
		DXCodeCount := A_Index
		j := 0
		for field in rs.fields
			{
			j := j + 1
			y := Field.Value
			DxCode%j%=%y%
			}
			LV_Add("",DXCode1,DXCode2,DXCode3,DXCode4,DXCode5,DXCode6,DXCode7,DXCode8,DXCode9,DXCode10,DXCode11,DXCode12,DXCode13,DXCode14,DXCode15,DXCode16,DXCode17,DXCode18)
			rs.MoveNext()
}

	rs.close()   
	adodb.close()

	LV_ModifyCol()  ; Auto-size each column to fit its contents. 
	LV_Modify(1, "Sort")
}

return
}

ReadHelpers()
{

	global

;Front Helper Load
IfExist, %A_MyDocuments%\fronthelpers.csv
{
	Loop, read, %A_MyDocuments%\fronthelpers.csv
	{
		Loop, parse, A_LoopReadLine, CSV
			DxCode%A_Index%=%A_LoopField%
		
		FrontofDiagnosisHelper%DxCode2% = %DxCode3%
	}
}
else
{
	connectstring := "DRIVER={SQL Server};SERVER=s-irv-sql02;DATABASE=winsurgehotkeys;uid=wshotkeys;pwd=hotkeys10;"
	adodb := ComObjCreate("ADODB.Connection")
	rs := ComObjCreate("ADODB.Recordset")
	rs.CursorType := "0"
	strRequest := "SELECT * FROM fronthelpers"
	adodb.open(connectstring)
	rs := adodb.Execute(strRequest)
	rs.MoveFirst()
	while rs.EOF = 0{
		DXCodeCount := A_Index
		j := 0
		for field in rs.fields
			{
			j := j + 1
			y := Field.Value
			DxCode%j%=%y%
			}
			FrontofDiagnosisHelper%DxCode2% = %DxCode3%
			rs.MoveNext()
	}
	rs.close()   
	adodb.close()
}

;Margin Load	
IfExist, %A_MyDocuments%\margins.csv
{
	Loop, read, %A_MyDocuments%\margins.csv
	{
		Loop, parse, A_LoopReadLine, CSV
			DxCode%A_Index%=%A_LoopField%
		
		BackofDiagnosisHelper%DxCode2% = %DxCode3%
	}
}
else
{
	connectstring := "DRIVER={SQL Server};SERVER=s-irv-sql02;DATABASE=winsurgehotkeys;uid=wshotkeys;pwd=hotkeys10;"
	adodb := ComObjCreate("ADODB.Connection")
	rs := ComObjCreate("ADODB.Recordset")
	rs.CursorType := "0"
	strRequest := "SELECT * FROM margins"
	adodb.open(connectstring)
	rs := adodb.Execute(strRequest)
	rs.MoveFirst()
	while rs.EOF = 0{
		DXCodeCount := A_Index
		j := 0
		for field in rs.fields
			{
			j := j + 1
			y := Field.Value
			DxCode%j%=%y%
			}
			BackofDiagnosisHelper%DxCode2% = %DxCode3%
			rs.MoveNext()
	}
	rs.close()   
	adodb.close()
}

;Comment Helper Load	
IfExist, %A_MyDocuments%\commenthelpers.csv
{
	Loop, read, %A_MyDocuments%\commenthelpers.csv
	{
		Loop, parse, A_LoopReadLine, CSV
			DxCode%A_Index%=%A_LoopField%
		
		CommentHelper%DxCode2% = %DxCode3%
	}
}
else
{
	connectstring := "DRIVER={SQL Server};SERVER=s-irv-sql02;DATABASE=winsurgehotkeys;uid=wshotkeys;pwd=hotkeys10;"
	adodb := ComObjCreate("ADODB.Connection")
	rs := ComObjCreate("ADODB.Recordset")
	rs.CursorType := "0"
	strRequest := "SELECT * FROM commenthelpers"
	adodb.open(connectstring)
	rs := adodb.Execute(strRequest)
	rs.MoveFirst()
	while rs.EOF = 0{
		DXCodeCount := A_Index
		j := 0
		for field in rs.fields
			{
			j := j + 1
			y := Field.Value
			DxCode%j%=%y%
			}
			CommentHelper%DxCode2% = %DxCode3%
			rs.MoveNext()
	}
	rs.close()   
	adodb.close()
}
return
}

FirstTimeSetup()
{
	global
	IfWinNotExist, WinSURGE
		{
			Msgbox, WinSURGE must be running the first time you fire the Caris CodeRocket.  Please login to WinSURGE and restart the Caris CodeRocket.
			ExitApp
		}
		
	perc := "%"
	StringLeft, y, A_UserName, 5
	StringUpper, y, y
	ComObjError(True)
	s = Select p.name, p.id, p.state, p.suid, p.wsid from pathologists p where p.abbr='%y%'
	CodeDataBaseQuery(s)

	if !Result_1
	{
		s = Select p.name, p.id, p.state, p.suid, p.wsid from pathologists p where p.ttwentyfive03 LIKE '%y%%perc%'
	CodeDataBaseQuery(s)
		if !Result_1
			{
				OpMode = T
				IniWrite, %A_Username%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WindowsLoginId
				IniWrite, %OpMode%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, OpMode
				Return
			}
			
	}
	UseSendMethod := 0
	WinSurgeLoginID = %y%
	WinSurgeFullName = %Result_1%
	SecurityUserId = %Result_4%
	WinSURGEPathologistID = %Result_5%
	CarisRocketWindowX := 0
	CarisRocketWindowY := 0
	CarisRocketWindowW := 0
	CarisRocketWindowH := 0
	
	
	OpMode = D

	IniWrite, %A_Username%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WindowsLoginId
	IniWrite, %WinSurgeFullName%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSurgeFullName
	IniWrite, %OpMode%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, OpMode
	IniWrite, %SecurityUserId%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, SecurityUserId
	IniWrite, %WinSURGEPathologistID%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSURGEPathologistID
	IniWrite, %CarisRocketWindowX%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowX
	IniWrite, %CarisRocketWindowY%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowY
	IniWrite, %CarisRocketWindowW%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowW
	IniWrite, %CarisRocketWindowH%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowH

	Msgbox, 4, Setup..., Caris CodeRocket is installed but is not setup for "automated" modes on your machine.  You must go through setup to use the automated modes or you can skip setup and use in manual mode.  Continue to setup for automation? 
	ifMsgbox, No
		{
			OpMode=M
			IniWrite, %OpMode%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, OpMode			
			Return
		}
		
			

	OpenCase("CD11-1000")
	Sleep, 1500

qiblabel:
	SplashTextOn, 200,100, Setup..., Click your mouse into the "Queue Into Batch" box and press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	WinGetTitle, active_title, A
	ControlGetFocus, x,	%active_title%
	StringGetPos, y, x, TextBox
	if (y>0)
		{
			QueueIntoBatchBox := x	
			IniWrite, %QueueIntoBatchBox%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, QueueIntoBatchBox
		}
	Else
		{
			Msgbox, You did not click into a text box.  Please click your mouse into the "Queue Into Batch" box and press F3.
			Goto, qiblabel
		}
		
p1label:
	SplashTextOn, 200,100, Setup..., Click your mouse into the "Image01:" box and press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	WinWait, WinSURGE WinsIMAGE
	WinClose, WinSURGE WinsIMAGE
	WinWaitClose, WinSURGE WinsIMAGE
	Sleep, 500
	WinGetTitle, active_title, A
	ControlGetFocus, x,	%active_title%
	StringGetPos, y, x, TextBox
	if (y>0)
		{
			Photo1TextBox := x	
			IniWrite, %Photo1TextBox%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, Photo1TextBox
		}
	Else
		{
			Msgbox, You did not click into a text box.  Please click your mouse into the "Image01" box and press F3.
			Goto, p1label
		}

p2label:
	SplashTextOn, 200,100, Setup..., Click your mouse into the "Image02:" box and press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	WinWait, WinSURGE WinsIMAGE
	WinClose, WinSURGE WinsIMAGE
	WinWaitClose, WinSURGE WinsIMAGE
	Sleep, 500
	WinGetTitle, active_title, A
	ControlGetFocus, x,	%active_title%
	StringGetPos, y, x, TextBox
	if (y>0)
		{
			Photo2TextBox := x	
			IniWrite, %Photo2TextBox%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, Photo2TextBox
		}
	Else
		{
			Msgbox, You did not click into a text box.  Please click your mouse into the "Image02" box and press F3.
			Goto, p2label
		}

pathlabel:
	SplashTextOn, 200,100, Setup..., Click your mouse into the "Pathologist" box and press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	WinGetTitle, active_title, A
	ControlGetFocus, x,	%active_title%
	StringGetPos, y, x, TextBox
	if (y>0)
		{
			PathologistTextBox := x	
			IniWrite, %PathologistTextBox%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, PathologistTextBox
		}
	Else
		{
			Msgbox, You did not click into a text box.  Please click your mouse into the "Pathologist" box and press F3.
			Goto, pathlabel
		}

finaldxlabel:
	SplashTextOn, 200,100, Setup..., Hover your mouse over the "Final Diagnosis" button BUT DONT CLICK IT!  Then press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	MouseGetPos, xpos, ypos	
	FinalDiagnosisButtonX := xpos
	FinalDiagnosisButtonY := ypos
	IniWrite, %FinalDiagnosisButtonX%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, FinalDiagnosisButtonX
	IniWrite, %FinalDiagnosisButtonY%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, FinalDiagnosisButtonY

grossdesclabel:
	SplashTextOn, 200,100, Setup..., Hover your mouse over the "Gross Description" button BUT DONT CLICK IT!  Then press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	MouseGetPos, xpos, ypos	
	GrossDescriptionButtonX := xpos
	GrossDescriptionButtonY := ypos
	IniWrite, %GrossDescriptionButtonX%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, GrossDescriptionButtonX
	IniWrite, %GrossDescriptionButtonY%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, GrossDescriptionButtonY
	
winSurgeSignoutPassword:
	SplashTextOff
	InputBox, wssop1, WinSurge Signout Password, Enter your WinSURGE signout password, HIDE
	InputBox, wssop2, WinSurge Signout Password, Enter your WinSURGE signout password again for verification, HIDE
	
	if (wssop1=wssop2)
	{
		IniWrite, %wssop1%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSurgeSignoutPassword
		WinSurgeSignoutPassword := wssop1
	}
	else
	{
		Msgbox,  Your entered passwords did not match.  Please enter them again.
		Goto, winSurgeSignoutPassword
	}

	SplashTextOn, 200, 100, Setup Complete!
	Sleep, 1500
	SplashTextOff
Return
}

ReadIniValues()
{
	global

	IniRead, WindowsLoginId, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WindowsLoginId
	IniRead, OpMode, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, OpMode
	IniRead, WinSurgeSignoutPassword, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSurgeSignoutPassword
	IniRead, WinSurgeLoginPassword, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSurgeLoginPassword
	IniRead, WinSurgeFullName, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSurgeFullName
	IniRead, HomeLabCasePrefix, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, HomeLabCasePrefix
	IniRead, HomeLabGrossAddendum, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, HomeLabGrossAddendum
	IniRead, WinSURGEPathologistID, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSURGEPathologistID
	IniRead, SecurityUserId, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, SecurityUserId

	IniRead, QueueIntoBatchBox, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, QueueIntoBatchBox 	
	IniRead, Photo1TextBox, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, Photo1TextBox 	
	IniRead, Photo2TextBox, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, Photo2TextBox 	
	IniRead, PathologistTextBox, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, PathologistTextBox 	
	IniRead, FinalDiagnosisButtonX, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, FinalDiagnosisButtonX
	IniRead, FinalDiagnosisButtonY, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, FinalDiagnosisButtonY 	
	IniRead, GrossDescriptionButtonX, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, GrossDescriptionButtonX 	
	IniRead, GrossDescriptionButtonY, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, GrossDescriptionButtonY 	
	IniRead, DictationSendX, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, DictationSendX 	
	IniRead, DictationSendY, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, DictationSendY 	

	IniRead, CarisRocketWindowX, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowX 	
	IniRead, CarisRocketWindowY, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowY 
	IniRead, CarisRocketWindowW, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowW 
	IniRead, CarisRocketWindowH, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowH

	IniRead, EsignoutTextBox, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, EsignoutTextBox
	IniRead, ApprovalPasswordControl, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, ApprovalPasswordControl 

	IniRead, ExpressSignout, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, ExpressSignout
	IniRead, SpeakEnabled, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, SpeakEnabled
	IniRead, UseSendMethod, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, UseSendMethod
	IniRead, BeepOnShiftEnter, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, BeepOnShiftEnter

Return
}

WinSurgeSetup() 
{
	global
	IfWinNotExist, WinSURGE
	{
		Run, %WinSurgeFilePath%		
		WinWait, WinSURGE Environment Login, , 2
		If ErrorLevel
			{
				Msgbox, Error Starting WinSurge
				ExitApp
			}
		Else
			{
			IfWinNotActive, WinSURGE Environment Login, , WinActivate, WinSURGE Environment Login, 	
			WinWaitActive, WinSURGE Environment Login, 
			Send, %WinSurgeLoginPassword%{ENTER}
			WinWait, Login Message, 
			IfWinNotActive, Login Message, , WinActivate, Login Message, 
			WinWaitActive, Login Message, 
			Send, {ENTER}
			}
	}
	Else
		Msgbox, WinSurge is currently running on your machine.  Some WinSurge errors close the visible windows but do not close the program.  If you get this message but do not see the windows, hit Ctrl-Alt-Delete and then Task Manager and use it to close WinSurge when this occurs.  Then reopen WinSURGE.
	Return 0
}

OpenAutoAssign() 
{
	global
	Run, http://autoassign/autoassign2/default.php
	WinWait, Caris AutoAssign Main Menu
	IfWinNotActive, Caris AutoAssign Main Menu
		WinActivate, Caris AutoAssign Main Menu	
	WinWaitActive, Caris AutoAssign Main Menu
	Send, %AutoAssignUsername%{Tab}
	Sleep, 500
	Send, {Enter}
	Sleep, 500
	Run, http://autoassign/autoassign2/report_path_case_status.php
	Return
}

MildDysplasticWarningCheck()
{
global
	If (ClientID="CT5995D" OR ClientID="RI7604D" OR ClientID="RI10484D" OR ClientID="RI10659D")
					{
						comma =,
						IfInString, mildCodes, %comma%%basediag%%comma%
							{
								SoundBeep
								SoundBeep
								SoundBeep
								Msgbox, 4,, This client has requested never to receive a melanocytic lesion with mild atypia and you used one.  Are you sure you want to continue with this code?
								IfMsgbox, No
									Return 1
							}
					}

return 0
}

JerriJohnsonMarginCheck()
{
	global
						If (ClientName="Johnson, Jerri")
						IfInstring, thisJarClinicalData, margins
						{
 						
						SoundBeep
						Msgbox, 4,, Jerri Johnson has requested margins on this case in the clinical history.  Are you sure you want to continue without one?
						IfMsgbox, No
 							Return 1
 						}

Return 0
}
  ; End of Internal Functions

