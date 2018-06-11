/*  WISH LIST    
/*  Preference Checking: 
 * Move Current functionality to Firebase
 * Use an external web page to enter preferences into a Firebase database.
 * Store the preferences that were used to generate the database entries and monitor if they have changed.
 * 
 * Create a master list of hashtags to tag DX Codes with
 * #normal #abnormal #tumor #malignant  #benign #NMSC #BCC #SCC #Melanoma #allDN #mildDN, #atypicalMelanocyticProliferation #inflammatory #rash #hematolymphoid
 * 
 * Allow granularity on the edit page, hashtags Y/N Any DX code added specifically
 * Allow the same granularity for recomendations
 * 
 * Do not use list: #hashtage, DXCODES, helperphrases, "literal phrases"
 * 
 * Micros  Same granularity as margins
 * Photos:  Same granularity as margins
 
	Explore separating the Coding, Automation, and Paper Replacer Functions
	Pull MsgToResults() into a function?

	Move Smart Substitution Editing to a webpage for html formatting
 
 */

*/
/*  TO-DO List
Move Patient Data into the WebView
Add DoNotUse Hotstring functionality could check on F8 also
Add CPT Checking support

Add A-Z replacements for all data
Add ProcedureNote Flag 

{
Standardize the ReDraw so that substitutions can only occur once.
Add more margin flags
Add a tooltip to the margin flag so that hover will show the margin flags or just add to the html the margin flags
}



*/

Startup:         ;MS done
{

mildCodes=,cnmild,cnmildr,cnmi,jnmild,jnmildr,jnmi,jnfs,cnfs,lcnmi,ljnmi,nljn,nlcn,jnami,cnami,
CurrentCodeRocketDisplayedCase = FirstTime   ;Prevents an error on startup
namedDxCodes := {}

#SingleInstance force
#WinActivateForce
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetTitleMatchMode 1
CoordMode, Mouse, Relative
ComObjError(false)

IfExist, %A_MyDocuments%\CarisCodeRocket.ini
	{
		ReadIniValues()
		LoadSmartSubstitutions()
	}
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
	

	Gui, Font, S12, Arial
	Gui, Add, Text, vCaseLoaderLbl,Case Loader:
	Gui, Add, Edit, vCaseScanBox ys,
	Gui, Add, Button, w0 h0 Default, OK
	Gui, Font, S24, Arial
	Gui, Add, Text, cBlue x580 y10 w200 vCaseNumberLabel, Case Number: %CurrentCaseNumber%

	Gui, Font, S14, Arial
	Gui, Add, Text, r1 cBlue vUsePhotos, <PHOTOS REQUIRED>
	Gui, Add, Text, r1 cRED vUseMicros, <MICROS REQUIRED>
	;Gui, Add, Text, r1 cBlack w500 vUseMargins, <Margin Preferences appear here>

	Gui, Add, Text, x10 y50 w500 vPatientLabel, Patient: %PatientName% --- Age:%PatientAge%
	Gui, Add, Text, r1 w500 vDoctorLabel, Doctor: %ClientName%
	Gui, Add, Text, r1 w500 vClientLabel, Client:  %ClientOfficeName%  --- %ClientState%
	gui, add, text, x10 y+10 w800 h1 0x7  ;Horizontal Line > Black
	html=<span id='main'><span id='orderedProcedures'></span><span id='attnPathologist' style='color:red'></span><strong>Preferences:</strong><br><span id='preferences'></span><br><br><strong>Clinical Information:<br></strong><span id='clinicalInformation' style="color:blue"></span><br><br><strong>Final Diagnosis:<br></strong><span id='finalDiagnosis' style="color:green"></span><br><br><strong>Gross Description:<br></strong><span id='grossDescription'></span><br><br><strong>Prior Case Information<br></strong><span id='priorCaseInformation'></span></span>

	Gui, Add, ActiveX, w800 h590 vWB hwndATLWinHWND, Shell.Explorer
	
	;These Lines are required for the copy function to work inside the browerWindow
	IOleInPlaceActiveObject_Interface:="{00000117-0000-0000-C000-000000000046}"
	pipa := ComObjQuery(WB, IOleInPlaceActiveObject_Interface)
	OnMessage(WM_KEYDOWN:=0x0100, "WM_KEYDOWN")
	OnMessage(WM_KEYUP:=0x0101, "WM_KEYDOWN")
	OnMessage(0x5000, Speak)
	
	WB.Navigate("about:blank")
	WB.document.write(html)
	
	Menu, FileMenu, Add, E&xit, GuiClose
	Menu, HelpMenu, Add, Search Diagnosis Codes  (F7), F7
	Menu, HelpMenu, Add, Search Extended Phrases  (Shift-F7), +F7
	Menu, HelpMenu, Add, Display All Helpers  (F9), F9

	Menu, SettingsMenu, Add, Speak Patient Name, SpeakPatientName
	Menu, SettingsMenu, Add, Beep On Shift-Enter, BeepOnShiftEnter
	Menu, SettingsMenu, Add, Use Smart Substitutions, UseSmartSubstitutions

	If SpeakEnabled
		Menu, SettingsMenu, Check, Speak Patient Name
	Else
		Menu, SettingsMenu, UnCheck, Speak Patient Name

	If BeepOnShiftEnter
		Menu, SettingsMenu, Check, Beep On Shift-Enter
	Else
		Menu, SettingsMenu, UnCheck, Beep On Shift-Enter

	If UseSmartSubstitutions
		Menu, SettingsMenu, Check, Use Smart Substitutions
	Else
		Menu, SettingsMenu, UnCheck, Use Smart Substitutions

	Menu, EditMenu, Add, Edit Client Preferences, EditDoctorPreferences

	Menu, MyMenuBar, Add, &File, :FileMenu  
	Menu, MyMenuBar, Add, &Edit, :EditMenu
	Menu, MyMenuBar, Add, &Settings, :SettingsMenu
	Menu, MyMenuBar, Add, &Help, :HelpMenu
	Gui, Menu, MyMenuBar

	Gui, 1:+Resize +MinSize850x850 +MaxSize850x850
	Gui, Show, x%CarisRocketWindowX% y%CarisRocketWindowY% w%CarisRocketWindowW% h%CarisRocketWindowH%

	Progress, 0 x400 y1 h130, Preparing for first time use..., Written by Matthew Muenster M.D.`n`nInitializing..., Caris CodeRocket 
	Progress, 40, Reading personalized values...
	Progress, 60, Getting the diagnosis codes from the database...
	ReadDXCodes()
	Progress, 80, Getting the helper codes from the database...
	ReadHelpers()
	Progress, 100, Initialization complete!
	Progress, Off

	SetTimer, WinSURGECaseDataUpdater, 2000

	Gosub, WinSurgeCaseDataUpdater

	Run, S:\CodeRocket\bin\EP\Autohotkey.exe S:\CodeRocket\bin\EP\DermpathExtendedPhrases.ahk, ,UseErrorLevel , epid
	If ErrorLevel
		MsgBox, 4112, Connectivity Error, Your connection to the S:\CodeRocket directory is not present.  Usually`, restarting your computer will correct this.  You can use the CodeRocket program but will have no extended phrase capabilities.

	LoadPersonalExtendedPhrases()
	
	;Prepare Text to Speech Function
	SAPI := ComObjCreate("SAPI.SpVoice")
	SAPI.rate := -2

	return
}

ButtonOK:  ;Automation
{

	SetTimer, WinSurgeCaseDataUpdater, Off
	
	Gui, Submit, NoHide
	If (UndoEnabled AND !DataEntered)
		{
			Msgbox, You must first save or undo the changes to the current case!
			Gosub, F12
			Return
		}
	
	foundCase := RegExMatch(CaseScanBox, "[A-Za-z][A-Za-z]\d\d-\d+", NewCaseNum)

	if !foundCase
		{
			Msgbox, You did not enter a valid case number!
			Gosub, F12
			Return
		}

	If (DataEntered AND UndoEnabled)
		{
			Gosub, F8
			If SaveError
			{
				Gosub, F12
				Return
			}
		}
		else
		{
			If (CurrentCodeRocketDisplayedCase<>NewCaseNum)
			{
				z1:=NewCaseNum
				Gosub, BuildMainGui
			}
		}
	
	CloseWinSURGEModalWindow("WinSURGE - Final Diagnosis:","","Close")

	SetTimer, UnblockInput, 5000
	Gosub, F12
	BlockInput, On
	OpenCase(NewCaseNum)
	DataEntered := 0
	OpenFinalDiagnosisModal()
	ActivateNextTripleAsterisk() 
	
	If SpeakEnabled
	{
		StringSplit, name, PatientName, `,
		p = %name2% %name1%, Age: %PatientAge%
		StringUpper, p, p, T
		SAPI.speak(p)
	}
	
	SetTimer, WinSURGECaseDataUpdater, 2000
	BlockInput, Off
	SetTimer, UnblockInput, Off 	

	Return
}

BuildMainGui:  ;Paper Replacer
{
	preferences := ""
	priorCaseInfo := ""
	orderedProcedures := ""
	
	if (z1="")
	{
		z1 := "No Current Case"
		RedrawGui()
		return
	}

	if(z1=CurrentCodeRocketDisplayedCase)
		return
	
	CurrentCodeRocketDisplayedCase := z1

	x := get_filled_case_number(CurrentCodeRocketDisplayedCase)
		
		s := "select s.dx, s.gross, s.numberofspecimenparts, s.custom03, s.clin, p.name, s.clindata, pt.name, s.Computed_PATIENTAGE, p.proficiencylog, p.comment, s.custom04, s.patient, s.zfield, s.Computed_PatientDOB from specimen s, physician p, patient pt where s.patient = pt.id and s.clin=p.id and computed_numberfilled='" . x . "'"

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
		rawPreferences := Result_10
		ClientOfficeName := Result_11
		ClientID := Result_12
		StringReplace, ClientID, ClientID, `n,,All
		StringSplit, PatientID, Result_13, .    ; The patient ID is in a variable called PatientID1
		attnPathologistField := Result_14
		PatientDOB := Result_15
		
		;Mandatory Replacements for basic formatting start here
		attnPathologistField := RegExReplace(attnPathologistField, "[a-z]+ \d+\/\d+\/\d+ \d+:\d+ \w+", "")
		StringReplace, attnPathologistField, attnPathologistField, `r,,All
		StringReplace, attnPathologistField, attnPathologistField, `n,,All
		
		StringLeft, ClientState, ClientID, 2
		
		StringSplit, PhysicianWinSurgeId, ClientWinSurgeId, .  ;physican WinSurge Id is stored in PhysicianWinSurgeId1
		
		StringReplace, preferences, preferences,`r,,All
		StringReplace, preferences, preferences,¥,,All
		StringReplace, preferences, preferences,***,<br>, All

		StringReplace, ClientOfficeName, ClientOfficeName, &, &&, All
		StringReplace, ClientOfficeName, ClientOfficeName, -Att, ,
		
		STringReplace, finaldiagnosistext, finaldiagnosistext, `%`%P`%`%%A_Space%, `%`%P`%`%<br>, All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%B., <br><br>B., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%C., <br><br>C., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%D., <br><br>D., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%E., <br><br>E., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%F., <br><br>F., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%G., <br><br>G., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%H., <br><br>H., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%I., <br><br>I., All
		STringReplace, finaldiagnosistext, finaldiagnosistext, %A_Space%Comment:, <br>Comment:, All
		STringReplace, finaldiagnosistext, finaldiagnosistext, <br>Comment:, <br><br>Comment:, All
		finaldiagnosistext := RegExReplace(finaldiagnosistext,"%%\w+%%","")
		
		StringReplace, tempVar, finaldiagnosistext, <br>, ¥, All
		siteList:= ""
		clinicalSiteArray := []
		Loop, Parse, tempVar, ¥
			{
				j:=RegExMatch(A_LoopField, "[A-Z]\. (.*?)[:,]", Subpat)
				if(Subpat1)
					siteList=%siteList%,%Subpat1%
				
			}

		STringReplace, ClinicalData, ClinicalData, %A_Space%B., <br>B., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%C., <br>C., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%D., <br>D., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%E., <br>E., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%F., <br>F., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%G., <br>G., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%H., <br>H., All
		STringReplace, ClinicalData, ClinicalData, %A_Space%I., <br>I., All

		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%A., <br>A., 
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%B., <br>B.,
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%C., <br>C.,
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%D., <br>D.,
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%E., <br>E.,
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%F., <br>F.,
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%G., <br>G.,
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%H., <br>H.,
		STringReplace, grossdescriptiontext, grossdescriptiontext, %A_Space%I., <br>I.,
		grossdescriptiontext := RegExReplace(grossdescriptiontext,"\([A-Z][A-Z]\/[a-z]+ \d+\/\d+\/\d+ \d+:\d+ \w+\)","")
		grossdescriptiontext := RegExReplace(grossdescriptiontext,"\([a-z]+ \d+\/\d+\/\d+ \d+:\d+ \w+\)","")

		displayedPreferences := preferences
		displayedClinicalData := ClinicalData
		displayedFinalDiagnosis := finaldiagnosistext
		displayedGrossDescription := grossdescriptiontext
		
		displayedClinicalData := RegExReplace(displayedClinicalData, "(([P|p]lease)?\s?(?i)(check margins\W?))", "<strong style='color:Red'>$0</strong>")
		ReDrawGui()
		
		;This section looks for "client" specific preferences and adds them to the preferences box if they exist
		additionalPreferences := ""
		s=select p.proficiencylog from physician p where p.number='%ClientID%'
		WinSurgeQuery(s)

		If msg
			{
			If(rawPreferences<>"")
				FoundPos1 := InStr(msg, rawPreferences)
			else
				FoundPos1 := 0
			if(msg<>"")
				FoundPos2 := InStr(rawPreferences, msg)
			else
				FoundPos2 := 0

			StringReplace, msg, msg, `n,<br>,All
			StringReplace, msg, msg, ***,<br>, All
			StringReplace, msg, msg, ¥,,All
			StringReplace, msg, msg, `r,,All

			if(!(FoundPos1 OR FoundPos2))
				{
				rawPreferences=%rawPreferences%%msg%
				displayedPreferences=%displayedPreferences%<br>%msg%
				RedrawGui()
				}
			else if(FoundPos1>0 AND FoundPos2=0)
				{
				rawPreferences=%msg%
				displayedPreferences=%msg%
				RedrawGui()
				}
			}
	
	;This section gets the ordered procedures
	s := "select s.* from specimen s where computed_numberfilled='" . x . "'"
	WinSurgeQuery(s)
	StringSplit, oput, msg, ¥
	orderedProcedures := oput515
	orderedProcedures := RegExReplace(orderedProcedures,"L 1-3#\d+,","")
	orderedProcedures := RegExReplace(orderedProcedures,"L 1-3#\d+","")
	orderedProcedures := RegExReplace(orderedProcedures,"L 1-3,","")
	orderedProcedures := RegExReplace(orderedProcedures,"L 1-3","")
	
	StringReplace, orderedProcedures, orderedProcedures, DLx2,Deepers x 2,All
	StringReplace, orderedProcedures, orderedProcedures, DLx3,Deepers x 3,All
	StringReplace, orderedProcedures, orderedProcedures, DLx4,Deepers x 4,All
	StringReplace, orderedProcedures, orderedProcedures, DLx6,Deepers x 6,All
	StringReplace, orderedProcedures, orderedProcedures, DLx8,Deepers x 8,All
	StringReplace, orderedProcedures, orderedProcedures, m1mar,Mart-1,All

	


	;PRIOR CASE INFO - This sections searches for and formats the prior cases info
	perc := "%%"
	msg := ""

	
	if(PatientID1)
	{
		s := "select s.number, s.sodate, p.name, s.dx from specimen s, physician p where s.patient=" . PatientID1 . " and s.path=p.id"
		WinSurgeQuery(s)

		if msg
		{
			;This sorts the msg by the signout date of the cases
			sortPriors := []
			Loop, Parse, msg, `n
				{
					StringSplit, oput, A_LoopField, ¥
					StringSplit, soDateArray, oput3, /
					If(soDateArray1<10)
						soDateArray1=0%soDateArray1%
					If(soDateArray2<10)
						soDateArray2=0%soDateArray2%
						
					soDate=%soDateArray3%%soDateArray1%%soDateArray2%
					;Msgbox, %soDate%
					sortPriors[A_Index]:= { "caseNumber":oput2, "signoutDate":soDate, "signoutPathologist":oput4, "signedOutText":oput5, "rawText":A_LoopField }
				}
				
				For key, val in sortPriors
				{
					Loop, % sortPriors.Length()-key
					{
						;Msgbox, key=%key%, A_Index=%A_Index%
						if(sortPriors[key].signoutDate < sortPriors[key+A_Index-1].signoutDate)
						{
								tempVar := sortPriors[key]
								sortPriors[key] := sortPriors[key+A_Index-1]
								sortPriors[key+A_Index-1] := tempVar
						}
					}
				}	

			msg := ""
			For key, val in sortPriors
				msg := msg . val.rawText . "`n"


			Loop, Parse, msg, `n
			{
					If A_LoopField
					{
						FoundPos := RegExMatch(A_LoopField, "¥[A-Z][A-Z]\d\d-\d+¥")
						If FoundPos>0
						{
							IfInString, A_LoopField, %CurrentCaseNumber%   ;This keeps the current case from appearing in the prior case information
								continue
							else
								{
									
									p:=RegExReplace(A_LoopField,"¥[A-Z][A-Z]\d\d-\d+¥","<u>$0</u>")
									priorCaseLine=<br>%p%
								}
						}
						else
							{
								priorCaseLine=%A_LoopField%
							}
					
						
						StringReplace, priorCaseLine, priorCaseLine, %perc%P%perc%,<br>, All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88305%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88304%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88312%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88346%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88350%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, %perc%88342%perc%, , All
						StringReplace, priorCaseLine, priorCaseLine, Comment:, <br><br>Comment: , All
						StringReplace, priorCaseLine, priorCaseLine, ¥A., <br>A., All
						StringReplace, priorCaseLine, priorCaseLine, ¥, %A_Space%-%A_Space%, All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%B., <br><br>B., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%C., <br><br>C., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%D., <br><br>D., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%E., <br><br>E., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%F., <br><br>F., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%G., <br><br>G., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%H., <br><br>H., All
						StringReplace, priorCaseLine, priorCaseLine, %A_Space%I., <br><br>I., All
						
						priorCaseInfo = %priorCaseInfo%<br>%priorCaseLine%
					}
			}
			priorCaseInfo := Trim(RegExReplace(priorCaseInfo,"%%\w+%%",""))
			
			if (priorCaseInfo="")
				priorCaseInfo := "None"

			; This section does the highlighting of identical case numbers in the clinical Info and the prior cases.
			FoundPos := RegExMatch(ClinicalData, "[A-Z][A-Z]\d\d-\d+", foundCaseNum)

			If FoundPos
				{
				priorCaseInfo:=RegExReplace(priorCaseInfo, foundCaseNum,"<span id='" . idTag . "' style='background-color:yellow'><strong>$0</strong></span>",ReplacementCount)
				If(ReplacementCount)
					displayedClinicalData := RegExReplace(displayedClinicalData, foundCaseNum, "<span id='" . idTag . "' style='background-color:yellow'><strong>$0</strong></span>")
				}

			; This section does the highlighting of identical sites in the final diagnosis and the prior cases.
			Loop, Parse, siteList, `,
			{
				StringReplace, idTag, A_LoopField, %A_Space%,,All
				If (A_Index=1)
					continue
				priorCaseInfo:=RegExReplace(priorCaseInfo, A_LoopField,"<span id='" . idTag . "' style='background-color:yellow'><strong>$0</strong></span>",ReplacementCount)
				If(ReplacementCount)
					StringReplace, displayedFinalDiagnosis, displayedFinalDiagnosis, %A_LoopField%, <span id='" . idTag . "' style='background-color:yellow'><strong>%A_LoopField%</strong></span>
			}

			j:=StrLen(priorCaseInfo)
			StringLeft, firstEight, priorCaseInfo, 8
			If (firstEight="<br><br>")
				StringRight, priorCaseInfo, priorCaseInfo, j-8
			
		}
		else
			priorCaseInfo=None
	}
		
	ReDrawGui()
	;END of PRIOR CASE INFO

			GuiControl, 1:Hide, UsePhotos
			GuiControl, 1:Hide, UseMicros
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
					Gosub, EditDoctorPreferences
					GuiControl, 2:, PhotoSelect, 0
					GuiControl, 2:, MicroSelect, 0
					;GuiControl, 2:, ICD9Select, 0
					GuiControl, 2:Text, DocPreferenceLabel, Enter preferences for %ClientName%
					Gui, 2:Show, , CarisDemo			;x%CarisRocketWindowX% y%CarisRocketWindowY% w%CarisRocketWindowW% h%CarisRocketWindowH%
					Gui, 2:-Disabled +AlwaysOnTop
					SoundBeep
					SoundBeep
				}

		
return
}

WinSURGECaseDataUpdater:  ;Automation
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

	; This sections checks the status of the WinSURGE Window and formats accordingly
	ifWinNotExist, WinSURGE
		{
		lastWinSurgeTitle=""
		z1 := "WinSURGE Not Open"
		Gosub, BuildMainGui
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

			;This disables the case scan box when there has been a change to WinSurge that did not use the coderocket Shift-Enter function
			if(UndoEnabled AND !DataEntered)
				{
				GuiControl, Disable, CaseScanBox
				GuiControl, Disable, CaseLoaderLbl
				GuiControl, Text, CaseScanBox, Data in Case
			}
			Else 
			{
				Gui, Submit, NoHide
				GuiControl, Enable, CaseScanBox
				GuiControl, Enable, CaseLoaderLbl
				if CaseScanBox=Data in Case
					GuiControl, Text, CaseScanBox, 
			}
			}

			;Gets the title of the WinSurge Window
			WinGetTitle, x, WinSURGE, &2 Open Case
			
			;Return if CodeRocket is already displaying the proper case (from ButtonOK)
			IfInString, x, %CurrentCodeRocketDisplayedCase%
			{
				lastWinSurgeTitle := x
				return
			}

			If(x=lastWinSurgeTitle)
				Return

			StringGetPos, y, x, No Current Case
			StringGetPos, z, x, New
			z := y + z

			If (z>0 OR x="WinSURGE")
				{
				lastWinSURGEtitle := x
				z1 := "No Current Case"
				Gosub, BuildMainGui	
				Return
				}
				
			lastWinSURGEtitle := x
			DataEntered = 0
			StringReplace, x, x, Case, |, All
			StringSplit, y, x, |, %A_Space%
			StringSplit, z, y2, %A_Space%, %A_space%
			;z1 := RegExMatch()
			Gosub, BuildMainGui
	}
	return
}

SpeakPatientName:    ;Automation
{
	Menu, SettingsMenu, ToggleCheck, Speak Patient Name
	If SpeakEnabled
		SpeakEnabled := 0
	Else
		SpeakEnabled := 1
	IniWrite, %SpeakEnabled%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, SpeakEnabled
Return
}

BeepOnShiftEnter:    ;Automation
{
	Menu, SettingsMenu, ToggleCheck, Beep On Shift-Enter
	If BeepOnShiftEnter
		BeepOnShiftEnter := 0
	else
		BeepOnShiftEnter := 1
	
	IniWrite, %BeepOnShiftEnter%,  %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, BeepOnShiftEnter
	
	return
}

UseSmartSubstitutions:  	;Paper replacer
{
		Menu, SettingsMenu, ToggleCheck, Use Smart Substitutions
	If UseSmartSubstitutions
		{
			UseSmartSubstitutions := 0
			CurrentCodeRocketDisplayedCase := 0
			Gosub, BuildMainGui
		}
	else
		{
			UseSmartSubstitutions := 1
			CurrentCodeRocketDisplayedCase := 0
			Gosub, BuildMainGui
		}
	
	IniWrite, %UseSmartSubstitutions%,  %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, UseSmartSubstitutions
	
	return
}

EditDoctorPreferences:		;Automation
{
	
	Gui, 2:Destroy
	Gui, 2:Font, S12, Verdana
	Gui, 2:Add, Text, x18 vDocPreferenceLabel, TestNameXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
	Gui, 2:Font, S8, Verdana
	Gui, 2:Add, Checkbox, vPhotoSelect, Photos Required
	Gui, 2:Add, Checkbox, vMicroSelect, Micros Required
	;Gui, 2:Add, Checkbox, vICD9Select, ICD9s Required
	Gui, 2:Add, Text, , Margin Preferences
	Gui, 2:Add, Edit, vMarginSelect w500, 
	Gui, 2:Add, Button, gSavePreferences vGo, Save Preferences

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
		return
}
	
SavePreferences:   			;Automation
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

UnblockInput:			;Automation
{
	BlockInput, Off
	return
}

GuiClose:				;Automation or synthesis
{
	ObjRelease(pipa)
	Process, Close, %epid%, 
	Process, Close, %ppid%,
	;Process, Close, Autohotkey.exe,
	
	If ErrorLevel
		Process, Close, %ErrorLevel%, 
	ExitApp	
	Return
}

SendEmail:	;Automation
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

5ButtonClose:	;Coding
{
Gui, 5:Destroy
return
}

11ButtonSave:
{
	Gui, 11:Submit
	tempVar := RegExMatch(FlagSelector, "(\d+):", numberOfSelection)
	If (numberOfSelection1=1)
	{
		replacementText := "<strong style='background-color:lightblue'>" . tempStringtoFlag . "</strong>"
		FileAppend, "%tempStringtoFlag%"`,"%replacementText%"`,"no-margins-nmsc"`n, %A_MyDocuments%\MySmartSubstitutions.csv
	} 
	else if (numberOfSelection1=3)
	{
		replacementText := "<strong style='background-color:lightblue'>" . tempStringtoFlag . "</strong>"
		FileAppend, "%tempStringtoFlag%"`,"%replacementText%"`,"none"`n, %A_MyDocuments%\MySmartSubstitutions.csv
	}
	LoadSmartSubstitutions()
	ReDrawGui()
	CurrentCodeRocketDisplayedCase := 0
	Gui 11:Destroy
	return
}

11ButtonCancel:
{
	Gui 11:Destroy
	return
}

F7::            ;Coding
{	
	counter := 0
	dt := ""
	InputBox, SearchWord, Diagnostic Code Search, Enter the single word or phrase you want to search for:
	if SearchWord
		{
	
	For key, val in namedDxCodes
		{
		RetrievedText1 := namedDXCodes[key]["code"]
		RetrievedText2 := namedDXCodes[key]["dxLine"]
		RetrievedText3 := namedDXCodes[key]["comment"]
		RetrievedText4 := namedDXCodes[key]["micro"]
		StringGetPos, t1, RetrievedText1, %SearchWord%
		StringGetPos, t2, RetrievedText2, %SearchWord%
		StringGetPos, t3, RetrievedText3, %SearchWord%
		StringGetPos, t4, RetrievedText4, %SearchWord%
		;Msgbox, %t1%,%t2%,%t3%,%t4%
		if (t1>-1 or t2>-1 or t3>-1 or t4>-1)
			{
			counter := counter + 1
			if (counter=1)
				dt=<h1>DIAGNOSIS CODE SEARCH RESULTS</h1><br><h3>CounterHolder Code(s) Found</h3><br><br>
			StringUpper, RetrievedText1, RetrievedText1
			dt = %dt%<strong>%RetrievedText1%</strong> <br>%RetrievedText2%<br>
			if (RetrievedText3 OR RetrievedText4)
				dt = %dt%Comment:  %RetrievedText3%  %RetrievedText4%<br><br>
			}
		}
	StringReplace, dt, dt, CounterHolder, %counter%
	StringUpper, tempVar, SearchWord
	dt := RegExReplace(dt, tempVar, "<span style='background-color:yellow'><strong>$0</strong></span>")
	StringLower, tempVar, SearchWord
	dt := RegExReplace(dt, tempVar, "<span style='background-color:yellow'><strong>$0</strong></span>")
	
	Gui, 5:Destroy
	Gui, 5:Add, ActiveX, w800 h590 vDxCodeSearch, Shell.Explorer
	DXCodeSearch.Navigate("about:blank")
	DXCodeSearch.document.write(dt)
	Gui, 5:Add, Button, Default, Close
	Gui, 5:Show
}
Return
}

+F7::			;Coding
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

F8::           ;Automation
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
	if (UsePhotos)
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
	if (NewCaseNum)
		z1 := NewCaseNum
	
	Gosub, BuildMainGui
	
	QueueandAssign()
	CloseandSaveCase()
	DataEntered = 0
	Gosub, F12
	Return
}

+F8::			;Automation
{
	CloseandSaveCase()	
	DataEntered = 0
	GuiControl, Hide, StatusLabel
	Gui, Show, NoActivate
	Gosub, F12
	Return
}

F9::			;Coding
{
	Gui 3:Destroy

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
	Gui, 3:Font, 
	Gui, 3:show, ,List of Available Helpers
Return
}

F11::			;Automation
{
	IfWinExist, WinSURGE - 
	{
		Send, %LastCodeUsed%	
		Gosub, Shift & Enter
	}
	Return
}

F12::			;Automation			
{
	WinActivate, WinSURGE , 	
	SetTitleMatchMode, 2
	WinActivate, CodeRocket, , SciTE4AutoHotkey
	SetTitleMatchMode, 1
	GuiControl, Text, CaseScanBox,  ;Blanks the data entry textbox 	
	GuiControl, Focus, CaseScanBox,
	
return
}

^!f::  ;Smart Flag Replacer
{
	tempStringtoFlag := clipboard
	StringReplace, tempStringtoFlag, tempStringtoFlag,`n,,All
	StringReplace, tempStringtoFlag, tempStringtoFlag,`r,,All
	Gui 11:Destroy
	Gui 11:Add, Text , , Which flag would you like to add to the phrase "%tempStringtoFlag%"?
	Gui 11:Add, DropDownList, w400 vFlagSelector, 1:No Margins on NMSC|2:Margins on All Nevi|3:No margins except excisions and when requested
	
	Gui, 11:Add, Button,  Default, Save
	Gui, 11:Add, Button, , Cancel
	
	Gui, 11:Show, , Add a Smart Flag...
	return

}

^!m::			;Paper Replacer
{
	StringReplace, clipboard, clipboard,`n,,All
	StringReplace, clipboard, clipboard,`r,,All
	InputBox, reText, Add A Smart Substitution..., What would you like to replace "%clipboard%" with?
	
	If ErrorLevel
		return
	
	StringReplace, reText, reText,`n,,All
	FileAppend, "%clipboard%"`,"%reText%"`n, %A_MyDocuments%\MySmartSubstitutions.csv
	LoadSmartSubstitutions()
	CurrentCodeRocketDisplayedCase := 0
	Gosub, BuildMainGui
	return
}

^!p::	;Paper Replacer
{
printText=
(
{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Calibri;}}
{\*\generator Msftedit 5.41.21.2510;}\viewkind4\uc1\pard\ri-720\sl240\slmult1\lang9\b\f0\fs40 %CurrentCaseNumber%\tab\tab\tab\b0 %PatientName%   \fs22 %PatientAge% years old - DOB: %PatientDOB%\par
\pard\sl240\slmult1\fs28 %ClientName% \par
%ClientOfficeName% --- %ClientState%\par
\fs22\par
\b PREFERENCES\par
\b0  %preferences%\par
\par
\b CLINICAL INDICATIONS/HISTORY\par
\b0 %ClinicalData%\par
\par
\b FINAL DIAGNOSIS\par
\b0 %finaldiagnosistext%\par
\par
\b GROSS DESCRIPTION\par
\b0 %grossdescriptiontext%\par
\par
\b PRIOR CASE INFORMATION\par
\b0 %priorCaseInfo% \par
\par
\par
}
 )
	IfNotExist, %A_MyDocuments%\CaseFile\
		FileCreateDir, %A_MyDocuments%\CaseFile\
	
	FileDelete, %A_MyDocuments%\CaseFile\*.rtf
	StringReplace, printText, printText,<br>,\par%A_Space%, All
	StringReplace, printText, printText,<u>,\ul%A_Space% , All
	StringReplace, printText, printText,</u>,\ulnone%A_Space% , All
	FileAppend, %printText%, %A_MyDocuments%\CaseFile\%CurrentCaseNumber%.rtf
	Run, Wordpad.exe %A_MyDocuments%\CaseFile\\%CurrentCaseNumber%.rtf
Return
}

^k::   ;Special Stain order   ;Automation
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

^!s::  ;Batch Signout		;Automation
{
	ifWinExist, WinSURGE - Final Diagnosis
	{
		SoundBeep
		return
	}
	
	SetTimer, WinSURGECaseDataUpdater, Off

	Progress, x10 y10 h150, Preparing to signout, Obtaining Routine Cases for signout`n Press Ctrl-Alt-R to stop the signout, Working....,

	s =	select s.number, s.zaudittraillast, s.yesno07 from specimen s, physician p where  s.sodate < '1950-01-01' and s.path = p.id and p.name ='%WinSurgeFullName%' and s.calculatedslidecountdate>'1950-01-01' order by 2 desc
	WinSurgeQuery(s)
	FileAppend, %msg%, %A_MyDocuments%\SatSignout.txt
	
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
		If(A_Index=1)
			CurrentCodeRocketDisplayedCase := 0
		;Progress, , Preparing to signout, Signing out the cases`n, Working....,

		Loop, parse, SignOutFileList, `n
		{
			If(A_Index=1)
				CurrentCodeRocketDisplayedCase := 0     ;This fixes when the first case of signout doesn't update in the paper replacer.

			
			caseToSignout := A_LoopField
			y := SignOutCount - A_Index + 1
			x := 100 * (A_Index / SignOutCount)
			;Progress, X100 Y100 %x%,  %y% of %SignOutCount% cases remaining...

			if caseToSignout =  ; Omit the last linefeed (blank item) at the end of the list.
				break
			CloseWinSURGEModalWindow("WinSURGE Case Lookup","","Cancel")
			Open4SignOut(caseToSignout)
			
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
			
				z1 := caseToSignout
				Gosub, BuildMainGui
				

				
		if(A_PriorKey="F3" and A_TimeIdleKeyboard<1500)
			{
				EnterSignoutPasswordandApprove()	
			}
		else
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
	}

	Progress, Off   
	SetTimer, WinSURGECaseDataUpdater, 2000
	Return
}

^!e::	;Automation
{
	Gui, 4:Destroy
	
	Gui, 4:Font, S12, Verdana
	Gui, 4:Add, Text, vDisplayCaseNumber , Case Number: XXXXXXXXXXX
	Gui, 4:Add, DropDownList, vLoc w500, Boston|Irving
	Gui, 4:Add, DropDownList, AltSubmit vEmailType w500, Patient Double Blind Error|Need Previous Biopsy Report|Clinical Note and Photos|Critical Result Call|Pull Bottles and Blocks
	GuiControl, 4:Choose, EmailType, 1
	Gui, 4:Add, Text, ,Comments
	Gui, 4:Add, Edit, vEmailComments w500, 
	Gui, 4:Add, Button, Default gSendEmail vSendEmail, Send Email

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

^!u::	;Melanoma case check		;Automation
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

^!c::	;Automation
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

^!x::
{
perc := "%"
Gui, 11:Destroy
	
CPhtml=
(
This Client's Preferences<br>%preferences%<br><br>
<table border="1" style="border-collapse: collapse; width: 100%perc%;">
<tbody>
<tr>
	<td style="width: 49.1361%perc%;">
		<p>
			<span style="text-decoration: underline;"><strong>Margins</strong></span>
		</p>
		<p>
			<input id="cbB9NeviMargins" type="checkbox" /> B9 Nevi<br />
			<input id="cbMildDNMargins" type="checkbox" /> Mild DN<br />
			<input id="cbModDNMargins" type="checkbox" /> Mod DN<br />
			<input id="cbSevDNMargins" type="checkbox" /> Sev DN/Sev Atyp<br />
			<input id="cbMelanomaMargins" type="checkbox" /> Melanoma
		</p>
		<p>
			<input id="cbNMSCMargins" type="checkbox" /> NMSC<br /> 
			<input id="cbAKMargins" type="checkbox" /> AK/HAK
		</p>
		<p>
			<input id="cbExcisionsMargins" type="checkbox" /> Excisions<br /> 
			<input id="cbAllNeoplasmMargins" type="checkbox" /> All Neoplasms
		</p>
	</td>
	<td style="width: 50.8639%perc%;">
		<p>
			<span style="text-decoration: underline;"><strong>Recommendations</strong></span>
		</p>
		<p>
			<input id="cbB9NeviRecomendations" type="checkbox" /> B9 Nevi<br />
			<input id="cbMildDNRecomendations" type="checkbox" /> Mild DN<br />
			<input id="cbModDNRecomendations" type="checkbox" /> Mod DN<br />
			<input id="cbSevDNRecomendations" type="checkbox" /> Sev DN/Sev Atyp<br />
			<input id="cbMelanomaRecomendations" type="checkbox" /> Melanoma
		</p>
		<p>
			<input id="cbNMSCRecomendations" type="checkbox" /> NMSC<br /> 
			<input id="cbAKRecomendations" type="checkbox" /> AK/HAK
		</p>
		<p>
			<input id="cbExcisionsRecomendations" type="checkbox" /> Excisions<br /> 
			<input id="cbAllNeoplasmRecomendations" type="checkbox" /> All Neoplasms
		</p>
	</td>
</tr>
</tbody>
</table>
<p></p>
<p>
	<input id="cbUseMicros" type="checkbox" /> <strong><span style="text-decoration: underline;">Micros</span></strong><br />
	<input id="cbUsePhotos" type="checkbox" /> <strong><span style="text-decoration: underline;">Photos</span></strong>
</p>
<p>Do Not Use: 
	<input id="editDoNotUseList" type="text" /></p>
<p>
	ALL CAPS=DX Codes, all lower case=extended phrases, "literal phrases"
</p>
)
	
	Gui, 11:Add, ActiveX, w800 h590 vCP hwndATLWinHWND, Shell.Explorer
	Gui, 11:Add, Button,  Default, Save
	Gui, 11:Add, Button, , Cancel
	
	Gui, 11:Show, , Edit Doctor Preferences
	CP.Navigate("about:blank")
	CP.document.write(CPhtml)


	return
}

^!y::   ;Hotkey to find out who signs out what % of a client's cases
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

^!z::   ;Test Hotkey for debugging
{
	Msgbox, %A_AhkVersion%
return
}

^!v::	;Debugging
{
	ListVars
	Return
}

^!l::  ;Debugging
{
	ListLines
	Return
}

^!1::	;coding
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
			;SplashTextOn, 100, 100, EP convertor, Doing %DXCodeCount%
			FileAppend, ::%DxCode2%::%DxCode3%`n, S:\CodeRocket\bin\EP\DermpathExtendedPhrases.ahk
			rs.MoveNext()
	}
	rs.close()   
	adodb.close()
	Msgbox, Done!
	Return
}

^!2::Msgbox, %marginFlags%
	
#h::  ; Win+H hotkey used to add hotstrings to the personal extended phrases 
{
; Get the text currently selected. The clipboard is used instead of
; "ControlGet Selected" because it works in a greater variety of editors
; (namely word processors).  Save the current clipboard contents to be
; restored later. Although this handles only plain text, it seems better
; than nothing:
AutoTrim Off  ; Retain any leading and trailing whitespace on the clipboard.
ClipboardOld = %ClipboardAll%
Clipboard =  ; Must start off blank for detection to work.
Send ^c
ClipWait 1
if ErrorLevel  ; ClipWait timed out.
    return
; Replace CRLF and/or LF with `n for use in a "send-raw" hotstring:
; The same is done for any other characters that might otherwise
; be a problem in raw mode:
StringReplace, Hotstring, Clipboard, ``, ````, All  ; Do this replacement first to avoid interfering with the others below.
StringReplace, Hotstring, Hotstring, `r`n, ``r, All  ; Using `r works better than `n in MS Word, etc.
StringReplace, Hotstring, Hotstring, `n, ``r, All
StringReplace, Hotstring, Hotstring, %A_Tab%, ``t, All
StringReplace, Hotstring, Hotstring, `;, ```;, All
Clipboard = %ClipboardOld%  ; Restore previous contents of clipboard.

; Show the InputBox, providing the default hotstring:
InputBox, codeHotstring, New Extended Phrase, Type your abreviation for this phrase '%Hotstring%':
if (ErrorLevel OR codeHotstring="")  
    return

; Otherwise, add the hotstring and reload the personal Extended phrases:
FileAppend, `n::%codeHotstring%::%Hotstring%`n, %A_MyDocuments%\PersonalExtendedPhrases.ahk ; Put a `n at the beginning in case file lacks a blank line at its end.
LoadPersonalExtendedPhrases()
If ErrorLevel
	MsgBox, 4,, The hotstring just added appears to be improperly formatted.  Would you like to open the script for editing? Note that the bad hotstring is at the bottom of the script.
return
}

Pause::Pause
^!r::
{
	Run, C:\Documents and Settings\All Users\Desktop\Launcher - Caris CodeRocket.exe
	ExitApp
	Return
}

^!q::Reload		;Debugging
^+!r::			
{
	FileDelete, %A_MyDocuments%\CarisCodeRocket.ini
	Run, C:\Documents and Settings\All Users\Desktop\Launcher - Caris CodeRocket.exe
	ExitApp
	Return
}

#IfWinActive, WinSURGE - 		;Coding
Shift & Enter::
{
		If (BeepOnShiftEnter)
			SoundBeep
		
		Send, {#}{#}
		Sleep, 200
		ControlGetText, finaldxcontents, TX202, WinSURGE - 	
		Send, {Backspace}{Backspace}
		
		;rawCode0 holds the entire match, rawCode2 holds the front helpers, 3 holds the DX code, 5 holds the margins, 7 holds the comment helpers, 8 holds the "?/" modifiers
		codePresent := RegExMatch(finaldxcontents,"(?m)[\s|\n]((\w*):)?(\w+)(\.(\w*))?(;(\w*))?([\?\/]{0,2}?)\#\#", rawCode)  
		
		dxCode := rawCode3
		frontHelpers := rawCode2
		marginHelper := rawCode5
		commentHelpers := rawCode7
		modifiers := rawCode8
		
		If (dxCode="" OR StrLen(marginHelper)>1)
		{			
		Msgbox, There is an error in your diagnosis code!  Please reenter!
			return
		}
		;DetermineWhichJar
		foundPos := RegExMatch(finaldxcontents,"(?m)(\w)\.[^\r]*%%P%%\r[^\r]*\#\#",jar)   ;jar1 will contain the letter of the jar for the entered code
		;Get the Clinical History for that Jar

		nextJar := Chr(Asc(jar1) + 1)
		needleRegex=%jar1%\.([\w\s\W\n\r]+)(%nextJar%\.|$)?

		foundPos := RegExMatch(ClinicalData, needleRegex, thisJarClinicalData)  ;This jars clinical data will be contained in thisJarClinicalData1
		
		If !thisJarClinicalData1
			thisJarClinicalData1 := ClinicalData
		
		StringReplace, thisJarClinicalData1, thisJarClinicalData1, <br>,@
		StringSplit, tempVar, thisJarClinicaldata1, @
		thisJarClinicalData1 := tempVar1
		
		;Msgbox, %marginFlags%
		;Msgbox, %jar1%`n%ClinicalData%`n%thisJarClinicalData1%
		
			TempMicros = 0	
			SuppressMicros = 0	
			SelectedCodeIndex  = 0
			
		If (Instr(modifiers, "?"))
			SuppressMicros := 1
		
		If (Instr(modifiers, "/"))
			TempMicros := 1
		
		If (namedDxCodes[dxCode].dxLine = "")
			{
			Msgbox, That is not a valid diagnosis code!
			return
			}

		SelDXCode := namedDxCodes[dxCode]["code"]
		SelDiagnosis := namedDxCodes[dxCode]["dxLine"]
		SelComment := namedDxCodes[dxCode]["comment"]
		SelMicro := namedDxCodes[dxCode]["Micro"]
		SelCPTCode := namedDxCodes[dxCode]["cptCode"]
		SelICD9 := namedDxCodes[dxCode]["ICD9"]
		SelICD10 := namedDxCodes[dxCode]["ICD10"]
		SelSnomed := namedDxCodes[dxCode]["SNOMED"]
		SelPre := namedDxCodes[dxCode]["premalignantFlag"]
		SelMal := namedDxCodes[dxCode]["malignantFlag"]
		SelDys := namedDxCodes[dxCode]["dysplasticFlag"]			
		SelInf := namedDxCodes[dxCode]["inflammatoryFlag"]
		SelMargInc := namedDxCodes[dxCode]["marginIncludedFlag"]
		SelLog := namedDxCodes[dxCode]["log"]
		
		ErrorLevel := MildDysplasticWarningCheck()
		If ErrorLevel
			return

				;Loop to add front helper codes to the diagnosis
				Loop, % StrLen(frontHelpers)
				{
					tempVar1 := Substr(frontHelpers, StrLen(frontHelpers)-A_Index+1, 1)
					tempVar2 := FrontofDiagnosisHelper%tempVar1%
					SelDiagnosis = %tempVar2% %SelDiagnosis%
				}

				;If there is one, add a margin code to the diagnosis
				Stringlen, i, marginHelper
				if (i=0)
				{
									
					tempVar := Instr(thisJarClinicalData1, "margin") AND !InStr(SelDiagnosis, "No Residual")
					tempVar += Instr(marginFlags, "margins-all-nevi") AND (Instr(SelDiagnosis, "nevus") OR Instr(SelDiagnosis, "melanocy"))  ;For All nevi
					if (tempVar)
					{
						SoundBeep 
						SoundBeep
						Msgbox, 4,Margin Warning!!!, The client has requested a margin and you did not use one.  Do you want to continue without a margin?
						IfMsgbox, No
							Return	
					}

 				}
				else if i=1
					{
					
					tempVar := Instr(marginFlags, "none")  ;Warn if they have requested no margins
					tempVar += Instr(marginFlags, "no-margins-nmsc") AND (Instr(SelDiagnosis, "carcinoma") OR Instr(SelDiagnosis,"actinic keratosis")) ;warn if they have requested no margins on NMSC
					if (tempVar)
						{
							SoundBeep
							SoundBeep
							Msgbox,4,Margin Warning!!!,This client has requested the following margin preferences (%UseMargins%) and you have used one!  Do you wish to continue?
							IfMsgbox, No
								Return	
						}

						p := BackofDiagnosisHelper%marginHelper%
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
				Loop, % StrLen(commentHelpers)
				{
					tempVar1 := Substr(commentHelpers, A_Index, 1)
					tempVar2 := CommentHelper%tempVar1%
					SelComment = %SelComment%  %tempVar2%  
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
				
				If (((TempMicros OR UseMicros) AND !SuppressMicros) AND !x5)
					Msgbox, Client has requested microscopic descriptions and there is not one for this diagnostic code!  Please enter manually.
				If (x4 or (x5 and ((TempMicros OR UseMicros) AND !SuppressMicros)))
					{
					If ((TempMicros OR UseMicros) AND !SuppressMicros)
						dxtext = %dxtext%`n`nComment:%A_Space%%x4%%A_Space%%A_Space%%x5%
					Else
						dxtext = %dxtext%`n`nComment:%A_Space%%x4%
					}

				dxtext = %dxtext%`n
				If x6
					{
					Loop, parse, x6,`;
						dxtext = %dxtext%%perc%%A_LoopField%%perc%	
					}
				dxtext = %dxtext%%perc%%SelDXCode%%perc%
					
				SetCapsLockState, Off
		
		rawCode := trim(rawCode)
		StringReplace, rawCode, rawCode, `n,, All
		StringReplace, newText, finaldxcontents, %rawCode%, %dxtext%		
		ControlSetText, TX202, %newtext%, WinSURGE -  
		LastCodeUsed := SubStr(rawCode, 1, StrLen(rawCode)-2)
		DataEntered = 1
			

			ifWinActive,  WinSURGE - Final Diagnosis:
			{
				Gui, Submit, NoHide	
				GuiControl, Enable, CaseScanBox
				GuiControl, Enable, CaseLoaderLbl
				if CaseScanBox=Data in Case
					GuiControl, Text, CaseScanBox, 
			
			finaldiag := WinSURGEFinalDiagnosisContents()
			StringGetPos, i, finaldiag, ***
			if i>0
				ActivateNextTripleAsterisk()
			Else 
				Gosub, F12
			}

Return
}

#IfWinActive, Special Stains Checklist		;Automation
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
	IfWinExist, %WinTitle%, %WinText%
	{
		Loop,
		{
			FirstTimeBeep := 0
			WinClose, %WinTitle%, %WinText%
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
			
				SetControlDelay, 100
				len := StrLen(WinSurgeFullName)
				len := len -1
				StringLeft, PartialName, WinSurgeFullName, %len%
				ControlSetText, %PathologistTextBox%, %PartialName%, WinSURGE [
				ControlClick, %PathologistTextBox%, WinSURGE [
				ControlSend, %PathologistTextBox%, {TAB}, WinSURGE [

				Loop,  
				{
					ControlGetText, t1, %PathologistTextBox%, WinSURGE [
					if (t1=WinSurgeFullName)
						Break
					else
						Sleep, 100
				}

				ControlSetText, %QueueIntoBatchBox%, Final rep, WinSURGE [
				ControlClick, %QueueIntoBatchBox%, WinSURGE [
				ControlSend, %QueueIntoBatchBox%, {TAB}, WinSURGE [

				Loop,  
				{
					ControlGetText, t2, %QueueIntoBatchBox%, WinSURGE [
					if (t2="Final Reports")
						Break
					else
						Sleep, 100
				}

return
}

SkipCaseSignout()
{
	global
	WinWaitClose, WinSURGE E-signout, abcdefgABCDEFG 1234567890, , [
	WinActivate, WinSURGE E-signout [, &Approve
	WinWaitActive, WinSURGE E-signout [, &Approve
	LabeledButtonPress("WinSURGE E-signout [","&Approve","&Skip")
	
	SetTitleMatchMode, 2
	WinWait, -- No Current Case, , 5
	
	WinWait, WinSURGE E-signout, , 5, -- No Current Case
	SetTitleMatchMode, 1
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
						
				ifWinNotExist, WinSURGE E-signout,abcdefgABCDEFG 1234567890
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
				return
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

LoadPersonalExtendedPhrases()
{
	global
	
		IfExist, %A_MyDocuments%\PersonalExtendedPhrases.ahk
		{
		Run, S:\CodeRocket\bin\EP\Autohotkey.exe "%A_MyDocuments%\PersonalExtendedPhrases.ahk", ,UseErrorLevel , ppid	
		If ErrorLevel
			Msgbox, There was an error loading your personal extended phrases file. %ErrorLevel%
		}

	return
}

ReadDXCodes()
{
	global
	IfExist, %A_MyDocuments%\dxcodes.csv
	{
		Loop, read, %A_MyDocuments%\dxcodes.csv
		{
			Loop, parse, A_LoopReadLine, CSV
					tempVar%A_Index%=%A_LoopField%
			
			namedDXCodes[tempVar2] := {}
			namedDXCodes[tempVar2]["code"] := tempVar2		
			namedDXCodes[tempVar2]["category"] := tempVar3
			namedDXCodes[tempVar2]["subcategory"] := tempVar4
			namedDXCodes[tempVar2]["dxLine"] := tempVar5
			namedDXCodes[tempVar2]["comment"] := tempVar6
			namedDXCodes[tempVar2]["micro"] := tempVar7
			namedDXCodes[tempVar2]["cptCode"] := tempVar8
			namedDXCodes[tempVar2]["ICD9"] := tempVar9
			namedDXCodes[tempVar2]["ICD10"] := tempVar10
			namedDXCodes[tempVar2]["SNOMED"] := tempVar11
			namedDXCodes[tempVar2]["premalignantFlag"] := tempVar12
			namedDXCodes[tempVar2]["malignantFlag"] := tempVar13
			namedDXCodes[tempVar2]["dysplasticFlat"] := tempVar14
			namedDXCodes[tempVar2]["melanocyticFlag"] := tempVar15
			namedDXCodes[tempVar2]["inflammatoryFlag"] := tempVar16
			namedDXCodes[tempVar2]["marginIncludedFlag"] := tempVar17
			namedDXCodes[tempVar2]["log"] := tempVar18			
		}
		
		;Msgbox, % namedDXCodes["bccn"]["dxLine"]
}
	else
	{
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
			tempVar%j%=%y%
			}
		
			if !namedDXCodes[tempVar2]["code"]
			{
				namedDXCodes[tempVar2] := {}
				namedDXCodes[tempVar2]["code"] := tempVar2		
				namedDXCodes[tempVar2]["category"] := tempVar3
				namedDXCodes[tempVar2]["subcategory"] := tempVar4
				namedDXCodes[tempVar2]["dxLine"] := tempVar5
				namedDXCodes[tempVar2]["comment"] := tempVar6
				namedDXCodes[tempVar2]["micro"] := tempVar7
				namedDXCodes[tempVar2]["cptCode"] := tempVar8
				namedDXCodes[tempVar2]["ICD9"] := tempVar9
				namedDXCodes[tempVar2]["ICD10"] := tempVar10
				namedDXCodes[tempVar2]["SNOMED"] := tempVar11
				namedDXCodes[tempVar2]["premalignantFlag"] := tempVar12
				namedDXCodes[tempVar2]["malignantFlag"] := tempVar13
				namedDXCodes[tempVar2]["dysplasticFlat"] := tempVar14
				namedDXCodes[tempVar2]["melanocyticFlag"] := tempVar15
				namedDXCodes[tempVar2]["inflammatoryFlag"] := tempVar16
				namedDXCodes[tempVar2]["marginIncludedFlag"] := tempVar17
				namedDXCodes[tempVar2]["log"] := tempVar18			
			}
			rs.MoveNext()
}

	rs.close()   
	adodb.close()
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
			Msgbox, WinSURGE must be running the first time you fire the Caris CodeRocket.  Please login to WinSURGE and restart the CodeRocket.
			ExitApp
		}
	
	startSetup:
	InputBox, WinSURGEFullName, CodeRocket Setup, Enter your full name exactly (including spaces) as it appears in the 'Pathologist' box in WinSurge (usually, last comma space first):
	s := "select p.name, p.id from physician p where p.name='" . WinSURGEFullName . "'"
	WinSurgeQuery(s)
	If RegExMatch(Result_2, "(\d+)\.", tempVar)
		WinSURGEPathologistID := tempVar1
	else
	{
		Msgbox, The name you entered does not match any WinSurge names EXACTLY.  Please enter your full Winsurge name as it appears in the pathologist box including any spaces and punctuation.
		Gosub, startSetup
	}

	InputBox, pathologistLocation, CodeRocket Setup, Enter the location where your computer is located (Irving`, Boston`, or Union only)  

	CarisRocketWindowX := 0
	CarisRocketWindowY := 0
	CarisRocketWindowW := 0
	CarisRocketWindowH := 0

	IniWrite, %pathologistLocation%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, pathologistLocation
	IniWrite, %WinSurgeFullName%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSurgeFullName
	IniWrite, %WinSURGEPathologistID%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSURGEPathologistID
	IniWrite, %CarisRocketWindowX%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowX
	IniWrite, %CarisRocketWindowY%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowY
	IniWrite, %CarisRocketWindowW%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowW
	IniWrite, %CarisRocketWindowH%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowH

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

	IniRead, WinSurgeSignoutPassword, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSurgeSignoutPassword
	IniRead, WinSurgeFullName, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSurgeFullName
	IniRead, WinSURGEPathologistID, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, WinSURGEPathologistID
	IniRead, pathologistLocation, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, pathologistLocation

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

	IniRead, SpeakEnabled, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, SpeakEnabled
	IniRead, BeepOnShiftEnter, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, BeepOnShiftEnter
	IniRead, UseSmartSubstitutions, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, UseSmartSubstitutions
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
 
LoadSmartSubstitutions()  ;PaperReplacer
{
	global
	smartSubs := []
	
	IfExist, %A_MyDocuments%\MySmartSubstitutions.csv
	{
		FileRead, fileContents, %A_MyDocuments%\MySmartSubstitutions.csv

		Loop, read, %A_MyDocuments%\MySmartSubstitutions.csv
		{

			Loop, parse, A_LoopReadLine, CSV
				sSub%A_Index%=%A_LoopField%
			
				;Msgbox, %sSub1%,%sSub2%,%sSub3%
			smartSubs[A_Index] := { "searchText":sSub1, "replacementText":sSub2, "flag":sSub3}
			
		}
	}

	return
}

WM_KEYDOWN(wParam, lParam, nMsg, hWnd)  ;Paper Replacer
{
   global wb
   static fields := "hWnd,nMsg,wParam,lParam,A_EventInfo,A_GuiX,A_GuiY"
   WinGetClass, ClassName, ahk_id %hWnd%
   if  (ClassName = "Internet Explorer_Server")
   {
   ;// Get the in place interface pointer
      pipa := ComObjQuery(wb.document, "{00000117-0000-0000-C000-000000000046}")
   ;// Build MSG Structure
      VarSetCapacity(Msg, 48)
      Loop Parse, fields, `,             ;`
         NumPut(%A_LoopField%, Msg, (A_Index-1)*A_PtrSize)
   ;// Call Translate Accelerator Method
      TranslateAccelerator := NumGet(NumGet(1*pipa)+5*A_PtrSize)
      Loop 2 ;// only necessary for Shell.Explorer Object
         r := DllCall(TranslateAccelerator, "Ptr",pipa, "Ptr",&Msg)
      until wParam != 9 || wb.document.activeElement != ""
   ;// Release the in place interface pointer
      ObjRelease(pipa)
		
      if r = 0 ;// S_OK: the message was translated to an accelerator.
         return 0
   }
}

ReDrawGui()  ;Paper replacer
{
	global

	if(z1="No Current Case" OR z1="")
	{
		GuiControl, Enable, CaseScanBox 
		GuiControl, , CaseNumberLabel, No Current Case Open
		;GuiControl, Enable, OK
		GuiControl, Hide, PatientLabel
		GuiControl, Hide, DoctorLabel
		GuiControl, Hide, ClientLabel
		GuiControl, Hide, UsePhotos
		GuiControl, Hide, UseMicros
		GuiControl, Hide, UseMargins
		GuiControl, Hide, WB
		Menu, EditMenu, Disable, Edit Client Preferences
		displayedPreferences := ""
		displayedClinicalData := ""
		displayedFinalDiagnosis := ""
		displayedGrossDescription := ""


	}
	else if (z1="WinSURGE Not Open")
	{
		
		GuiControl, Disable, CaseScanBox
		GuiControl, Disable, CaseLoaderLbl
		GuiControl, , CaseNumberLabel, WinSURGE is not open!
		;GuiControl, Disable, OK
		GuiControl, Hide, PatientLabel
		GuiControl, Hide, DoctorLabel
		GuiControl, Hide, ClientLabel
		GuiControl, Hide, UsePhotos
		GuiControl, Hide, UseMicros
		GuiControl, Hide, UseMargins
		GuiControl, Hide, WB
		Menu, EditMenu, Disable, Edit Client Preferences
		displayedPreferences := ""
		displayedClinicalData := ""
		displayedFinalDiagnosis := ""
		displayedGrossDescription := ""
	}
	else
	{
		GuiControl, Enable, CaseScanBox 
		GuiControl, , CaseNumberLabel, %CurrentCaseNumber%
		;GuiControl, Enable, OK
		GuiControl, Show, PatientLabel
		GuiControl, , PatientLabel, Patient: %PatientName% (%PatientAge%) - DOB: %PatientDOB%
		GuiControl, Show, DoctorLabel
		GuiControl, , DoctorLabel, Doctor: %ClientName%
		GuiControl, Show, ClientLabel
		GuiControl, , ClientLabel, Client: %ClientOfficeName%  --- %ClientState%
		GuiControl, Show, UsePhotos
		GuiControl, Show, UseMicros
		GuiControl, Show, UseMargins
		GuiControl, Show, WB
		Menu, EditMenu, Enable, Edit Client Preferences
		
		;SMART Replacements start here
		if UseSmartSubstitutions
		{
			marginFlags := ""
			Loop, % smartSubs.Length()+1
			{
				m := smartSubs[A_Index].searchText
				n := smartSubs[A_Index].replacementText
				o := smartSubs[A_Index].flag
					
								
				if (Instr(displayedPreferences, m) AND !Instr(displayedPreferences, n))
					{
						StringReplace, displayedPreferences, displayedPreferences, %m%, %n%, All
						marginFlags=%marginFlags%, %o%
					}

				StringReplace, displayedClinicalData, displayedClinicalData, %m%, %n%, All
				StringReplace, displayedGrossDescription, displayedGrossDescription, %m%, %n%, All
			}
		}

		
		Loop,
		{
			StringReplace, displayedPreferences, displayedPreferences, <br><br>, <br>, All
			IfNotInString, displayedPreferences, <br><br>
				break
		}
		
		tempVar := displayedPreferences
		StringReplace, tempVar, tempVar, <br>,¥ , All
		
		tempVar3 := ""
		Loop, Parse, tempVar, ¥
			{
				tempVar2 := Trim(A_LoopField)
				if(tempVar2<>"")
					tempVar3 := tempVar3 . tempVar2 . "<br>"
			}
			
		displayedPreferences := tempVar3	
		
		if (attnPathologistField)
			WB.document.getElementById("attnPathologist").innerHTML := "<h4> ATTENTION PATHOLOGIST:  " . attnPathologistField . "</h4>"
		else
			WB.document.getElementById("attnPathologist").innerHTML := ""

		if (orderedProcedures)
			WB.document.getElementById("orderedProcedures").innerHTML := "<h4>Ordered Procedures:  " . orderedProcedures . "</h4>"
		else
			WB.document.getElementById("orderedProcedures").innerHTML := ""

		WB.document.getElementById("preferences").innerHTML := displayedPreferences
		WB.document.getElementById("clinicalInformation").innerHTML := displayedClinicalData
		WB.document.getElementById("finalDiagnosis").innerHTML := displayedFinalDiagnosis
		WB.document.getElementById("grossDescription").innerHTML := displayedGrossDescription
		WB.document.getElementById("priorCaseInformation").innerHTML := priorCaseInfo
	}
	
return
}

Json_Load(ByRef src, args*)
{
	static q := Chr(34)

	key := "", is_key := false
	stack := [ tree := [] ]
	is_arr := { (tree): 1 }
	next := q . "{[01234567890-tfn"
	pos := 0
	while ( (ch := SubStr(src, ++pos, 1)) != "" )
	{
		if InStr(" `t`n`r", ch)
			continue
		if !InStr(next, ch, true)
		{
			ln := ObjLength(StrSplit(SubStr(src, 1, pos), "`n"))
			col := pos - InStr(src, "`n",, -(StrLen(src)-pos+1))

			msg := Format("{}: line {} col {} (char {})"
			,   (next == "")      ? ["Extra data", ch := SubStr(src, pos)][1]
			  : (next == "'")     ? "Unterminated string starting at"
			  : (next == "\")     ? "Invalid \escape"
			  : (next == ":")     ? "Expecting ':' delimiter"
			  : (next == q)       ? "Expecting object key enclosed in double quotes"
			  : (next == q . "}") ? "Expecting object key enclosed in double quotes or object closing '}'"
			  : (next == ",}")    ? "Expecting ',' delimiter or object closing '}'"
			  : (next == ",]")    ? "Expecting ',' delimiter or array closing ']'"
			  : [ "Expecting JSON value(string, number, [true, false, null], object or array)"
			    , ch := SubStr(src, pos, (SubStr(src, pos)~="[\]\},\s]|$")-1) ][1]
			, ln, col, pos)

			throw Exception(msg, -1, ch)
		}

		is_array := is_arr[obj := stack[1]]

		if i := InStr("{[", ch)
		{
			val := (proto := args[i]) ? new proto : {}
			is_array? ObjPush(obj, val) : obj[key] := val
			ObjInsertAt(stack, 1, val)
			
			is_arr[val] := !(is_key := ch == "{")
			next := q . (is_key ? "}" : "{[]0123456789-tfn")
		}

		else if InStr("}]", ch)
		{
			ObjRemoveAt(stack, 1)
			next := stack[1]==tree ? "" : is_arr[stack[1]] ? ",]" : ",}"
		}

		else if InStr(",:", ch)
		{
			is_key := (!is_array && ch == ",")
			next := is_key ? q : q . "{[0123456789-tfn"
		}

		else ; string | number | true | false | null
		{
			if (ch == q) ; string
			{
				i := pos
				while i := InStr(src, q,, i+1)
				{
					val := StrReplace(SubStr(src, pos+1, i-pos-1), "\\", "\u005C")
					static end := A_AhkVersion<"2" ? 0 : -1
					if (SubStr(val, end) != "\")
						break
				}
				if !i ? (pos--, next := "'") : 0
					continue

				pos := i ; update pos

				  val := StrReplace(val,    "\/",  "/")
				, val := StrReplace(val, "\" . q,    q)
				, val := StrReplace(val,    "\b", "`b")
				, val := StrReplace(val,    "\f", "`f")
				, val := StrReplace(val,    "\n", "`n")
				, val := StrReplace(val,    "\r", "`r")
				, val := StrReplace(val,    "\t", "`t")

				i := 0
				while i := InStr(val, "\",, i+1)
				{
					if (SubStr(val, i+1, 1) != "u") ? (pos -= StrLen(SubStr(val, i)), next := "\") : 0
						continue 2

					; \uXXXX - JSON unicode escape sequence
					xxxx := Abs("0x" . SubStr(val, i+2, 4))
					if (A_IsUnicode || xxxx < 0x100)
						val := SubStr(val, 1, i-1) . Chr(xxxx) . SubStr(val, i+6)
				}

				if is_key
				{
					key := val, next := ":"
					continue
				}
			}

			else ; number | true | false | null
			{
				val := SubStr(src, pos, i := RegExMatch(src, "[\]\},\s]|$",, pos)-pos)
			
			; For numerical values, numerify integers and keep floats as is.
			; I'm not yet sure if I should numerify floats in v2.0-a ...
				static number := "number", integer := "integer"
				if val is %number%
				{
					if val is %integer%
						val += 0
				}
			; in v1.1, true,false,A_PtrSize,A_IsUnicode,A_Index,A_EventInfo,
			; SOMETIMES return strings due to certain optimizations. Since it
			; is just 'SOMETIMES', numerify to be consistent w/ v2.0-a
				else if (val == "true" || val == "false")
					val := %value% + 0
			; AHK_H has built-in null, can't do 'val := %value%' where value == "null"
			; as it would raise an exception in AHK_H(overriding built-in var)
				else if (val == "null")
					val := ""
			; any other values are invalid, continue to trigger error
				else if (pos--, next := "#")
					continue
				
				pos += i-1
			}
			
			is_array? ObjPush(obj, val) : obj[key] := val
			next := obj==tree ? "" : is_array ? ",]" : ",}"
		}
	}

	return tree[1]
}

Json_Dump(obj, indent:="", lvl:=1)
{
	static q := Chr(34)

	if IsObject(obj)
	{
		static Type := Func("Type")
		if Type ? (Type.Call(obj) != "Object") : (ObjGetCapacity(obj) == "")
			throw Exception("Object type not supported.", -1, Format("<Object at 0x{:p}>", &obj))

		is_array := 0
		for k in obj
			is_array := k == A_Index
		until !is_array

		static integer := "integer"
		if indent is %integer%
		{
			if (indent < 0)
				throw Exception("Indent parameter must be a postive integer.", -1, indent)
			spaces := indent, indent := ""
			Loop % spaces
				indent .= " "
		}
		indt := ""
		Loop, % indent ? lvl : 0
			indt .= indent

		lvl += 1, out := "" ; Make #Warn happy
		for k, v in obj
		{
			if IsObject(k) || (k == "")
				throw Exception("Invalid object key.", -1, k ? Format("<Object at 0x{:p}>", &obj) : "<blank>")
			
			if !is_array
				out .= ( ObjGetCapacity([k], 1) ? Json_Dump(k) : q . k . q ) ;// key
				    .  ( indent ? ": " : ":" ) ; token + padding
			out .= Json_Dump(v, indent, lvl) ; value
			    .  ( indent ? ",`n" . indt : "," ) ; token + indent
		}

		if (out != "")
		{
			out := Trim(out, ",`n" . indent)
			if (indent != "")
				out := "`n" . indt . out . "`n" . SubStr(indt, StrLen(indent)+1)
		}
		
		return is_array ? "[" . out . "]" : "{" . out . "}"
	}

	; Number
	else if (ObjGetCapacity([obj], 1) == "")
		return obj

	; String (null -> not supported by AHK)
	if (obj != "")
	{
		  obj := StrReplace(obj,  "\",    "\\")
		, obj := StrReplace(obj,  "/",    "\/")
		, obj := StrReplace(obj,    q, "\" . q)
		, obj := StrReplace(obj, "`b",    "\b")
		, obj := StrReplace(obj, "`f",    "\f")
		, obj := StrReplace(obj, "`n",    "\n")
		, obj := StrReplace(obj, "`r",    "\r")
		, obj := StrReplace(obj, "`t",    "\t")

		static needle := (A_AhkVersion<"2" ? "O)" : "") . "[^\x20-\x7e]"
		while RegExMatch(obj, needle, m)
			obj := StrReplace(obj, m[0], Format("\u{:04X}", Ord(m[0])))
	}
	
	return q . obj . q
}