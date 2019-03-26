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

	ReadIniValues()
	LoadSmartSubstitutions()

	html=<!DOCTYPE html><html><meta name="viewport" content="width=device-width, initial-scale=1.0"><body><span id='main' style='white-space:pre-line'><span id='caseNumber' style='color:blue;font-size:36px;'></span><br><span id='procedureNote' style='color:red'></span><span id='additionalClinicalInformation' style='color:red'></span><span id='orderedProcedures'></span><span id='attnPathologist' style='color:red'></span><strong>Preferences:</strong><br><span id='preferences' style='white-space:pre-line'></span><br><br><strong>Clinical Information:<br></strong><span id='clinicalInformation' style="color:blue"></span><br><br><strong>Final Diagnosis:<br></strong><span id='finalDiagnosis' style="color:green"></span><br><br><strong>Gross Description:<br></strong><span id='grossDescription'></span><br><br><strong>Prior Case Information<br></strong><span id='priorCaseInformation'></span></span></body></html>

	Gui, Add, ActiveX, w600 h750 vWB hwndATLWinHWND, Shell.Explorer
	ComObjConnect(WB, new Event)
	
	;These Lines are required for the copy function to work inside the browerWindow
	IOleInPlaceActiveObject_Interface:="{00000117-0000-0000-C000-000000000046}"
	pipa := ComObjQuery(WB, IOleInPlaceActiveObject_Interface)
	OnMessage(WM_KEYDOWN:=0x0100, "WM_KEYDOWN")
	OnMessage(WM_KEYUP:=0x0101, "WM_KEYDOWN")
	OnMessage(0x5000, Speak)
	
	WB.Navigate("about:blank")
	WB.document.write(html)
	
	Menu, FileMenu, Add, E&xit, GuiClose
	
	Menu, SettingsMenu, Add, Use Smart Substitutions, UseSmartSubstitutions

	If UseSmartSubstitutions
		Menu, SettingsMenu, Check, Use Smart Substitutions
	Else
		Menu, SettingsMenu, UnCheck, Use Smart Substitutions

	Menu, MyMenuBar, Add, &File, :FileMenu  
	Menu, MyMenuBar, Add, &Settings, :SettingsMenu
	Gui, Menu, MyMenuBar

	Gui, +Resize	
	Gui, Show, x%JustPaperWindowX% y%JustPaperWindowY% w%JustPaperWindowW% h%JustPaperWindowH%

	SetTimer, WinSURGECaseDataUpdater, 1000
	SetTimer, WinSURGEFinalDXDetector, 1000

	Gosub, WinSurgeCaseDataUpdater

	SetTitleMatchMode, 2
	WinMove, Caris CodeRocket v2.5a.exe, , , , 518, 445 
	WinMove, Just Paper.ahk, , , , 775, 1046, SciTE4

	return
}

BuildMainGui:  ;Paper Replacer
{
	finaldxcontents := ""
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
		
		s := "select s.dx, s.gross, s.numberofspecimenparts, s.custom03, s.clin, p.name, s.clindata, pt.name, s.Computed_PATIENTAGE, p.proficiencylog, p.comment, s.custom04, s.patient, s.zfield, s.Computed_PatientDOB, s.computed_procabs, z.proficiencylog, s.image14, s.image15 from specimen s, physician p, physician z, patient pt where s.patient = pt.id and s.clin=p.id and s.client=z.id and computed_numberfilled='" . x . "'"

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
		rawPreferences := Result_10
		ClientOfficeName := Result_11
		ClientID := Result_12
		StringReplace, ClientID, ClientID, `n,,All
		StringSplit, PatientID, Result_13, .    ; The patient ID is in a variable called PatientID1
		attnPathologistField := Result_14
		PatientDOB := Result_15
		orderedProcedures := Result_16
		clientPreferences := Result_17
		procedureNote := Result_18
		additionalClinicalInformation := Result_19
		
		rawPreferences=%rawPreferences%`n%clientPreferences%
		
		FileDelete, RawPreferences.txt
		FileAppend, %rawPreferences%, RawPreferenes.txt
		
		
		if (ClientWinSurgeID)
		{
		s = Select top 1 c.name,c.photo_pref,c.micro_pref,c.margin_pref, c.icd9_pref, c.log from clinipref c where c.WinSurge_id=%ClientWinSurgeId%
	CodeDatabaseQuery(s)
	;Msgbox, %msg%
		}
	
		 IfInString, rawPreferences, ICD10
			{
				SoundBeep
				SoundBeep
				SoundBeep
				SoundBeep
				;Msgbox, ICD10's are required!
			}

			
		
		;Mandatory Replacements for basic formatting start here
		attnPathologistField := RegExReplace(attnPathologistField, "[a-z]+ \d+\/\d+\/\d+ \d+:\d+ \w+", "")
		StringReplace, attnPathologistField, attnPathologistField, `r,,All
		StringReplace, attnPathologistField, attnPathologistField, `n,,All
		
		StringLeft, ClientState, ClientID, 2
		
		StringSplit, PhysicianWinSurgeId, ClientWinSurgeId, .  ;physican WinSurge Id is stored in PhysicianWinSurgeId1
		
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

		displayedClinicalData := ClinicalData
		displayedFinalDiagnosis := finaldiagnosistext
		displayedGrossDescription := grossdescriptiontext
		
		displayedClinicalData := RegExReplace(displayedClinicalData, "(([P|p]lease)?\s?(?i)((check)?\s?margins\W?))", "<strong style='color:Red'>$0</strong>")
		
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

			
		ReDrawGui() ;Uses rawPreferences
		

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
			
if (!DisableDermCoding)
{
					IfInString, alertFlags, use-photos
						{
						GuiControl, 1:Show, UsePhotos
						UsePhotos := 1
					}
					Else	
						{
						GuiControl, 1:Hide, UsePhotos
						UsePhotos := 0
					}

					IfInString, alertFlags, use-micros
						{
						GuiControl, 1:Show, UseMicros	
						UseMicros := 1
					}	
					Else	
					{
						GuiControl, 1:Hide, UseMicros
						UseMicros := 0
					}
	}
return
}

WinSURGEFinalDXDetector:
{
	IfWinExist, WinSURGE - Final Diagnosis:
		{
				ControlGetText, t, TX202, WinSURGE - Final Diagnosis:
				If (t<>finaldxContents)
					{
						finaldxContents := t
						Gosub, CheckForWarnings
					}
		}
	return
}

WinSURGECaseDataUpdater:  ;Automation
{

	;This section gets and saves the position of the CodeRocket Window if it has moved
		SetTitleMatchMode, 2
		WinGetPos, x, y, w, h, Just Paper, , SciTE4AutoHotkey
		SetTitleMatchMode, 1
	
		
		If(w>400 AND h>400)
			{
			if((x<>JustPaperWindowX OR y<>JustPaperWindowY OR w<>JustPaperWindowW OR h<>JustPaperWindowH) AND x<>-32000)
				{
					JustPaperWindowX := x	
					JustPaperWindowY := y
					JustPaperWindowW := w
					JustPaperWindowH := h
			
					IniWrite, %JustPaperWindowX%, %A_MyDocuments%\JustPaper.ini, Window Positions, JustPaperWindowX
					IniWrite, %JustPaperWindowY%, %A_MyDocuments%\JustPaper.ini, Window Positions, JustPaperWindowY
					IniWrite, %JustPaperWindowW%, %A_MyDocuments%\JustPaper.ini, Window Positions, JustPaperWindowW
					IniWrite, %JustPaperWindowH%, %A_MyDocuments%\JustPaper.ini, Window Positions, JustPaperWindowH
				}
			}

			; This sections checks the status of the WinSURGE Window and formats accordingly
			ifWinNotExist, WinSURGE
				{
				lastWinSurgeTitle=""
				z1 := "WinSURGE Not Open"
				Gosub, BuildMainGui
				}

			ControlGetText, x, Edit1, Caris CodeRocket v3.0.exe
			IfInstring, x, -
			{
				StringSplit, casenum, x, %A_Space%
				z1 := casenum1
				lastWinSURGEtitle := z1
				caseFromCodeRocket := 1
				Gosub, BuildMainGui
				Return
			}
			else
			{
				If caseFromCodeRocket    ;  The last time we redrew the GUI, there was a case in the case loader
				{
					Sleep, 1500
					caseFromCodeRocket := 0
					return
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
			StringReplace, x, x, Case, |, All
			StringSplit, y, x, |, %A_Space%
			StringSplit, z, y2, %A_Space%, %A_space%
			Gosub, BuildMainGui
	

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

GuiClose:				;Automation or synthesis
{
	ExitApp	
	Return
}

CheckForWarnings:
{
	percP := "%%P%%"
	
	;Msgbox, %finaldxContents%	
	LineArray := StrSplit(finalDxContents, "`n")
	jarText := []
	
	;Establish the contents of the final diagnosis box into the component jar texts
	Loop, 26
	{
		jarIndex := A_Index
		thisJar := Chr(A_Index+64)
		nextJar := Chr(A_Index+65)
		start := stop := 0
		
		Loop % LineArray.MaxIndex()
		{
			skipThisLine := 0
			this_line := LineArray[A_Index]
			
			If (Instr(this_line, thisJar . ".") AND Instr(this_line, percP))
				{
					start := 1
					skipThisLine := 1
				}
			If (Instr(this_line, nextJar . ".") AND Instr(this_line, percP))
				stop := 1
			
			if(start and !stop and !skipThisLine)
				jarText[jarIndex] := jarText[jarIndex] . this_line . "`n"
		}
		
	}
	
	;Loop Through jar text to find potential warnings
	{
		Loop % jarText.MaxIndex()
		{	
			thisJarIndex := A_Index
			;Msgbox, % jarText[thisJarIndex]
			;Get the Clinical History for that Jar
		thisJar := Chr(thisJarIndex+64)
		nextJar := Chr(thisJarIndex+64 +1)

		;Msgbox, thisJar=%thisJar%`nnextJar=%nextJar%
		
		;Msgbox, % thisJarIndex
		
		tempVar=<br>%nextJar%.
		;Msgbox, searchString = '%tempVar%'
		StringReplace, tempVar2, ClinicalData, %tempVar%, ¥, All
		;Msgbox, tempVar2=%tempVar2%
		StringSplit, clinDataSplit, tempVar2, ¥
		;Msgbox, clinDataSplit1=%clinDataSplit1%
		
		If (thisJar="A")
			tempVar=A.%A_Space%
		else
			tempVar=<br>%thisJar%.%A_Space%
		
		StringCaseSense, On
		StringReplace, tempVar2, clinDataSplit1, %tempVar%, ¥, All
		StringCaseSense, Off
		;Msgbox, %tempVar2%
		StringSplit, clinDataSplits, tempVar2, ¥
		thisJarClinicalData1:=clinDataSplits%clinDataSplits0%
		;Msgbox, %thisJarClinicalData1%
			
		;Get the first WinSurge Result Key that starts with a letter
			x := RegexMatch(jarText[thisJarIndex], "%%([a-z].+?)%%", resultKey)
			resultKey1=-%resultKey1%-
			
			
				If(InStr(alertFlags, "alert-on-use-") AND Instr(alertFlags, resultKey1))
				{
								SoundBeep 
								SoundBeep
								Msgbox, , Possible Preference Code Violation!!!, The client has requested a unique preference for this code.  Please double-check the preferences?

				}
		
	if InStr(jarText[thisJarIndex], "COMPLETELY EXCISED") OR InStr(jarText[thisJarIndex], "EXTENDING TO") OR InSTr(jarText[thisJarIndex], "EXAMINED MARGINS")
		marginPresent := 1
	else
		marginPresent := 0
	
	;Msgbox, % thisJarIndex . "----" . jarText[thisJarIndex] . "`n" . marginPresent
	
			if (!marginPresent AND RegexMatch(jarText[thisJarIndex], "%%([a-z].+?)%%", resultKey))
				{

					tempVar := Instr(thisJarClinicalData1, "margin") AND !InStr(jarText[thisJarIndex], "No Residual")
					tempVar += Instr(alertFlags, "margins-all-nevi") AND (Instr(jarText[thisJarIndex], "nevus") OR Instr(SelDiagnosis, "melanocy"))  ;For All nevi
					if (tempVar)
					{
						SoundBeep 
						SoundBeep
						Msgbox, ,Margin Warning!!!, The client has requested a margin and you did not use one.  Please double-check and correct your work if necessary.
					
					}
 				}
				else if (marginPresent)
					{
				
					tempVar := Instr(alertFlags, "none")  ;Warn if they have requested no margins
					tempVar += Instr(alertFlags, "no-margins-nmsc") AND (Instr(jarText[thisJarIndex], "carcinoma") OR Instr(jarText[thisJarIndex],"actinic keratosis")) ;warn if they have requested no margins on NMSC
					if (tempVar AND !Instr(thisJarClinicalData1, "margin"))
						{
							SoundBeep
							SoundBeep
							Msgbox, ,Margin Warning!!!,This client has requested the following margin preferences (%alertFlags%) and you have used one!  Please double-check and correct your work if necessary.	
						}
					}	

		firstCharacter := SubStr(jarText[thisJarIndex], 1, 1)
		
		j := Asc(firstCharacter)
		If j>96
			lowerCaseProblem = 1
		else
			lowerCaseProblem = 0
		
		resultKeyPresent := RegExMatch(jarText[thisJarIndex], "%%([a-z].+?)%%")
		if (lowerCaseProblem and resultKeyPresent)
			{
			SoundBeep
			SoundBeep
			Msgbox,  STOP! There appears to be a lower case letter where the TOP LINE Diagnosis should be!
			Return
			}
	
		;Improper crr code used
		tempVar := Instr(jarText[thisJarIndex], "Complete removal") AND !Instr(jarText[thisJarIndex], "dysplastic")
		If tempVar
		{
			SoundBeep
			SoundBeep
			Msgbox, , Complete removal warning!, You have recommended complete removal without the presence of a "dysplastic" nevus.  Please double-check and correct your work if necessary.
		}
	}
	return
	}
}

F1::
{
	SetTitleMatchMode, 2
	WinGetPos, x, y, w, h, Just Paper.ahk, ,SciTE4
	
	If(x<500) ;Window is minimized
	{
		WinActivate, Just Paper.ahk, , SciTE4
		WinActivate, Caris CodeRocket v2.5a.exe
		return
	}
	
	WinMinimize, Just Paper.ahk, ,SciTE4
	WinActivate, Caris CodeRocket v2.5a.exe
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

^!a::			;Paper Replacer
{	
	Send, ^c
	StringReplace, clipboard, clipboard,`n,,All
	StringReplace, clipboard, clipboard,`r,,All
	
	IfNotInString, displayedPreferences, %clipboard%
	{
		Msgbox, There is a problem with the text "%clipboard%" that prevents it becoming a smart substitution
		return
	}

	
	InputBox, alertCodes, Add A Diagnosis Code Alert..., What diagnosis codes (separated by hyphens) would you like %clipboard% to alert you on?
	
	If ErrorLevel
		return

	reText=<span style='background-color:orange'>%clipboard%</span>

	FileAppend, "%clipboard%"`,"%reText%"`,"alert-on-use-%alertCodes%-"`n, %A_MyDocuments%\MySmartSubstitutions.csv
	LoadSmartSubstitutions()
	CurrentCodeRocketDisplayedCase := 0
	Gosub, BuildMainGui
	return
}

^!m::			;Paper Replacer
{	
	Send, ^c
	StringReplace, clipboard, clipboard,`n,<br>,All
	StringReplace, clipboard, clipboard,`r,,All
	
	IfNotInString, displayedPreferences, %clipboard%
	{
		Msgbox, There is a problem with the text "%clipboard%" that prevents it becoming a smart substitution
		return
	}
	
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

^!u::  ;Undo current case and set for next scan
{
	WinActivate, WinSURGE
	Send, !3
	WinWaitActive, WinSURGE, Cancel
	Send {Left}{Enter}
	Sleep, 1500
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
\b0  %rawPreferences%\par
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
	
	WinWaitActive,  WordPad
	Sleep, 500
	SendLevel, 10
	Send ^p
	Sleep, 500
	Send !g
	Sleep, 100
	Send 1
	sleep, 100
	Send {Enter}
	Sleep, 100
	SendLevel, 0
	
	;~ WinWaitActive, Print
	;~ Sleep, 500
	;~ Send {Enter}
	;~ WinWaitActive, %CurrentCaseNumber%
	;~ Sleep 500
	;~ Send !f
	;~ sleep, 100
	;~ Send x
	;~ SendLevel, 0
	
Return
}

^!x::	;Cases for a specific client
{

			s := "select s.number, s.numberofspecimenparts, s.dx from specimen s where s.custom04 ='MI6970D' and s.sodate >= '2018-11-01'"
		WinSurgeQuery(s)
		Msgbox, %msg%
		FileAppend, %msg%, C:\Users\mmuenster\Desktop\LegacyClient.txt
		return

}

^!k::
{
	totalJarCount := 0
	daysInMonths := [31,29,31,30,31,30,31,31,30,31,30,31 ]
	;InputBox, monthToSearch, What Month and Year Do You Want To Search? MM-YYYY
	;StringSplit, m, monthToSearch, -
	;monthSearch := m1
	monthSearch := 12
	yearSearch := 2017
	;yearSearch := m2
	StringLeft, y, A_UserName, 5
	StringUpper, y, y
	y := "DWIMM"
	ComObjError(True)
	s = Select p.name, p.id, p.state, p.suid, p.wsid from pathologists p where p.abbr='%y%'
	CodeDataBaseQuery(s)
	pathSearch := Result_5
	Msgbox, %pathSearch%
	numdays := daysInMonths[monthSearch]
	Loop, %numdays%
	{
		jarcount:=casecount:=mxcount:=blockcount:=0

		if (thisJarIndex<10)
			daySearch=0%A_Index%
		else
			daysearch=%A_Index%
		
		FormatTime, dow, %yearSearch%%monthSearch%%daySearch%, dddd
		todaydate = %yearSearch%-%monthSearch%-%daySearch%
		s := "select s.number, s.numberofspecimenparts, s.dx, s.calculatedslidecount from specimen s where s.path =" . pathSearch . " and s.sodate = '" . todaydate . "'"
		WinSurgeQuery(s)
		Loop, parse, msg, `n 
		{
			If(A_LoopField)
			{
				
				StringSplit, res, A_LoopField, ¥
				
				caseNum=%res2%
				StringLeft, casePrefix, caseNum, 2
				x := get_filled_case_number(caseNum)
				s=select specimen_id from specimen where computed_numberfilled ='%x%'
				WinSurgeQuery(s)
				;Msgbox, %msg%
				StringSplit, specimenID, Result_1, .
				;Msgbox, %specimenID1%
				s=select max(h.histologyblock_id) from histologyblock h left join specimen s on h.specimen_id=s.specimen_id where h.specimen_id=(%specimenID1%) and h.prefix ='%casePrefix%' and h.prefix=s.prefix and NOT h.histologyblock_id  in (select histologyblock_id from histologyprocedure where specimen_id=%specimenID1% and prefix ='%casePrefix%' and histoproccancelleddate is NOT NULL)
				WinSurgeQuery(s)
				StringSplit, bc, Result_1, .
				
				;Msgbox, %caseNum%,%res3%,%blockCount1%
				

				casecount := casecount + 1
				jarcount := jarcount + res3
				if (bc1>0)
					blockcount := blockcount + bc1

				IfInString, res2, MX
					mxcount := mxcount + 1
			}
		}
		
		daysJarCount := jarcount + mxcount*2
		if(dow="Saturday" OR dow="Sunday")
			totalJarCount := totalJarCount + daysJarCount
		else if (daysJarCount>142)
			totalJarCount := totalJarCount + daysJarCount - 142
		
		appendText := pathSearch . "," . todaydate . "," . dow . "," . daysJarCount . "," . totalJarCount . "," . blockcount . "`n"
		
		FileAppend, %appendText%, C:\Users\mmuenster\Documents\j_b_summary.csv 
	}

	Msgbox, Done!
	return	
}

^!9::
{
	InputBox, daySearch, What Day in December 2017 Do You Want To Search? DD
	monthSearch := 11
	yearSearch := 2017
	pathName='Matthews, Mark R.'    ;233388
	todaydate = %yearSearch%-%monthSearch%-%daySearch%
	
		jarcount:=casecount:=mxcount:=blockcount:=tcpcCount:=0
		
		FormatTime, dow, %yearSearch%%monthSearch%%daySearch%, dddd

		s := "select s.number, s.numberofspecimenparts, s.resultkeylog, s.num05 from specimen s, physician p where s.path = p.id and p.name = " . pathName . " and s.sodate = '" . todaydate . "'"
		WinSurgeQuery(s)
		Msgbox, %msg%
		Loop, parse, msg, `n 
		{
			If(A_LoopField)
			{
				
				StringSplit, res, A_LoopField, ¥
				
				caseNum=%res2%
				StringLeft, casePrefix, caseNum, 2
				StringSplit, bc, res5, .
				casecount := casecount + 1
				jarcount := jarcount + res3
				
				If(casePrefix="PY" OR casePrefix="DP")
					{
						tcpcCount:=tcpcCount + 1
						x := get_filled_case_number(caseNum)
						s=select max(h.histologyblock_id) from histologyblock h left join specimen s on h.specimen_id=s.specimen_id where h.specimen_id=(select specimen_id from specimen where computed_numberfilled ='%x%') and h.prefix ='%casePrefix%' and h.prefix=s.prefix and NOT h.histologyblock_id  in (select histologyblock_id from histologyprocedure where specimen_id=(select specimen_id from specimen where computed_numberfilled ='%x%') and prefix ='%casePrefix%' and histoproccancelleddate is NOT NULL)
									WinSurgeQuery(s)
									StringSplit, bc, Result_1, .\
									;Msgbox, %x%,%bc1%
					}


				
				if (bc1>0)
					blockcount := blockcount + bc1
				
				if(!blockcount)
					Msgbox, %caseNum%,%res3%,%bc1%,%blockcount%

				If(casePrefix="MX")
					mxcount := mxcount + 1
				
				If(casePrefix="PY" OR casePrefix="DP")
					tcpcCount:=tcpcCount + 1
				
				Msgbox, %caseNum%,%res3%,%bc1%,%blockcount%
				
			}
		}
		
		daysJarCount := jarcount + mxcount*2
		
		appendText := pathName . "," . todaydate . "," . dow . "," . daysJarCount . "," . blockcount . "," . mxcount . "," . tcpcCount .  "`n"
		
		Msgbox, %appendText%
		FileAppend, %appendText%, C:\Users\mmuenster\Documents\j_b_summary.csv 


	Msgbox, Done!
	return	
}

^!o::
{
	WinActivate, WinSURGE
	Send, !3
	WinWaitActive, WinSURGE, Cancel
	Send {Left}{Enter}
	Sleep, 1500
	Send, {F12}
	return
return
}

^!t::  ;List of all cases distributed today but not signed out
{
	unsignedCaseList := ""

	Loop, 5
	{
		SplashTextOn, 200, 200, Unsigned Cases, Day %A_Index%
		
		day = %a_now%
		;day += -1, days
		day += -%A_Index%, days
		FormatTime, day, %day%, yyyy-MM-dd 
		;Msgbox, %day%
		
		tempVar := "http://s-irv-autoasgn/autoassign2/report_path_case_status.php?order_by=sodate&date=" . day
		;Msgbox, %tempVar%
		URLDownloadToFile, %tempVar%, distHtml.txt

		FileRead, html, distHTML.txt
		FileDelete, distHTML.txt

	;Msgbox, % html
		document := ComObjCreate("HTMLfile")
		document.write(html)
		all := document.getElementsByTagName("table")

		Sleep, 1000
		tempVar := 2
		table := all[tempVar]


		Loop, % table.rows.length - 1
		{	
			tempVar := A_Index-1
			tempVar2 := table.rows[tempVar].cells[0].innerHTML
			;Msgbox, % table.rows[tempVar].cells[4].innerHTML
			
			
			if (table.rows[tempVar].cells[1].innerHTML>0 AND table.rows[tempVar].cells[2].innerHTML<>"&nbsp;" AND table.rows[tempVar].cells[4].innerHTML="&nbsp;")
				unsignedCaseList=%unsignedCaseList%%tempVar2%`n
			
		}

	;Msgbox, %unsignedCaseList%
	}
	SplashTextOff
	Msgbox, Cases Distributed Last 5 days But Not Signed Out NOT INCLUDING TODAY`n--------------------------------------------`n %unsignedCaseList%
	
	Return

}

^!w::  ;Distribution Summary
{
URLDownloadToFile, http://s-irv-autoasgn/autoassign2/report_path_case_status.php, distHtml.txt

FileRead, html, distHTML.txt
FileDelete, distHTML.txt

;  Login info  Username=mmuenster  Password=mmuenster@2015

document := ComObjCreate("HTMLfile")
document.write(html)
all := document.getElementsByTagName("table")
	
Sleep, 1000
tempVar := 2
table := all[tempVar]

notDistributedCount := distributedCount := 0

Loop, % table.rows.length - 1
{	
	tempVar := A_Index-1
	if (table.rows[tempVar].cells[1].innerHTML>0 AND table.rows[tempVar].cells[2].innerHTML<>"&nbsp;")
		distributedCount += table.rows[tempVar].cells[1].innerHTML
	else
		notDistributedCount += table.rows[tempVar].cells[1].innerHTML
	
}

	totalJarsInQueue := distributedCount + notDistributedCount
	Msgbox, Distribution Summary`n-------------------------`n %distributedCount% of %totalJarsInQueue% distributed
	Return

}

^!y::   ;Hotkey to find out who signs out what % of a client's cases
{
	names := []
	nameCounts := {}
	
	InputBox, clientID, ,Enter the Client ID you want to search...,
	
	s := "select s.number, p.name, s.numberofspecimenparts, s.sodate from specimen s, physician p where s.custom04='" . clientID . "' and s.path=p.id and s.sodate>='2019-03-17'"
	;. ", physician p, patient pt where s.patient = pt.id and s.clin=p.id and computed_numberfilled='" . x . "'"
	WinSurgeQuery(s)
	FileAppend, %msg%, %clientID%.txt

	
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

^!b::Reload

^!4::
{
		percP := "%%P%%"
		
		IfWinExist, WinSURGE - Final Diagnosis:
		{
				ControlGetText, t, TX202, WinSURGE - Final Diagnosis:
				If (t<>finaldxContents)
					Gosub, CheckForWarnings
				
				finaldxContents := t
				finaldxActive := 1
		}

	IfWinNotExist, WinSURGE - Final Diagnosis:
		{
			if (finaldxActive=1)
				Gosub, CheckForWarnings
			
			finaldxContents := ""
			finaldxActive := 0
		}

	LineArray := StrSplit(finalDxContents, "`n")
	jarText := []
	
	Loop, 26
	{
		jarIndex := A_Index
		thisJar := Chr(A_Index+64)
		nextJar := Chr(A_Index+65)
		start := stop := 0
		
		Loop % LineArray.MaxIndex()
		{
			skipThisLine := 0
			this_line := LineArray[A_Index]
			
			If (Instr(this_line, thisJar . ".") AND Instr(this_line, percP))
				{
					start := 1
					skipThisLine := 1
				}
			If (Instr(this_line, nextJar . ".") AND Instr(this_line, percP))
				stop := 1
			
			if(start and !stop and !skipThisLine)
				jarText[jarIndex] := jarText[jarIndex] . this_line . "`n"
		}
		
	}
	
	Loop % jarText.MaxIndex()
		Msgbox, % jarText[A_Index]

	return
}

^!3::   ;Test Hotkey for debugging
{
s := "select s.number, s.custom04, p.proficiencylog, z.proficiencylog from specimen s, physician p, physician z, patient pt where s.patient = pt.id and s.clin=p.id and s.client=z.id and s.sodate>'2019-01-18'"

		WinSurgeQuery(s)

;s:= "select s.number, s.dx from specimen s, physician p where s.clin=p.id and s.dx LIKE '%dfsp%'  and s.sodate > '2018-05-01'"
;s:= "select s.number, s.dx, s.sodate from specimen s where s.dx LIKE '%alopecia%' and s.sodate > '2018-01-01'"
;s=select * from specimen where computed_numberfilled='DD18-173265'

;WinSurgeQuery(s)
Msgbox, %msg%
FileAppend, %msg%, preferencelog.txt

return
}

^!g::	;Debugging
{
	ListVars
	Return
}

^!h::  ;Debugging
{
	ListLines
	Return
}

^!2::Msgbox, %alertFlags%

F6::
{
if(A_Username<>"mmuenster")
	return

	;Run, "C:\Users\mmuenster\Desktop\PR Development\KeepActive.ahk"

	Process,Close,WinSURGE.exe
	Process,WaitClose,WinSURGE.exe,2
	If(!ErrorLevel)
	{
		Run, "C:\Program Files (x86)\WinSURGE\WinSURGE.exe"
		WinWaitActive, WinSURGE
		tempVar := "W{!}nter2019"
		Send, %tempVar%
		Send, {Enter}
		WinWaitActive, Login Message
		Send, {Enter}
		
		WinWaitActive, WinSURGE
		Click, 760, 100
	}
	else
	{
		Msgbox, Could not close WinSurge
	}

	return
}

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
^!z::Reload		;Debugging

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

	array312 := ["PAS","acid-Schiff","GMS","AFB","Fite","Giemsa","Gram","Mucicarmine","Steiner","PAS-positive","Schiff","acid-fast"]
	array313 := ["Fontana-Masson","Fontana","alcian","Congo","VVG","oil","Prussian","Iron","reticulin","violet"]
	array342 := ["Mart1","Mart-1","S100","S-100","tryptase","Alk-1","Alk1","Ber-EP4","Beta-catenin","cKit","c-kit","c-myc","CD10","CD138","CD1a","CD20","CD21","CD23","CD3","CD30","CD31","CD34","CD4","CD43","CD45","CD5","CD56","CD68","CD7","CD79A","CD8","CDX2","CEA","chromogranin","CK19","CK20","CK5/6","ck8/18","ae1/ae3","cam5.2","Factor","cd99","ck7","desmin","actin","gata-3","hmb45","IgA","IgG","kappa","lambda","mammoglobin","Mitf","sox10","p16","p63","pancytokeratin","PSA","RCC","myosin","TTF1","TTF-1","Pax8","HSV","VZV","pallidum","vimentin","antigen","EMA","CD20+","CD3+","ki-67","KI67","p40","mib","mib1","CD45/LCA","synaptophysin","chromogranin","(CD3+)","(CD20+)","LCA","CK","oscar","HHV8","hhv-8","erg"]
	array365 := ["(16,18)","(16,","(6,11)","(6,"]


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

ReadIniValues()
{
	global

	IniRead, JustPaperWindowX, %A_MyDocuments%\JustPaper.ini, Window Positions, JustPaperWindowX 	
	IniRead, JustPaperWindowY, %A_MyDocuments%\JustPaper.ini, Window Positions, JustPaperWindowY 
	IniRead, JustPaperWindowW, %A_MyDocuments%\JustPaper.ini, Window Positions, JustPaperWindowW 
	IniRead, JustPaperWindowH, %A_MyDocuments%\JustPaper.ini, Window Positions, JustPaperWindowH

	IniRead, UseSmartSubstitutions, %A_MyDocuments%\JustPaper.ini, Window Positions, UseSmartSubstitutions
	
Return
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
		displayedPreferences := ""
		displayedClinicalData := ""
		displayedFinalDiagnosis := ""
		displayedGrossDescription := ""
	}
	else
	{
		GuiControl, Enable, CaseScanBox 
		GuiControl, , CaseNumberLabel, %CurrentCaseNumber%
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

	;Msgbox, 3508 - %rawPreferences%
	
		StringReplace, displayedPreferences, displayedPreferences,`r,,All
		StringReplace, displayedPreferences, rawPreferences, `n,<br>,All
		StringReplace, displayedPreferences, displayedPreferences,¥,,All
		StringReplace, displayedPreferences, displayedPreferences,***,<br>, All
		StringReplace, displayedPreferences, displayedPreferences,%A_Space%%A_Space%,%A_Space%,All
		StringReplace, displayedPreferences, displayedPreferences,%A_Space%%A_Space%%A_Space%,%A_Space%,All

		;Sample for testing certain replacements
		;StringReplace, displayedPreferences, displayedPreferences, Diagnostic Text:, Matt Muenster, All


		FileDelete, displayedPreferences.txt
		FileAppend, %displayedPreferences%, displayedPreferenes.txt

	;SMART Replacements start here
		if UseSmartSubstitutions
		{
			alertFlags := ""

			Loop, 2  ;This is required to get the links to work that are below the first substitution
			{
				Loop, % smartSubs.Length()+1
				{
					m := smartSubs[A_Index].searchText
					n := smartSubs[A_Index].replacementText
					o := smartSubs[A_Index].flag
					
					if (Instr(displayedPreferences, m))   ;AND !Instr(displayedPreferences, n))
						{
							;Msgbox, %displayedPreferences%`n%m%`n%n%`n%o%
							StringReplace, displayedPreferences, displayedPreferences, %m%, %n%, All
							if (o AND !InStr(alertFlags, o))
								alertFlags=%alertFlags%   %o%
						}
						
					StringReplace, displayedClinicalData, displayedClinicalData, %m%, %n%, All
					StringReplace, displayedGrossDescription, displayedGrossDescription, %m%, %n%, All
				}
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

		if (Instr(procedureNote,"File") OR Instr(procedureNote,"Image"))
			WB.document.getElementById("procedureNote").innerHTML := "****PROCEDURE NOTE AVAILABLE ON PATHOLOGIST TAB****<br><br>"
		else
			WB.document.getElementById("procedureNote").innerHTML := ""
		
		if (Instr(additionalClinicalInformation, "File") OR Instr(additionalClinicalInformation, "Image"))
			WB.document.getElementById("additionalClinicalInformation").innerHTML := "****ADDITIONAL CLINICAL INFORMATION AVAILABLE ON PATHOLOGIST TAB****<br>"
		else
			WB.document.getElementById("additionalClinicalInformation").innerHTML := ""
		WB.document.getElementById("caseNumber").innerHTML := CurrentCodeRocketDisplayedCase . " ---- " . PatientName . " (" . PatientAge . ")<br>"	. ClientOfficeName . " (" . ClientState . ")<br>"
		WB.document.getElementById("preferences").innerHTML := displayedPreferences
		WB.document.getElementById("clinicalInformation").innerHTML := displayedClinicalData
		WB.document.getElementById("finalDiagnosis").innerHTML := displayedFinalDiagnosis
		WB.document.getElementById("grossDescription").innerHTML := displayedGrossDescription
		WB.document.getElementById("priorCaseInformation").innerHTML := priorCaseInfo
	}

;Msgbox, 3665 - %alertFlags%
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


class Event {
	DocumentComplete(wb) {
		static doc
		ComObjConnect(doc:=wb.document, new Event)
	}
	OnKeyPress(doc) {
		static keys := {1:"selectall", 3:"copy", 22:"paste", 24:"cut"}
		keyCode := doc.parentWindow.event.keyCode
		if keys.HasKey(keyCode)
			Doc.ExecCommand(keys[keyCode])
	}
}