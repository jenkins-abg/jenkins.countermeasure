#include-once

#cs	==================================================================================================================
	Title 				:	jnknsMainInitialize
	AutoIt Version	: 	3.3.14.5
	Language		: 	English
	Description		:	Initialize variables needed
	Author				: 	prdedumo
    Version            :    0.1
#ce	==================================================================================================================

;~ #AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w- 7

; #INCLUDES# ===========================================================================================================
	#include <ButtonConstants.au3>
	#include <EditConstants.au3>
	#include <GUIConstantsEx.au3>
	#include <MsgBoxConstants.au3>
	#include <Excel.au3>
	#include <File.au3>
	#include <Misc.au3>

	#include "ErrHandler\jnknsErrHandler.au3"
    #include "TraceLog\jnknsProcessLogger.au3"
; ====================================================================================================================

; #GLOBAL VARIABLES# =====================================================================================================
Global		$g_sJMI_TestDesign_File, _																									;	Test Design File name
				$g_sJMI_Spider_Version, _																									;	Spider Version Name
				$g_sJMI_TPRJ_Path, _																											;	TPRJ path on the build software
				$g_sJMI_Spider_Latest = 'C:\Program Files (x86)\DENSO\D-SPIDER\D-SPIDER.exe', _		;	DSpider exe file path
				$g_sJMI_Spider_Latest_Title = 'ソフトウェア単体テスト自動化ツール D-SPIDER  Ver.1.0.0', _		;	DSpider Title name
				$g_sJMI_Spider_Old = 'C:\Program Files (x86)\DENSO\FSUnitTest\FSUnitTest.exe', _		;	OLD FSUnit exe file path
				$g_sJMI_Spider_Old_Title = '機能安全対応ソフトウェア単体テストツール  ver.0.9.13.1', _			;	OLD FSUnit Title Name
                $g_sJMI_StatusClass = ""
Global		$g_iJMI_Spider_win_x = 0, _																								;	DSpider X-position in the monitor
				$g_iJMI_Spider_win_y = 0, _																								;	DSpider Y-position in the monitor
				$g_iJMI_Spider_win_width = 691, _																						;	DSpider resize-width
				$g_iJMI_Spider_win_height = 202, _																					;	DSpider resize-height
                $g_iJM_Spider_Software_Path_Class, _
                $g_iJM_Assembly_TPRJ, _
                $g_iJEH_PLError_Check, _
                $g_iJM_Assembly_Error
; ====================================================================================================================

; #CURRENT# ============================================================================================================
;_JMI_jnknsShowForm
;_JMI_jnknsCallDSpider﻿
;_JMI_jnknsSpiderSettings
;_JMI_jnknsPressF5
;_JMI_jnknsBuildTree
;_JMI_jnknsWinGetClassesByText
;_JMI_jnknsWinGetControlIDs
;_JMI_jnknsAddClass
;_JMI_jnknsBrowseFolder
; ====================================================================﻿================================================

; #FUNCTION# ===========================================================================================================
; Name					:	_JMI_jnksEnvironmentLog
; Description		:	Creates initial log text to set the test design and other required parameters.
; Syntax				:	_JMI_jnksEnvironmentLog()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	0 - No Error Occured
;                  				1 - Error Occured
; Author				:	prdedumo
; Modified			:	None
;=====================================================================================================================
Func _JMI_jnksEnvironmentLog()
    Local   $sLogTextFile = @ScriptDir & "\Log.txt"
    Local   $hTextFile = FileOpen($sLogTextFile, $FO_READ)
    Local   $sTestSheetFile
    Local   $sTPRJPath
    Local   $sTrimPath
    Local   $ret

    $ret = 0
    ; Test Sheet Line number
    $sTestSheetFile = FileReadLine($hTextFile,2)
    $sTestSheetFile = StringTrimLeft($sTestSheetFile,20)
    $sTPRJPath = FileReadLine($hTextFile,1)
    $sTPRJPath = StringTrimLeft($sTPRJPath,11)
    $sTrimPath = StringStripWS($sTPRJPath,4)
    ;MsgBox($MB_SYSTEMMODAL, "",$sTrimPath)
    $g_sJMI_TestDesign_File = $sTestSheetFile
    If $sTrimPath = "" Then
        If _JMI_jnknsSpiderSettings() Then
            ;_JMI_jnknsInitLog($g_sJMI_Spider_Version)
        EndIf
        $ret = 0
    Else
        _JMI_jnknsSpiderSettings()
		;_JMI_jnknsInitLog($g_sJMI_Spider_Version)
        $ret = 1
    EndIf
    Return $ret
EndFunc ;==>_JMI_jnksEnvironmentLog

; #FUNCTION# ===========================================================================================================
; Name					:	_JMI_jnknsShowForm
; Description		:	Launches GUI Form and set the test design and other required parameters.
; Syntax				:	_JMI_jnknsShowForm()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	0 - No Error Occured
;                  				1 - Error Occured
; Author				:	prdedumo
; Modified			:	None
;=====================================================================================================================
Func _JMI_jnknsShowForm()
	Local	$sForm_Title = "Get Test Design", _
				$sTestDesign_File = ""
	Local	$iErrHandler = 0, _
				$iReturnCallDSpider, _
				$iReturnShowForm

	if @error Then
		$iReturnShowForm = 1
		$iErrHandler = 1
		Exit
	EndIf
	$iReturnShowForm = 0
	; GUI form
	#Region ### START Koda GUI section ### Form=
		$form_GTD = GUICreate($sForm_Title, 559, 45, 199, 136)
		GUISetFont(9, 400, 0, "ＭＳ Ｐゴシック")
		$cmd_GetTD = GUICtrlCreateButton("Get", 424, 8, 33, 25)
		$sTestDesign_File = GUICtrlCreateInput("SUT Test Design here...", 16, 8, 393, 25,$ES_READONLY)
		$cmd_Run = GUICtrlCreateButton("Run Test", 472, 8, 75, 25)
	#EndRegion ### END Koda GUI section ###
	; Checks if there is only one instance of the window
	if _Singleton($sForm_Title,1) <> 0 then
		GUISetState(@SW_SHOW)
	endif
	; Activates the form so that it won't be hidden
	WinSetOnTop($form_GTD,"",1)
	; Check action, loop until canceled by the user
	While 1
		; Checks action
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE		; Cancelled by the user
				Exit
			Case $cmd_GetTD					; Search File
				_JMI_jnknsBrowseFolder()
				GUICtrlSetData($sTestDesign_File,$g_sJMI_TestDesign_File)
				; Checking condition for the assembly error
				; Edited by: cjhernandez
                ; Edit starts here
                ; ======================================
				if $g_iJM_Assembly_Error Then
					$g_iJM_getWin = WinActivate( $sForm_Title )
					$g_iJM_errorlogFile = WinGetText ( $g_iJM_getWin )
					ExitLoop
				EndIf
                ; ======================================
                ; Edit ends here
			Case $cmd_Run						; Run Test
				; Checks if TestDesign is blank
				if $g_sJMI_TestDesign_File <> "" Then
					; Checks version of DSpider tool
					$iReturnCallDSpider = _JMI_jnknsCallDSpider()
					if $iReturnCallDSpider = 0 Then
						Exit
					Else
						Sleep(2000)
						WinSetState($g_sJMI_Spider_Latest_Title,"",@SW_MAXIMIZE)
						If _JMI_jnknsSpiderSettings($g_sJMI_Spider_Version) Then
                            ; create initial log text for assembly
                            _JMI_jnknsInitLog($g_sJMI_Spider_Version)
                            ; removed the pressF5
							If _JMI_jnknsPressF5($g_sJMI_Spider_Version) Then
								If $g_iJEH_FormClose Then	ExitLoop
							EndIf
						EndIf
					EndIf
				Else
				EndIf
		EndSwitch
	WEnd
	Return $iReturnShowForm
EndFunc		;==>_JMI_jnknsShowForm

; #FUNCTION# =========================================================================================================
; Name					:	_JMI_jnknsCallDSpider
; Description		:	Sets DSpider Parameters
; Syntax				:	_JMI_jnknsCallDSpider()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	0 - No DSpider Exists in the window
;                  				1 - DSpider Exists in the window
; Author				:	prdedumo
; Modified			:	None
;=====================================================================================================================
Func _JMI_jnknsCallDSpider()
	Local $hRunHandler
	Local $hTestToolHandler
	; Check if D-Spider Latest UnitTest Tool is already running
	if ProcessExists("D-SPIDER.exe") Then
		$hTestToolHandler = WinGetHandle($g_sJMI_Spider_Latest_Title)
		WinMenuSelectItem($hTestToolHandler,"","","")
		;WinActivate($g_sJMI_Spider_Latest_Title,"")
		;WinSetState($g_sJMI_Spider_Latest_Title,"",@SW_MAXIMIZE)
	Else
		; Run D-Spider, returns error if no D-Spider installed
		$hRunHandler = Run($g_sJMI_Spider_Latest)
		if @error Then
			if ProcessExists("FSUnitTest.exe") Then
				$hTestToolHandler = WinGetHandle(g_sJMI_Spider_Old)
				WinMenuSelectItem($hTestToolHandler,"","","")
				;WinActivate($g_sJMI_Spider_Old,"")
				WinSetState($g_sJMI_Spider_Old,"",@SW_MAXIMIZE)
			Else
				$hRunHandler = Run($dspider_non_upgraded)
				$hTestToolHandler = WinGetHandle(g_sJMI_Spider_Old_Title)
				if @error Then
					Exit
				EndIf
				; Wait until FSUnitTest opens
				WinWait(hTestToolHandler,"","")
				;WinWaitActive($g_sJMI_Spider_Old_Title,"","")
			EndIf
		EndIf
		; Wait until D-Spider Opens
		WinWait(g_sJMI_Spider_Latest_Title,"","")
		;WinWaitActive($g_sJMI_Spider_Latest_Title,"","")
	endif
	; Checks if Window exists
	if WinExists($g_sJMI_Spider_Latest_Title) Then
		WinMove($g_sJMI_Spider_Latest_Title, "", $g_iJMI_Spider_win_x, $g_iJMI_Spider_win_y, $g_iJMI_Spider_win_width, $g_iJMI_Spider_win_height)
		$g_sJMI_Spider_Version = $g_sJMI_Spider_Latest_Title
        $g_sJMI_Spider_Process = "D-SPIDER.exe"
		Return 1
	ElseIf WinExists($g_sJMI_Spider_Old_Title) Then
		WinMove($g_sJMI_Spider_Old_Title, "", $g_iJMI_Spider_win_x, $g_iJMI_Spider_win_y, $g_iJMI_Spider_win_width, $g_iJMI_Spider_win_height)
		$g_sJMI_Spider_Version = $g_sJMI_Spider_Old_Title
        $g_sJMI_Spider_Process = "FSUnitTest.exe"
		Return 1
	Else
		Return 0
	EndIf
endFunc		;==>_JMI_jnknsCallDSpider

; #FUNCTION# =========================================================================================================
; Name					:	_JMI_jnknsSpiderSettings
; Description		:	Sets DSpider Parameters
; Syntax				:	_JMI_jnknsSpiderSettings(ByRef $spiderTitle)
; Parameters		:	$spiderTitle   - gets the current spider title - only required parameter!
; Requirement(s)	:	Requirement(s).: v3.3 +
; Return values		: 	Return values .: None
; Author				:	prdedumo
; Modified			:	None
;=====================================================================================================================
#comments-start
	Func _JMI_jnknsSpiderSettings(ByRef $sSpiderTitle)
	Local $sSpider_class_filename = $g_iJM_Spider_Software_Path_Class					;	Class value depending in the AutoIt v3 Window info
    $g_sJMI_TPRJ_Path = ControlGetText($sSpiderTitle,"", $sSpider_class_filename)
	Return 1
EndFunc		;==>_JMI_jnknsSpiderSettings

#comments-end

Func _JMI_jnknsSpiderSettings()
	Local   $sLogTextFile = "C:\work\Jenkins\automation-jenkins\Log.txt"
	Local   $hTextFile = FileOpen($sLogTextFile, $FO_READ)
	Local   $sTPRJPath
	Local   $sTrimPath
	Local   $ret
	Local   $sTestSheetFile

	; Test Sheet Line number
	$sTestSheetFile = FileReadLine($hTextFile,2)
	$sTestSheetFile = StringTrimLeft($sTestSheetFile,20)

	$sTPRJPath = FileReadLine($hTextFile,1)
	$sTPRJPath = StringTrimLeft($sTPRJPath,11)
	$sTrimPath = StringStripWS($sTPRJPath,4)

	; Assign to global
	$g_sJMI_TPRJ_Path = $sTrimPath
	$g_sJMI_TestDesign_File = $sTestSheetFile

	Return 1
EndFunc		;==>_JMI_jnknsSpiderSetting




; #FUNCTION# =========================================================================================================
; Name					:	_JMI_jnknsPressF5
; Description		:	Launches the test design in the DSpider
; Syntax				:	_JMI_jnknsPressF5(ByRef $spiderTitle)
; Parameters		:	$spiderTitle   - gets the current spider title - only required parameter!
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:	prdedumo
; Modified			:	None
;=====================================================================================================================
Func _JMI_jnknsPressF5($sSpiderTitle)
	Local	$sSpider_F5_Class = $g_iJM_Spider_F5_Class, _												;	Class value depending in the AutoIt v3 Window info
				$sSpider_File_Class = $g_iJM_Spider_File_Class, _											;	Class value depending in the AutoIt v3 Window info
				$sSpider_Software_Path_Class = $g_iJM_Spider_Software_Path_Class, _		;	Class value depending in the AutoIt v3 Window info
				$sSpider_Run_Class = "テスト実行中", _															;	Class value depending in the AutoIt v3 Window info
				$sSpider_Path = "", _																						;	Setting initial value to null
				$sSpider_Software_Path = "", _																		;	Setting initial value to null
				$sSSpider_Local = "", _																					;	Setting initial value to null
				$sSpider_Log_TxtFile = "", _																				;	Setting initial value to null
				$sUnitTest_Log_TxtFile = ""
	Local	$iReturnF5
	Local	$hTestToolHandler

	$iReturnF5 = 0
	$sSpider_Software_Path = ControlGetText($sSpiderTitle,"",$sSpider_Software_Path_Class)
	$sSpider_Path =  StringTrimRight($sSpider_Software_Path,21)
	$sUnitTest_Log_TxtFile = $sSpider_Path & "\UnitTest\log.txt"
	$sSpider_Log_TxtFile = @ScriptDir & '\Log.txt'
	$hTestToolHandler = WinGetHandle($sSpiderTitle)
	
	; Copy the test Design File
	;ClipPut($g_sJMI_TestDesign_File)
	; Send Keys
	;Send("{ALT}")
	;Send("{F}")
	;Send("{T}")
	;ControlSend($hTestToolHandler, "", "&F", "{T}")
	;ControlSend($sSpiderTitle, "", $sSpider_File_Class, "!aft")
	; Wait for 1 second
	;Sleep(2000)
	;WinWait("","",1)
	; Pastes the copied test design file path
	;Send("^v")
	;Send("{ENTER}")

	;Local $tprjHwnd = WinGetHandle("テストブックを選択")
	;$sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle("テストブックを選択"))
	;ControlFocus($tprjHwnd, "", $sTextClasses[8][0])   ; Set focus on the tprj input bar
	;ControlSend($tprjHwnd, "", "^v")
	;ControlSend($tprjHwnd, "", "{ENTER}")
	;Sleep(500)
	;WinWait("","",5)
	; Save the configuration of the DSpider
	;Send("^s")
	;ControlSend($tprjHwnd, "", "^s")
	Sleep(5000)
	;WinWait("","",5)
	; Presses F5 in the DSpider Tool
	ControlClick($sSpiderTitle,"",$sSpider_F5_Class)
	Sleep(5000)
	;WinWait("","",10)
	While (1)
		if WinExists($sSpider_Run_Class) Then
			ExitLoop
		EndIf
	WEnd

	;$sSpider_Local = WinActivate($sSpider_Run_Class)
	$sSpider_Local = WinGetHandle($sSpider_Run_Class)
	; Loop to wait until running of the tool is done
	While 1
		;$sSpider_Local = WinActivate($sSpider_Run_Class)
		$sSpider_Local = WinExists($sSpider_Run_Class,"") ;edited by ryan
;~             If _JPE_jnknsErrorLogger($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile) Then
;~                 _JEH_jnknsCheckErrHandler($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile)
;~                 $g_iJEH_PLError_Check = 1
;~                 ExitLoop
;~             EndIf
            if $sSpider_Local <> 0 Then
            Else
                ExitLoop
            EndIf
	WEnd
	Sleep(2000)
    If $g_iJEH_PLError_Check = 1 Then
        $iReturnF5 = 1
;~         While 1
;~             If _JPE_jnknsErrorLogger() Then
;~                 ExitLoop
;~             EndIf
;~         WEnd
    Else
        _JMI_jnknsReCheckIfError($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile)
    EndIf
	$iReturnF5 = 1
	
	Return $iReturnF5
EndFunc		;==>_JMI_jnknsPressF5

; #FUNCTION# =========================================================================================================
; Name					:	_JMI_jnknsInitLog
; Description		:	Initializes Log text file
; Syntax				:	_JMI_jnknsInitLog(ByRef $spiderTitle)
; Parameters		:	$spiderTitle   - gets the current spider title - only required parameter!
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:	prdedumo
; Modified			:	None
;=====================================================================================================================
Func _JMI_jnknsInitLog($sSpiderTitle)
    Local	$sSpider_F5_Class = $g_iJM_Spider_F5_Class, _												;	Class value depending in the AutoIt v3 Window info
				$sSpider_File_Class = $g_iJM_Spider_File_Class, _											;	Class value depending in the AutoIt v3 Window info
				$sSpider_Software_Path_Class = $g_iJM_Spider_Software_Path_Class, _		;	Class value depending in the AutoIt v3 Window info
				$sSpider_Run_Class = "テスト実行中", _															;	Class value depending in the AutoIt v3 Window info
				$sSpider_Path = "", _																						;	Setting initial value to null
				$sSpider_Software_Path = "", _																		;	Setting initial value to null
				$sSSpider_Local = "", _																					;	Setting initial value to null
				$sSpider_Log_TxtFile = "", _																				;	Setting initial value to null
				$sUnitTest_Log_TxtFile = ""
	Local	$iReturnF5
	Local	$hTestToolHandler

	$iReturnF5 = 0

	$sSpider_Software_Path = ControlGetText($sSpiderTitle,"",$sSpider_Software_Path_Class)
	$sSpider_Path =  StringTrimRight($sSpider_Software_Path,21)
	$sUnitTest_Log_TxtFile = $sSpider_Path & "\UnitTest\log.txt"
	$sSpider_Log_TxtFile = @ScriptDir & '\Log.txt'
	$hTestToolHandler = WinGetHandle($sSpiderTitle)

;~     Local $hWnd = WinWait($g_sJMI_Spider_Process)
;~     Local $hControl = ControlGetHandle($hWnd, "", "Edit1")
	; Copy the test Design File
	ClipPut($g_sJMI_TestDesign_File)
    WinActivate($sSpiderTitle)
    ; Send Keys

    Sleep(200)
    ControlSend($sSpiderTitle, "", $sSpider_File_Class, "!aft")
;~     ControlSend($sSpiderTitle, "", $sSpider_File_Class, "{F}", 1)
;~     ControlSend($sSpiderTitle, "", $sSpider_File_Class, "{T}", 1)
;~ 	Send("{ALT}")
;~ 	Send("{F}")
;~ 	Send("{T}")
	; Wait for 1 second
;~ 	WinWait("","",1)
	; Pastes the copied test design file path
;~     ControlSend($sSpiderTitle, "", $sSpider_File_Class, "^v", 1)
;~     ControlSend($sSpiderTitle, "", $sSpider_File_Class, "{ENTER}", 0)
	ControlSend($sSpiderTitle, "", $sSpider_File_Class, "^v")
	ControlSend($sSpiderTitle, "", $sSpider_File_Class, "{ENTER}")
	;Send("^v")
	;Send("{ENTER}")
;~ 	WinWait("","",5)
	; Save the configuration of the DSpider
    ControlSend($sSpiderTitle, "", $sSpider_File_Class, "^s", 0)
	;Send("^s")
;~ 	WinWait("","",5)

;~     Sleep(2000)
    _JEH_jnknsCreateLogFile("0", "Starting", $sSpider_Log_TxtFile)
EndFunc		;==>_JMI_jnknsInitLog

; #FUNCTION# =========================================================================================================
; Name					:	_JMI_jnknsBuildTree
; Description		:	Get the specified text snippets and associated ClassNameNNs.
; Syntax				:	_JMI_jnknsBuildTree()
; Parameters		:	Const ByRef $TextClasses
; Requirement(s)	:	v3.3 +
; Return values		: 	1 - No Error Occured
;								0 - Error Occured
; Author				:	prdedumo
; Modified			:	None
;=====================================================================================================================
Func _JMI_jnknsBuildTree(Const ByRef $TextClasses)
	; Delete any existing TreeView; this is the easiest way to get rid of all
	; existing window data.
	If $g_iJM_Handles[$HAN_TREE] <> '' Then GUICtrlDelete($g_iJM_Handles[$HAN_TREE])
	; Create a new TreeView.
	$g_iJM_Handles[$HAN_TREE] = GUICtrlCreateTreeView(0, 0, 0, 0, _
			$GUI_SS_DEFAULT_TREEVIEW, $WS_EX_CLIENTEDGE)
	GUICtrlSetResizing($g_iJM_Handles[$HAN_TREE], $GUI_DOCKBORDERS)
	; Populate with text snippets and associated ClassNameNNs.
	For $I = 1 To $TextClasses[0][0]
		Local $TextNode = GUICtrlCreateTreeViewItem( _
				"'" & $TextClasses[$I][0] & "'", $g_iJM_Handles[$HAN_TREE])
		Local $Classes = $TextClasses[$I][1]
		; Check if class contains tprj
		if StringInStr( $TextClasses[$I][0] , 'UnitTestProject.tprj') Then
			$g_iJM_Spider_Software_Path_Class = $Classes
            $g_iJM_Assembly_TPRJ = $TextClasses[$I][0]
		EndIf
		; Checks class
		Switch $TextClasses[$I][0]
			; Class of F5 button
			Case 'テスト開始(F5)'
				$g_iJM_Spider_F5_Class = $Classes
			; Class of File
			Case 'menuStrip1'
				$g_iJM_Spider_File_Class = $Classes
			Case Else
		EndSwitch
		if @error Then
			Return 0
			ExitLoop
		EndIf
	Next
	Return 1
EndFunc		 ;==>_JMI_jnknsBuildTree

; #FUNCTION# =========================================================================================================
; Name					:	_JMI_jnknsWinGetClassesByText
; Description		:	Gets class in reference with the window being selected
; Syntax				:	_JMI_jnknsWinGetClassesByText($Title, $Text = '')
; Parameters		:	$Title	-	Title window of the FS Unit test tool
;								$Text	-	Optional parameter
; Requirement(s)	:	v3.3 +
; Return values		: 	Array of $Texts
; Author				:	prdedumo
; Modified			:	None
; ====================================================================================================================
Func _JMI_jnknsWinGetClassesByText($Title, $Text = '')
	Local $Classes = _JMI_jnknsWinGetControlIDs($Title, $Text)
	Local $Texts[$Classes[0] + 1][2]

    $Texts[0][0] = 0
	For $I = 1 To $Classes[0]
		_JMI_jnknsAddClass($Texts, ControlGetText($Title, $Text, $Classes[$I]), $Classes[$I])
	Next
	Return $Texts
EndFunc		;==>WinGetClassesByText

; #INTERNAL_USE_ONLY#====================================================================================================
; Name				:	_JMI_jnknsWinGetControlIDs
; Description	:	Returns an array of ClassNameNNs for a window where element 0 is a count.
; Author			:	prdedumo
; Remarks			:
; ====================================================================================================================
Func _JMI_jnknsWinGetControlIDs($sTitle, $sText = '')
	Local $avClasses[1], $iCounter, $sClasses, $sClassStub, $sClassStubList

    ; Request an unnumbered class list.
	$sClassStubList = WinGetClassList($sTitle, $sText)
	; Return an empty response if no controls exist.
	; Additionally set @Error if the specified window was not found.
	If $sClassStubList = '' Then
		If @error Then SetError(1)
		$avClasses[0] = 0
		Return $avClasses
	EndIf
	; Prepare an array to hold the numbered classes.
	ReDim $avClasses[StringLen($sClassStubList) - _
			StringLen(StringReplace($sClassStubList, @LF, '')) + 1]
	; The first element will contain a count.
	$avClasses[0] = 0
	; Count each unique class, enumerate them in the array and remove them from the string.
	Do
		$sClassStub = _
				StringLeft($sClassStubList, StringInStr($sClassStubList, @LF))
		$iCounter = 0
		While StringInStr($sClassStubList, $sClassStub)
			$avClasses[0] += 1
			$iCounter += 1
			$avClasses[$avClasses[0]] = _
					StringTrimRight($sClassStub, 1) & $iCounter
			$sClassStubList = _
					StringReplace($sClassStubList, $sClassStub, '', 1)
		WEnd
	Until $sClassStubList = ''
	Return $avClasses
EndFunc		;==>WinGetControlIDs

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_JMI_jnknsAddClass
; Description	:	Adds a class to a text entry in the given text/class list.
;							If the given text is not already contained then a new element is created.
; Author			:	prdedumo
; Remarks			:
; ====================================================================================================================
Func _JMI_jnknsAddClass(ByRef $Texts, $Text, $Class)
	For $I = 1 To $Texts[0][0]
		If $Text == $Texts[$I][0] Then
			$Texts[$I][1] &= @LF & $Class
			Return
		EndIf
	Next
	; This point is reached if the text doesn't already exist in the list.
	$Texts[0][0] += 1
	$Texts[$Texts[0][0]][0] = $Text
	$Texts[$Texts[0][0]][1] = $Class
EndFunc		;==>AddClass

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_JMI_jnknsBrowseFolder
; Description	:	Browse Test Design
; Author			:	prdedumo
; Remarks			:
; ====================================================================================================================
Func _JMI_jnknsBrowseFolder()
	; Create a constant variable in Local scope of the message to display in FileOpenDialog.
    Local Const $sMessage = "Select your Test Design."

    ; Display an open dialog to select a file.
    Local $sFileOpenDialog = FileOpenDialog($sMessage, @WindowsDir & "\", "All (*.*)", $FD_FILEMUSTEXIST)
    If @error Then
        ; Display the error message.
;~         MsgBox($MB_SYSTEMMODAL, "", "No file was selected.")
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)
    Else
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)
        ; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
        $sFileOpenDialog = StringReplace($sFileOpenDialog, "|", @CRLF)
        ; Display the selected file.
		$g_sJMI_TestDesign_File = $sFileOpenDialog
    EndIf
EndFunc		;==>_JMI_jnknsBrowseFolder

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_JMI_jnknsReCheckIfError
; Description	:	Recheck run time if error still occurs
; Author			:	prdedumo
; Remarks			:
; ====================================================================================================================
Func _JMI_jnknsReCheckIfError($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile)
    if FileExists($sUnitTest_Log_TxtFile) Then
		_JEH_jnknsCheckErrHandler($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile)
	Else
		_JEH_jnknsCreateLogFile("0", "OK", $sSpider_Log_TxtFile)
	EndIf
    Return 1
EndFunc     ; ==>_JMI_jnknsReCheckIfError
