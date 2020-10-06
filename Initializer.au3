
#include-once

#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <StringConstants.au3>
#include <Array.au3>
#include <Excel.au3>
#include <WinAPIFiles.au3>
#include <File.au3>

#include <WinAPI.au3>
#Include <Misc.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include <GuiTab.au3>

; #GLOBAL VARIABLES# =====================================================================================================
Global Enum $HAN_GUI, $HAN_TREE, $HAN_BTN, $HAN_BTN2, $HAN_COUNT
Global	$g_iJM_Handles[$HAN_COUNT], _
		$g_iJM_Spider_F5_Class , _				;	Class value depending in the AutoIt v3 Window info
		$g_iJM_Spider_File_Class, _				;	Class value depending in the AutoIt v3 Window info
		$g_iJM_Spider_Software_Path_Class , _	;	Class value depending in the AutoIt v3 Window info
		$g_iJM_Spider_Run_Class, _				;	Class value depending in the AutoIt v3 Window info
		$g_iJM_Assembly_Error, _				;	Class value depending in the AutoIt v3 Window info
		$g_iJM_Assembly_TPRJ

Global $g_iJM_getWin
Global $g_JPE_errlogtxt
; ====================================================================================================================

; #GLOBAL VARIABLES#	=====================================================================================================
Global              $g_sJEH_FixValue
Global              $g_iJEH_FormClose, _
                    $g_iJEH_PLError_Check
Global Const    $STD_INPUT_HANDLE = -10
Global Const    $STD_OUTPUT_HANDLE = -11
Global Const    $STD_ERROR_HANDLE = -12
Global Const    $_CONSOLE_SCREEN_BUFFER_INFO = _
                        "short dwSizeX;" & _
                        "short dwSizeY;" & _
                        "short dwCursorPositionX;" & _
                        "short dwCursorPositionY;" & _
                        "short wAttributes;" & _
                        "short Left;" & _
                        "short Top;" & _
                        "short Right;" & _
                        "short Bottom;" & _
                        "short dwMaximumWindowSizeX;" & _
                        "short dwMaximumWindowSizeY"
Global	$TRUE=1 , _
        $FALSE= 0
; ====================================================================﻿================================================


; #GLOBAL VARIABLES# =====================================================================================================
Global		$g_sJMI_TestDesign_File, _																									;	Test Design File name
				$g_sJMI_Spider_Version, _																									;	Spider Version Name
				$g_sJMI_TPRJ_Path, _																											;	TPRJ path on the build software
				$g_sJMI_Spider_Latest = 'C:\Program Files (x86)\DENSO\D-SPIDER\D-SPIDER.exe', _		;	DSpider exe file path
				$g_sJMI_Spider_Latest_Title = 'ソフトウェア単体テスト自動化ツール D-SPIDER  Ver.1.0.0', _		;	DSpider Title name
				$g_sJMI_Spider_Old = 'C:\Program Files (x86)\DENSO\FSUnitTest\FSUnitTest.exe', _		;	OLD FSUnit exe file path
				$g_sJMI_Spider_Old_Title = '機能安全対応ソフトウェア単体テストツール  ver.0.9.13.1'				;	OLD FSUnit Title Name
Global		$g_iJMI_Spider_win_x = 0, _																								;	DSpider X-position in the monitor
				$g_iJMI_Spider_win_y = 0, _																								;	DSpider Y-position in the monitor
				$g_iJMI_Spider_win_width = 691, _																						;	DSpider resize-width
				$g_iJMI_Spider_win_height = 202, _																					;	DSpider resize-height
                $g_iJM_Spider_Software_Path_Class, _
                $g_iJM_Assembly_TPRJ, _
                $g_iJEH_PLError_Check, _
                $g_iJM_Assembly_Error
; ====================================================================================================================
Global		$g_iJM_Handles[$HAN_COUNT], _
				$g_iJM_Spider_F5_Class , _							;	Class value depending in the AutoIt v3 Window info
				$g_iJM_Spider_File_Class, _							;	Class value depending in the AutoIt v3 Window info
				$g_iJM_Spider_Software_Path_Class , _		;	Class value depending in the AutoIt v3 Window info
				$g_iJM_Spider_Run_Class							;	Class value depending in the AutoIt v3 Window info


; ====================================================================﻿================================================
; #CONSTANT VARIABLES# =====================================================================================================


Global Const $_SMALL_RECT = _
            "short Left;" & _
            "short Top;" & _
            "short Right;" & _
            "short Bottom"
Global Const $_COORD = _
            "short X;" & _
            "short Y"
Global Const $_CHAR_INFO = _
            "wchar UnicodeChar;" & _
            "short Attributes"

; ====================================================================================================================
; #GLOBAL VARIABLES# =====================================================================================================
Global  $g_JPE_pCmd, _                                                                                               ;Current Cmd process
            $g_JPE_hCmd, _                                                                                               ;Current Handle of Cmd Process
            $g_JPE_hCon, _                                                                                                ;Current Console Attachment
            $g_JPE_Cmdtext, _                                                                                            ;Current Text Display on Console Process
            $g_JPE_isErrorValid                                                                                        ;Check as first layer detection if error exist

; ====================================================================================================================


; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_JEH_RefreshSettings
; Description	:	Check output file of makeclean
; Author			:	_JAE_Rebuild_Software ($sSoftwarePath)
; Remarks			:
; ================================================================================================================

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_JEH_RefreshSettings
; Description	:	Check output file of makeclean
; Author			:	_JAE_Rebuild_Software ($sSoftwarePath)
; Remarks			:
; ================================================================================================================
Func _JEH_RefreshSettings($sSoftwarePath)
	Local	$sUnitTestTprjPath, _
				$iFileSize, _
                $i
    Local   $sTextClasses, _
                $sPassClass, _
                $sObjectSourceClass, _
                $sSettingStorageClass, _
                $sTestSheetClass
    Local   $hTab
    Local   $tIndex
    Local   $aTestSheetClass
	Local	$hTestToolHandler

	$hTestToolHandler = WinGetHandle($g_sJMI_Spider_Version)

    Sleep(200)
    ; Activate FSUnit window
    WinActivate($g_sJMI_Spider_Version)
    $sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle($g_sJMI_Spider_Version))
    For $i = 1 To $sTextClasses[0][0]
        If StringInStr($sTextClasses[$i][1], "SysListView32") Then
            $aTestSheetClass = StringSplit($sTextClasses[$i][1], @LF)
            For $j = 1 To $aTestSheetClass[0]
                If StringInStr($aTestSheetClass[$j], "SysListView32") Then
                    $sTestSheetClass = $aTestSheetClass[$j]
                    ExitLoop 2
                EndIf
            Next
        Else
            $sTestSheetClass = $sTextClasses[$i][1]
        EndIf
    next
    Sleep(200)
	; Refresh tprj settings
    $sUnitTestTprjPath = $sSoftwarePath & '\UnitTestProject.tprj'
	ClipPut($sUnitTestTprjPath)
	
	;Send("{ALT}")               ; Send Keys
	;Send("{F}")
	;Send("{O}")
	ControlSend($hTestToolHandler, "", "{ALT}")
	ControlSend($hTestToolHandler, "", "{F}")
	ControlSend($hTestToolHandler, "", "{O}")
	Sleep(1000)                 ; Wait for 1 second
	;Send("^v")                  ; Pastes the copied tprj path
	;Send("{ENTER}")
	ControlSend($hTestToolHandler, "", "^v")
	ControlSend($hTestToolHandler, "", "{ENTER}")
    Sleep(200)

    ; Refresh header files in the パス Tab
    ;Send("{ALT}")
    ;Send("{E}")
    ;Send("{RIGHT}")
    ;Send("{DOWN}")
    ;Send("{P}")
	ControlSend($hTestToolHandler, "", "{ALT}")
	ControlSend($hTestToolHandler, "", "{E}")
	ControlSend($hTestToolHandler, "", "{RIGHT}")
	ControlSend($hTestToolHandler, "", "{DOWN}")
	ControlSend($hTestToolHandler, "", "{P}")

    Sleep(200)
    WinActivate('プロジェクト設定')     ; Title of the header window
    $sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle('プロジェクト設定'))
    If StringInStr($sTextClasses[7][1], @LF) Then
			$sClassTrim = StringLeft($sTextClasses[7][1], StringInStr($sTextClasses[7][1], @LF) - 1)
    EndIf
    $hTab = ControlGetHandle('プロジェクト設定', "",$sClassTrim)
    $tIndex = _GUICtrlTab_FindTab($hTab, 'パス', True, 0)
    If $tIndex = -1 Then
    Else
        _GUICtrlTab_SetCurFocus($hTab, $tIndex)
    EndIf
    WinActivate('プロジェクト設定')
    Sleep(200)
    $sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle('プロジェクト設定'))
    For $i = 1 To $sTextClasses[0][0]
        If $sTextClasses[$i][0] = '自動取得' Then
            If StringInStr($sTextClasses[$i][1], @LF) Then
                $sPassClass = StringLeft($sTextClasses[$i][1], StringInStr($sTextClasses[$i][1], @LF) - 1)
            Else
                $sPassClass = $sTextClasses[$i][1]
            EndIf
        EndIf
    next
    Sleep(200)
    ControlClick('プロジェクト設定',"",$sPassClass)
    ;Send("{ENTER}")
	ControlSend($hTestToolHandler, "", "{ENTER}")

    Sleep(200)
    ; Refresh 対象ソース tab
    WinActivate('プロジェクト設定')
    $hTab = ControlGetHandle('プロジェクト設定', "",$sClassTrim)
    $tIndex = _GUICtrlTab_FindTab($hTab, '対象ソース', True, 0)
    If $tIndex = -1 Then
    Else
        _GUICtrlTab_SetCurFocus($hTab, $tIndex)
    EndIf
    $sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle('プロジェクト設定'))
    For $i = 1 To $sTextClasses[0][0]
        If $sTextClasses[$i][0] = '自動取得' Then
            If StringInStr($sTextClasses[$i][1], @LF) Then
                $sObjectSourceClass = StringRight($sTextClasses[$i][1], StringInStr($sTextClasses[$i][1], @LF) - 1)
            Else
                $sObjectSourceClass = $sTextClasses[$i][1]
            EndIf
        EndIf
        If $sTextClasses[$i][0] = '設定保存' Then
            If StringInStr($sTextClasses[$i][1], @LF) Then
                $sSettingStorageClass = StringRight($sTextClasses[$i][1], StringInStr($sTextClasses[$i][1], @LF) - 1)
            Else
                $sSettingStorageClass = $sTextClasses[$i][1]
            EndIf
        EndIf
    next
    Sleep(200)
    ControlClick('プロジェクト設定',"",$sObjectSourceClass)
	;Send("{ENTER}")
	ControlSend($hTestToolHandler, "", "{ENTER}")
    WinActivate('プロジェクト設定')
    ControlClick('プロジェクト設定',"",$sSettingStorageClass)
    Sleep(200)

    ; Check again the test sheet
    WinActivate($g_sJMI_Spider_Version)
    ;ControlSend($g_sJMI_Spider_Version,"","[NAME:lvwFileList]"," ")
	ControlClick($g_sJMI_Spider_Version,"","[NAME:lvwFileList]")

    ; Save changes
	;Send("^s")
	ControlSend($hTestToolHandler, "", "^s")
	ControlSend($g_sJMI_Spider_Version, "", "[NAME:lvwFileList]", "{space}")
	ControlSend($g_sJMI_Spider_Version, "", "[NAME:lvwFileList]", "{space}")
EndFunc		; ==>_JEH_RefreshSettings

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
			Case '?????(F5)'
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
	; Check if D-Spider Latest UnitTest Tool is already running
	if ProcessExists("D-SPIDER.exe") Then
		WinActivate($g_sJMI_Spider_Latest_Title,"")
		WinSetState($g_sJMI_Spider_Latest_Title,"",@SW_MAXIMIZE)
	Else
		; Run D-Spider, returns error if no D-Spider installed
		$hRunHandler = Run($g_sJMI_Spider_Latest)
		if @error Then
			if ProcessExists("FSUnitTest.exe") Then
				WinActivate($g_sJMI_Spider_Old,"")
				WinSetState($g_sJMI_Spider_Old,"",@SW_MAXIMIZE)
			Else
				$hRunHandler = Run($g_sJMI_Spider_Latest)
				if @error Then
					Exit
				EndIf
				; Wait until FSUnitTest opens
				WinWaitActive($g_sJMI_Spider_Old_Title,"","")
			EndIf
		EndIf
		; Wait until D-Spider Opens
		WinWaitActive($g_sJMI_Spider_Latest_Title,"","")
	endif
	; Checks if Window exists
	if WinExists($g_sJMI_Spider_Latest_Title) Then
		WinMove($g_sJMI_Spider_Latest_Title, "", $g_iJMI_Spider_win_x, $g_iJMI_Spider_win_y, $g_iJMI_Spider_win_width, $g_iJMI_Spider_win_height)
		$g_sJMI_Spider_Version = $g_sJMI_Spider_Latest_Title
		Return 1
	ElseIf WinExists($g_sJMI_Spider_Old_Title) Then
		WinMove($g_sJMI_Spider_Old_Title, "", $g_iJMI_Spider_win_x, $g_iJMI_Spider_win_y, $g_iJMI_Spider_win_width, $g_iJMI_Spider_win_height)
		$g_sJMI_Spider_Version = $g_sJMI_Spider_Old_Title
		Return 1
	Else
		Return 0
	EndIf
endFunc		;==>_JMI_jnknsCallDSpider


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
				$sSpider_Run_Class = "??????", _															;	Class value depending in the AutoIt v3 Window info
				$sSpider_Path = "", _																						;	Setting initial value to null
				$sSpider_Software_Path = "", _																		;	Setting initial value to null
				$sSSpider_Local = "", _																					;	Setting initial value to null
				$sSpider_Log_TxtFile = "", _																				;	Setting initial value to null
				$sUnitTest_Log_TxtFile = ""
	Local	$iReturnF5
	Local	$hTestToolHandler
	$iReturnF5 = 0
;~     $g_iJEH_PLError_Check = 0

	$sSpider_Software_Path = ControlGetText($sSpiderTitle,"",$sSpider_Software_Path_Class)
	$sSpider_Path =  StringTrimRight($sSpider_Software_Path,21)
	$sUnitTest_Log_TxtFile = $sSpider_Path & "\UnitTest\log.txt"
	$sSpider_Log_TxtFile = @ScriptDir & '\Log.txt'
	$hTestToolHandler = WinGetHandle($sSpiderTitle)

	; Copy the test Design File
	ClipPut($g_sJMI_TestDesign_File)
	; Send Keys
	;Send("{ALT}")
	;Send("{F}")
	;Send("{T}")
	ControlSend($sSpiderTitle, "", $sSpider_File_Class, "!aft")
	; Wait for 1 second
	WinWait("","",1)
	; Pastes the copied test design file path
	;Send("^v")
	;Send("{ENTER}")
	ControlSend($hTestToolHandler, "", "^v")
	ControlSend($hTestToolHandler, "", "{ENTER}")
	WinWait("","",5)
	; Save the configuration of the DSpider
	;Send("^s")
	ControlSend($hTestToolHandler, "", "^s")
	WinWait("","",5)
	; Presses F5 in the DSpider Tool
	ControlClick($sSpiderTitle,"",$sSpider_F5_Class)
	WinWait("","",10)
	$sSpider_Local = WinActivate($sSpider_Run_Class)

	; Loop to wait until running of the tool is done
	While 1
		$sSpider_Local = WinActivate($sSpider_Run_Class)
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
    ;   _JMI_jnknsReCheckIfError($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile)
    EndIf
	$iReturnF5 = 1
	Return $iReturnF5
EndFunc		;==>_JMI_jnknsPressF5
