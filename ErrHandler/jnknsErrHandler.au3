
#include-once

#cs	==================================================================================================================
	Title 				:	jnknsMainErrHandler
	AutoIt Version	: 	3.3.14.5
	Language		: 	English
	Description		:	Error Handling script
	Author				: 	prdedumo
    Version            :    0.1
#ce	==================================================================================================================

;~ #AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w- 7

; #INCLUDES#	============================================================================================================
	#include <Array.au3>
	#include <File.au3>
	#include <MsgBoxConstants.au3>
	#include <String.au3>
	#include <FileConstants.au3>
	#include <WinAPIFiles.au3>
    #include <GuiTab.au3>

	#include "jnknsErrorMessage.au3"
    #include "..\jnknsMainInitialize.au3"
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
; ====================================================================﻿================================================

; #CURRENT# ============================================================================================================
;_JEH_jnknsCheckErrHandler
;_JEH_jnknsCreateLogFile﻿
;_JEH_FindInFile
;_JEH_BackUpFile
;_JEH_EditFile
;_JEH_Rebuild_Software
;_JEH_Check_OutPutFile
; ====================================================================﻿================================================

; #FUNCTION# =========================================================================================================
; Name					:	_JEH_jnknsCheckErrHandler
; Description		:	Main process of identifying error depending on the error log of the DSpider Tool
; Syntax				:	_JEH_jnknsCheckErrHandler($unitTest_ERRlog_txt, $spider_ERRlog_txt)
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:	prdedumo
; Modified			:	None
;=====================================================================================================================
Func _JEH_jnknsCheckErrHandler($unitTest_ERRlog_txt, $spider_ERRlog_txt)
	Local	$iLine = 0, _
				$iValid = 0
	Local	$sLine, _
				$sFileRead
	Local	$aArray
	Local	$hFileOpen

    If $g_iJEH_PLError_Check = 1 Then
        _JEH_jnknsCreateLogFile("2", "Pending", $spider_ERRlog_txt)
    Else
        ; Open the file for reading and store the handle to a variable.
        $hFileOpen = FileOpen($unitTest_ERRlog_txt, $FO_READ)
        If $hFileOpen = -1 Then
            MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
            Return False
        EndIf
        ; Read the contents of the file using the handle returned by FileOpen.
        $sFileRead = FileRead($hFileOpen)
        if Not _FileReadToArray($unitTest_ERRlog_txt,$sLine) Then
            MsgBox(4096,"Error", " Error reading log to Array error:" & @error)
            Exit
        EndIf
        ; Looping all lines in the log txt
        for $iLine = 1 to $sLine[0]
            $aArray = _StringBetween($sLine[$iLine], ";%SYSTEMG-E-EVAL, ", " ")
            ; First Condition if Error found
            If  StringInStr($sLine[$iLine], '(Error)') and Not $iValid Then
                ; Condition if error fix is rev2.0
                If StringInStr($sLine[$iLine], $ERR_MESSAGE_COMMENT_RESULT1) and Not $iValid Then
                    $g_sJEH_FixValue = "Comment_カバレッジ結果"
                    $iValid = 1
                    _JEH_jnknsCreateLogFile("3", "Pending", $spider_ERRlog_txt)
                    ExitLoop
                EndIf
                ; Condition if error fix is rev3.0
                If StringInStr($sLine[$iLine], $ERR_MESSAGE_COMMENT_RESULT2) and Not $iValid Then
                    If StringInStr($sLine[$iLine], $ERR_MESSAGE_COMMENT_RESULT2) and Not $iValid Then
                        $g_sJEH_FixValue = "Comment_Result"
                        $iValid = 1
                        _JEH_jnknsCreateLogFile("3", "Pending", $spider_ERRlog_txt)
                        ExitLoop
                    EndIf
                EndIf
            ; Condition if error is ambiguous
            ElseIf StringInStr($sLine[$iLine], $ERR_MESSAGE_COMMENT_RESULT3) and Not $iValid Then
                If UBound($aArray) > 0 and Not $iValid Then
                    $aArray = _StringBetween($sLine[$iLine], '"', '"')
                    $g_sJEH_FixValue = $aArray[0]
                    $iValid = 1
                    _JEH_jnknsCreateLogFile("4", "Pending", $spider_ERRlog_txt)
                EndIf
                ExitLoop
            Else
            EndIf
        Next
    EndIf
	; Close log txt
	FileClose($hFileOpen)
EndFunc

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_JEH_jnknsCreateLogFile
; Description	:	Creates/Writes Log File in the Script Directory
; Author			:	prdedumo
; Remarks			:
; ====================================================================================================================
Func _JEH_jnknsCreateLogFile($logText, $logStatus, $logTextFile)
    Local	$aArray[7]
	Local	$sErrlogFile, _
				$sString = ""
	Local	$hFileOpen
	;	Default Log Path
	$sErrlogFile = $logTextFile ; @ScriptDir & '\Log.txt'
	; Open the file for reading and store the handle to a variable.
	$hFileOpen = FileOpen($sErrlogFile, 2)
	If $hFileOpen = -1 Then
		$g_iJEH_FormClose = 0
        Return False
    EndIf
	;	Default Entry
	$aArray[0] = "Tprj_Path:	" & $g_sJMI_TPRJ_Path
	$aArray[1] = "TestSheet_Filename:	" & $g_sJMI_TestDesign_File
	$aArray[2] = "Error_Number:	" & $logText
	$aArray[3] = "Status:	" & $logStatus
	$aArray[4] = "Spider Version:	" & $g_sJMI_Spider_Version
	$aArray[5] = "Spider Log Text:	" & $sErrlogFile
	$aArray[6] = "Error##_Fix:	" & $g_sJEH_FixValue

	For $vElement In $aArray
        $sString = $sString & $vElement & @CRLF
    Next
	; Write log txt
	FileWrite($hFileOpen,$sString)
	; Close log txt
	FileClose($hFileOpen)
	; Updates global variable to 1
	$g_iJEH_FormClose = 1
EndFunc

; #INTERNAL_USE_ONLY# ================================================================================================
; Name					:	_JEH_FindInFile
; Description		:	Search for a string within files located in a specific directory.
; Syntax				:	_JEH_FindInFile($sSearch, $sFilePath[, $sMask = '*'[, $fRecursive = True[, $fLiteral = Default[,
;                  				$fCaseSensitive = Default[, $fDetail = Default]]]]])
; Parameters		:	$sSearch					- The keyword to search for.
;                  				$sFilePath				- The folder location of where to search.
;                  				$sMask					- [optional] A list of filetype extensions separated with ';' e.g. '*.au3;*.txt'. Default is all files.
;                  				$fRecursive				- [optional] Search within subfolders. Default is True.
;                  				$fLiteral            		- [optional] Use the string as a literal search string. Default is False.
;                  				$fCaseSensitive		- [optional] Use Search is case-sensitive searching. Default is False.
;                  				$fDetail					- [optional] Show filenames only. Default is False.
; Requirement(s)	:	v3.3 +
; Return values		:	Success - Returns a one-dimensional and is made up as follows:
;                            		$aArray[0] = Number of rows
;                            		$aArray[1] = 1st file
;                            		$aArray[n] = nth file
;                  				Failure - Returns an empty array and sets @error to non-zero
; ====================================================================================================================
Func _JEH_FindInFile($sSearch, $sFilePath, $sMask = '*', $fRecursive = True, $fLiteral = Default, $fCaseSensitive = True, $fDetail = Default)
    Local $sCaseSensitive = $fCaseSensitive ? '' : '/i', $sDetail = $fDetail ? '/n' : '/m', $sRecursive = ($fRecursive Or $fRecursive = Default) ? '/s' : ''
    If $fLiteral Then
        $sSearch = ' /c:' & $sSearch
    EndIf
    If $sMask = Default Then
        $sMask = '*'
    EndIf
    $sFilePath = StringRegExpReplace($sFilePath, '[\\/]+$', '') & '\'
    Local Const $aMask = StringSplit($sMask, ';')
    Local $iPID = 0, $sOutput = ''
    For $i = 0 To $aMask[0]
        $iPID = Run(@ComSpec & ' /c ' & 'findstr ' & $sCaseSensitive & ' ' & $sDetail & ' ' & $sRecursive & ' "' & $sSearch & '" "' & $sFilePath & $aMask[$i] & '"', @SystemDir, @SW_HIDE, $STDOUT_CHILD)
        ProcessWaitClose($iPID)
        $sOutput &= StdoutRead($iPID)
    Next
    Return StringSplit(StringStripWS(StringStripCR($sOutput), BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING)), @LF)
EndFunc   	;==>_JEH_FindInFile

; #INTERNAL_USE_ONLY#====================================================================================================
; Name				:	_JEH_BackUpFile
; Description	:	Backup File if ambiguous was found
; Author			:	prdedumo
; Remarks			:
; ====================================================================================================================
Func _JEH_BackUpFile ($sFileName, $sSoftwarePath, $sPathSplitFileName, $sPathExtension)
	; Create a constant variable in Local scope of the filepath that will be read/written to.
    Local Const $sFilePath = _WinAPI_GetTempFileName($sSoftwarePath)
	Local $iResult
	$iResult = 0
	; Create a temporary file to copy.
    If Not FileWrite($sFilePath, "Creating temporary file: " & $sFilePath) Then
		Return False
		Exit
    EndIf
	; Copy the file in the same path
	$iResult = FileCopy($sFileName, $sSoftwarePath & '\' & $sPathSplitFileName & '_BAK' &$sPathExtension , 9)
	; Delete the temporary file.
    FileDelete($sFilePath)
	Return $iResult
EndFunc		;==>_JEH_BackUpFile

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_JEH_EditFile
; Description	:	Backup File if ambiguous was found
; Author			:	prdedumo
; Remarks			:
; ====================================================================================================================
Func _JEH_EditFile ($sFileName, $sFix_Value)
	Local	$iResult, _
				$iRet
	Local	$sFileAttrib

	$iRet = 0
	; Checking file if readonly
	$sFileAttrib = FileGetAttrib($sFileName)
	If $sFileAttrib = 'RA' Then
		FileSetAttrib($sFileName, "-R")
	EndIf
	; Copy the file in the same path
	$iResult = _ReplaceStringInFile($sFileName, $sFix_Value, $sFix_Value & '_1')
	If $iResult = -1 Then
		Exit
	Else
		$iRet = 1
	EndIf
	Return $iRet
EndFunc		;==>_JEH_EditFile

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_JEH_SetAttrib
; Description	:	Release Read-Only
; Author			:	prdedumo
; Remarks			:
; ====================================================================================================================
Func _JEH_SetAttrib($sCPUFileName)
    Local    $sFileAttrib
    $sFileAttrib = FileGetAttrib($sCPUFileName)
    If $sFileAttrib = 'RA' Then
        FileSetAttrib($sCPUFileName, "-R")
    EndIf
EndFunc ;==>_JEH_SetAttrib

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_JEH_Rebuild_Software
; Description	:	Rebuild software after editing
; Author			:	prdedumo
; Remarks			:
; ====================================================================================================================
Func _JEH_Rebuild_Software ($sSoftwarePath)
	Local	$hFileOpen
	Local	$sFileRead, _
				$sSetSource, _
				$sNewClass, _
				$sMakeBat, _
				$sMakeCommand, _
				$sTextClasses

	; Open cygwin process and check one instance only
	if ProcessExists("mintty.exe") Then
		ProcessClose("mintty.exe")
	endif
    Sleep(200)
	Local $iMinty = Run("C:\cygwin\bin\mintty.exe -i /Cygwin-Terminal.ico -", "", @SW_SHOW, 0x8)
	; Wait 5 seconds for the Cygwin window to appear.
	WinWaitActive("~","",5)
	; Sleep for 2 seconds.
	Sleep(2000)
	; Loop to check if cygwin terminal is ready
	While 1
		; if -sh is not error cygwin is still opening
		Local $hWnd = WinGetHandle("-sh")
		If @error Then
			$hWnd = WinGetHandle("~")
		Else
			$hWnd = 0
		EndIf
		; If true, cygwin is ready to be written
		if $hWnd <> '0x00000000' Then
			ExitLoop
		EndIf
        Sleep(1000)
	WEnd
	; Start of make clean process
	; ========================================================
	; Change Directory
	$sSoftwarePath = StringReplace($sSoftwarePath,"\", "/")
	ClipPut($sSoftwarePath)
	WinActivate("[CLASS:mintty]", "")
	Sleep(3000)																								; Sleep for 3 seconds
	Send('cd ' & $sSoftwarePath & "{ENTER}" )											; Change directory
	$sNewClass = WinGetTitle("[ACTIVE]")													; Get New Class after Change Directory
	Sleep(3000)																								; Sleep for 3 seconds
	$sSoftwarePath = StringReplace($sSoftwarePath,"/", "\")						; Change software path to access the makeFile
	$hFileOpen = FileOpen($sSoftwarePath & "\makefile", $FO_READ)		; Open Makefile
	If $hFileOpen = -1 Then
        Return False
    EndIf
	$sFileRead = FileReadLine($hFileOpen, 12)
	$sSetSource = _StringBetween($sFileRead, "(", ")", $STR_ENDNOTSTART)
	Send('source ' & $sSetSource[0] & '.sh' & "{ENTER}")
	Sleep(2000)		; Sleep for 2 seconds
	Send('make clean  > makeClean_output_and_error.txt 2>&1' & "{ENTER}")
	; Loop to check if make clean is done
	While (1)
		If _JEH_Check_OutPutFile($sSoftwarePath & "\makeClean_output_and_error.txt") Then
			ExitLoop
		EndIf
        Sleep(1000)
	WEnd
	; End of make clean process
	Sleep(2000)
	; Start of make processmake clean  > output_and_
	; ========================================================
	; Check make.bat setting if factory or not
	; Loop to make.bat to get make factory or make only command
	_FileReadToArray($sSoftwarePath & "\make実行.bat", $sMakeBat)
	For $i = 1 To UBound($sMakeBat) -1
		If StringInStr($sMakeBat[$i], 'make') Then
			$sMakeCommand = $sMakeBat[$i]
			ExitLoop
		EndIf
	Next
	WinActivate("[CLASS:mintty]", "")
	Send($sMakeCommand & ' > makeBuild_output_and_error.txt 2>&1' & "{ENTER}" )
	; Loop to check if building is done
	While (1)
		If _JEH_Check_OutPutFile($sSoftwarePath & "\makeBuild_output_and_error.txt") Then
			ExitLoop
		EndIf
        Sleep(1000)
	WEnd
	; End of build process
    Sleep(2000)
    Return 1
EndFunc		;==>_JEH_Rebuild_Software

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_JEH_Check_OutPutFile
; Description	:	Check output file of makeclean
; Author			:	_JAE_Rebuild_Software ($sSoftwarePath)
; Remarks			:
; ================================================================================================================
Func _JEH_Check_OutPutFile($sOutputFile)
	Local	$iRetOutputFile, _
				$iFileSize, _
                $iSize, _
                $iGetFileSize
    Local   $handle

    $iSize = 0
	$iRetOutputFile = 0
	While (1)
        $iGetFileSize = FileGetSize($sOutputFile)
        $iFileSize = FileReadToArray($sOutputFile)    ;  Array to store all lines in the file found
        $iLineCountB = @extended
        ; Sleep for 30 seconds
        Sleep(30000)
		; FileSize will be equal to size if done writing
		If UBound($iFileSize) = $iSize And $iGetFileSize = FileGetSize($sOutputFile) Then
			$iRetOutputFile = 1
			ExitLoop
		EndIf
        $iSize = $iLineCountB
	WEnd
	Return $iRetOutputFile
EndFunc		; ==>_JEH_Check_OutPutFile

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_JEH_RefreshSettings
; Description	:	Check output file of makeclean
; Author			:	_JAE_Rebuild_Software ($sSoftwarePath)
; Remarks			:
; ================================================================================================================
Func _JEH_RefreshSettings($sSoftwarePath, $sStartUpAddress = "", $sComment_Result = "")
	Local	$sUnitTestTprjPath, _
				$iFileSize, _
                $i
    Local   $sTextClasses, _
                $sPassClass, _
                $sObjectSourceClass, _
                $sSettingStorageClass, _
                $sTestSheetClass, _
                $sCapturedTitle
    Local   $hTab
    Local   $tIndex
    Local   $aTestSheetClass
    Local   $clickClass
    Local   $hTestToolHandler

    Sleep(200)
    ; Activate FSUnit window
    ;WinActivate($g_sJMI_Spider_Version)
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

    ; Get the TPRJ from the log.txt file
	Local $toolHwnd = WinGetHandle($g_sJMI_Spider_Version)
    Sleep(2000)

    $sUnitTestTprjPath = $sSoftwarePath & 'UnitTestProject.tprj'
    ClipPut($sUnitTestTprjPath)

    ControlFocus($toolHwnd, "", $sTextClasses[3][1])    ; Set focus on the status bar class
    ControlSend($toolHwnd, "", "", "!fo")               ; Click alt + f + o
    Sleep(2000)                                         ; Delay to get the class of the pop-up item

	Local $tprjHwnd = WinGetHandle("プロジェクトファイルを選択")
	$sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle("プロジェクトファイルを選択"))

	ControlFocus($tprjHwnd, "", $sTextClasses[8][0])   ; Set focus on the tprj input bar
    ControlSend($tprjHwnd, "", "", "^v")
    ControlSend($tprjHwnd, "", "", "{ENTER}")
	  ControlSend($toolHwnd,"","","^s")

    Sleep(2000)

    ; Refresh header files in the パス Tab
	$toolHwnd = WinGetHandle($g_sJMI_Spider_Version)
	Sleep(2000)
    ;ControlFocus($toolHwnd, "", $sTextClasses[3][1])    ; Set focus on the status bar class
    ControlSend($toolHwnd, "", "", "!ae{RIGHT}{DOWN}{P}")
    ;ControlSend($toolHwnd, "", "", "{RIGHT}")
    ;ControlSend($toolHwnd, "", "", "{DOWN}")
    ;ControlSend($toolHwnd, "", "", "{P}")

    Sleep(2000)                                         ; Delay to get the class of the pop-up item

    Local $popHwnd = WinGetHandle("プロジェクト設定")
    $sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle("プロジェクト設定"))
	ControlFocus($popHwnd, "", $sTextClasses[7][1])

    If StringInStr($sTextClasses[7][1], @LF) Then
        $sClassTrim = StringLeft($sTextClasses[7][1], StringInStr($sTextClasses[7][1], @LF) - 1)
    EndIf
    $hTab = ControlGetHandle('プロジェクト設定', "",$sClassTrim)
    $tIndex = _GUICtrlTab_FindTab($hTab, 'パス', True, 0)
    If $tIndex = -1 Then
    Else
        _GUICtrlTab_SetCurFocus($hTab, $tIndex)
    EndIf

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
    ControlSend('プロジェクト設定', "", "", "{ENTER}")

    Sleep(200)

    ; Refresh 対象ソース tab
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
    ControlSend('プロジェクト設定', "", "", "{ENTER}")
    ControlClick('プロジェクト設定',"",$sSettingStorageClass)
    Sleep(200)

    If  $sStartUpAddress <> "" Then
        Sleep(200)
        ; Refresh StartUp tab
        WinActivate('プロジェクト設定')
        $hTab = ControlGetHandle('プロジェクト設定', "",$sClassTrim)
        $tIndex = _GUICtrlTab_FindTab($hTab, '実行時設定', True, 0)
        If $tIndex = -1 Then
        Else
            _GUICtrlTab_SetCurFocus($hTab, $tIndex)
        EndIf
        $sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle('プロジェクト設定'))
        For $i = 1 To $sTextClasses[0][0]
            If  StringInStr($sTextClasses[$i][0],"0x") Then
                $sTextClasses[$i][1] = $sStartUpAddress
            EndIf
        next
	 EndIf

   Sleep(2000)

   If $sComment_Result <> "" Then
        ; Get Cmment_result From TestSheet

		ControlSend($toolHwnd, "", "", "!ae{RIGHT}{DOWN}{A}")

        Sleep(2000)
        $sCapturedTitle = WinGetHandle("アプリケーション設定")
        $sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle("アプリケーション設定"))
            ClipPut($sComment_Result)
            Sleep(200)
			ControlSend($sCapturedTitle, "", "", "{TAB}{TAB}{TAB}{TAB}{TAB}")
			ControlSend($sCapturedTitle, "", "", "{DELETE}^v")

        For $i = 1 To $sTextClasses[0][0]
            If  StringInStr($sTextClasses[$i][0],"設定保存") Then
                $clickClass = $sTextClasses[$i][1]
            EndIf
        next
        ControlClick($sCapturedTitle,"",$clickClass)
	 EndIf

	 ControlSend($toolHwnd,"","","^s")
	 Sleep(2000)

    ; Get the Testsheet from the log.txt file
	$toolHwnd = WinGetHandle($g_sJMI_Spider_Version)
    Sleep(2000)

    ClipPut($g_sJMI_TestDesign_File)

    ControlFocus($toolHwnd, "", $sTextClasses[3][1])    ; Set focus on the status bar class
    ControlSend($toolHwnd, "", "", "!ft")               ; Click alt + f + o
    Sleep(2000)                                         ; Delay to get the class of the pop-up item

	$tprjHwnd = WinGetHandle("テストブックを選択")
	$sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle("テストブックを選択"))

	ControlFocus($tprjHwnd, "", $sTextClasses[8][0])   ; Set focus on the tprj input bar
    ControlSend($tprjHwnd, "", "", "^v")
    ControlSend($tprjHwnd, "", "", "{ENTER}")
	Sleep(200)
    ControlSend($toolHwnd,"","","^s")
    Sleep(2000)

EndFunc		; ==>_JEH_RefreshSettings

; #FUNCTION# ===========================================================================================================
; Name					:	_JPE_jnknsErrorLogger
; Description		:	Get information string from current cmd window
; Syntax				:	_JPE_jnknsGetCmdData()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:   rdbayanado
; Modified			:	None
;=====================================================================================================================
Func _JPE_jnknsErrorLogger($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile)
    Local $aProcessList

    Sleep(10000)
    If ProcessExists("cmd.exe") Then
        AutoItSetOption("WinTitleMatchMode",4)
        $aProcessList = ProcessList("cmd.exe")
        _JPE_jnknsProcessExist($aProcessList, $sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile)
        ProcessWait("cmd.exe")
        Sleep(2000)
;~     Else
;~         $g_iJEH_PLError_Check = 0
    EndIf

Return $g_iJEH_PLError_Check
EndFunc     ; ==>_JPE_jnknsErrorLogger

; #FUNCTION# ===========================================================================================================
; Name					:	_JPE_jnknsProcessExist
; Description		:	Check if cmd process exist
; Syntax				:	_JPE_jnknsProcessExist($aProcessList)
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:   rdbayanado
; Modified			:	None
;=====================================================================================================================
Func _JPE_jnknsProcessExist($aProcessList, $sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile)

    Local $i
    Local $g_JPE_Cmdtext
    If $aProcessList[0][0] = 1 Then
        Send("{Enter}")
        $g_JPE_isErrorValid = 0
        Return 0
    ElseIf $aProcessList[0][0] <> 0 Then
         For $i = 1 To $aProcessList[0][0]
            If StringInStr($aProcessList[$i][0], "-java") Then
            Else
                While (1)
                    $g_JPE_Cmdtext =_JPE_jnkns_cmdGetText( $aProcessList[$i][0], True )
                        Sleep(15000)
                        If StringInStr($g_JPE_Cmdtext,"doesn't fit in memory block retention_memory ") Then
;~                             MsgBox(0,"","$sUnitTest_Log_TxtFile: " & $sUnitTest_Log_TxtFile & "     $sSpider_Log_TxtFile: " & $sSpider_Log_TxtFile)
                            Sleep(200)
                            _JEH_jnknsCheckErrHandler($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile)
                            $g_iJEH_PLError_Check = 1
                             ProcessClose($aProcessList[$i][1])
                            Return 1
                            ExitLoop
                        Else
                            WinActivate($aProcessList[$i][0])
                            Send("{Enter}")
                            $g_JPE_isErrorValid = 1
                            Return 0
                            ExitLoop
                        EndIf
                WEnd
            EndIf
        Next
    Else
        Return 0
        $g_JPE_isErrorValid = 0
    EndIf
;~     if($aProcessList[0][0] <> 0) Then
;~         Send("{Enter}")
;~         $g_JPE_isErrorValid = 1
;~         Return 1
;~     Else
;~         Return 0
;~         $g_JPE_isErrorValid = 0
;~     EndIf
EndFunc     ; ==>_JPE_jnknsProcessExist

; #FUNCTION# ===========================================================================================================
; Name					:	_JPE_jnkns_cmdAttachConsole
; Description		:	Get current window handle of cmd process
; Syntax				:	_JPE_jnkns_cmdAttachConsole()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:	None
; Modified			:	None
;=====================================================================================================================
Func _JPE_jnkns_cmdAttachConsole($nPid)
    ; Try to attach to the console of the PID.
    Local $aRet = DllCall("kernel32.dll", "int", "AttachConsole", "dword", $nPid)
    If @error Then Return SetError(@error, @extended, False)
    If $aRet[0] Then
        Local $vHandle[2]
        $vHandle[0] = _CmdGetStdHandle($STD_OUTPUT_HANDLE)  ; STDOUT Handle
        $vHandle[1] = DllStructCreate($_CONSOLE_SCREEN_BUFFER_INFO) ; Screen Buffer structure

        ; Return the handle on success.
        Return $vHandle
    EndIf
    Return 0
EndFunc ; _CmdAttachConsole()

Func _CmdGetStdHandle($nHandle)
    Local $aRet = DllCall("kernel32.dll", "hwnd", "GetStdHandle", "dword", $nHandle)
    If @error Then Return SetError(@error, @extended, $INVALID_HANDLE_VALUE)
    Return $aRet[0]
EndFunc ; _CmdGetStdHandle()

; #FUNCTION# ===========================================================================================================
; Name					:	_JPE_jnkns_cmdGetText
; Description		:	Get current window handle of cmd process
; Syntax				:	__JPE_jnkns_cmdGetText()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:	None
; Modified			:	None
;=====================================================================================================================
Func _JPE_jnkns_cmdGetText(ByRef $vHandle, $bAll = False)
    ; Basic sanity check to validate the handle.
    If UBound($vHandle) = 2 Then
        ; Create some variables for convenience.
        Local Const $hStdOut = $vHandle[0]
        Local Const $pConsoleScreenBufferInfo = $vHandle[1]

        ; Try to get the screen buffer information.
        If _GetConsoleScreenBufferInfo($hStdOut, $pConsoleScreenBufferInfo) Then
            ; Load the SMALL_RECT with the projected text position.
            Local $iLeft = DllStructGetData( $pConsoleScreenBufferInfo, "Left")
            Local $iRight = DllStructGetData( $pConsoleScreenBufferInfo, "Right")
            Local $iTop = DllStructGetData( $pConsoleScreenBufferInfo, "Top")
            Local $iBottom = DllStructGetData( $pConsoleScreenBufferInfo, "Bottom")

            Local $iWidth = $iRight - $iLeft + 1

            ; Set up the coordinate structures.
            Local $coordBufferCoord = _CmdWinAPI_MakeDWord(0, 0)
            Local $coordBufferSize = _CmdWinAPI_MakeDWord($iWidth, 1)

            Local $pBuffer = DllStructCreate("dword[" & $iWidth & "]")

            Local Const $pRect = DllStructCreate($_SMALL_RECT)
            ; This variable holds the output string.
            Local $sText = ""
            For $j = _IIf( $bAll, 0, $iTop ) To $iBottom
                Local $sLine = ""
                DllStructSetData( $pRect, "Left", $iLeft )
                DllStructSetData( $pRect, "Right", $iRight )
                DllStructSetData( $pRect, "Top", $j )
                DllStructSetData( $pRect, "Bottom", $j )

                ; Read the console output.
                If _CmdReadConsoleOutput($hStdOut, $pBuffer, $coordBufferSize, $coordBufferCoord, $pRect) Then
                    Local $pPtr = DllStructGetPtr($pBuffer)

                    For $i = 0 To $iWidth - 1
                        ; We offset the buffer each iteration by 4 bytes because that is the size of the CHAR_INFO
                        ; structure.  We do this so we can read each individual character.
                        Local $pCharInfo = DllStructCreate($_CHAR_INFO, $pPtr)
                        $pPtr += 4
                        ; Append the character.
                        $sLine &= DllStructGetData($pCharInfo, "UnicodeChar")
                    Next
                    $sText &= StringStripWS( $sLine, 2 ) & @CRLF
                EndIf
            Next
            $sText = StringStripWS( $sText, 2 )
            Return $sText
        EndIf
        Return SetError( 2, 0, "" )
    EndIf
    Return SetError( 1, 0, "" )
EndFunc   ;==>_CmdGetText

Func _CmdWinAPI_MakeDWord($LoWORD, $HiWORD)
    Local $tDWord = DllStructCreate("dword")
    Local $tWords = DllStructCreate("word;word", DllStructGetPtr($tDWord))
    DllStructSetData($tWords, 1, $LoWORD)
    DllStructSetData($tWords, 2, $HiWORD)
    Return DllStructGetData($tDWord, 1)
EndFunc   ;==>_CmdWinAPI_MakeDWord

Func _GetConsoleScreenBufferInfo($hConsoleOutput, $pConsoleScreenBufferInfo)
    Local $aRet = DllCall("kernel32.dll", "int", "GetConsoleScreenBufferInfo", "hwnd", $hConsoleOutput, _
        "ptr", _CmdSafeGetPtr($pConsoleScreenBufferInfo))
    If @error Then Return SetError(@error, @extended, False)
    Return $aRet[0]
EndFunc ; _GetConsoleScreenBufferInfo()

Func _CmdReadConsoleOutput($hConsoleOutput, $pBuffer, $coordBufferSize, $coordBufferCoord, $pRect)

    Local $aRet = DllCall("kernel32.dll", "int", "ReadConsoleOutputW", "ptr", $hConsoleOutput, _
        "ptr", _CmdSafeGetPtr($pBuffer), "int", $coordBufferSize, "int", $coordBufferCoord, _
        "ptr", _CmdSafeGetPtr($pRect))
    If @error Then SetError(@error, @extended, False)
    Return $aRet[0]
EndFunc ; _CmdReadConsoleOutput()

Func _CmdSafeGetPtr(Const ByRef $ptr)
    Local $_ptr = DllStructGetPtr($ptr)
    If @error Then $_ptr = $ptr
    Return $_ptr
EndFunc ; _CmdSafeGetPtr()

Func _Iif($fTest, $vTrueVal, $vFalseVal)
    If $fTest Then
        Return $vTrueVal
    Else
        Return $vFalseVal
    EndIf
EndFunc   ;==>_Iif

; #FUNCTION# ;=================================================================================
; Function Name ...: _FO_FileSearch (__FO_FileSearchType, __FO_FileSearchMask, __FO_FileSearchAll)
; AutoIt Version ....: 3.3.2.0+ , versions below this @extended should be replaced by of StringInStr(FileGetAttrib($sPath&'\'&$sFile), "D")
; Description ........: Search files by mask in subdirectories.
; Syntax................: _FO_FileSearch ( $sPath [, $sMask = '*' [, $fInclude=True [, $iDepth=125 [, $iFull=1 [, $iArray=1 [, $iTypeMask=1 [, $sLocale=0[, $vExcludeFolders = ''[, $iExcludeDepth = -1]]]]]]]]] )
; Parameters:
;		$sPath - Search path
;		$sMask - Allowed two options for the mask: using symbols "*" and "?" with the separator "|", or a list of extensions with the separator "|"
;		$fInclude - (True / False) Invert the mask, that is excluded from the search for these types of files
;		$iDepth - (0-125) Nesting level (0 - root directory)
;		$iFull - (0,1,2,3)
;                  |0 - Relative
;                  |1 - Full path
;                  |2 - File names with extension
;                  |3 - File names without extension
;		$iArray - if the value other than zero, the result is an array (by default ),
;                  |0 - A list of paths separated by @CRLF
;                  |1 - Array, where $iArray[0]=number of files ( by default)
;                  |2 - Array, where $iArray[0] contains the first file
;		$iTypeMask - (0,1,2) defines the format mask
;                  |0 - Auto detect
;                  |1 - Forced mask, for example *.is?|s*.cp* (it is possible to specify a file name with no characters * or ? and no extension will be found)
;                  |2 - Forced mask, for example tmp|bak|gid (that is, only files with the specified extension)
;		$sLocale - Case sensitive.
;                  |-1 - Not case sensitive (only for 'A-z').
;                  |0 - Not case sensitive, by default. (for any characters)
;                  |1 - Case sensitive (for any characters)
;                  |<symbols> - not case sensitive, specified range of characters from local languages. For example 'À-ÿ¨¸'. 'A-z' is not required, they are enabled by default.
;		$vExcludeFolders - Excludes folders from search. List the folder names via the "|", for example, "Name1|Name2|Name3|".
;		$iExcludeDepth - Nesting level for the parameter $vExcludeFolders. -1 by default, which means disabled.
; Return values ....: Success - Array or a list of paths separated by @CRLF
;					Failure - Empty string, @error:
;                  |0 - No error
;                  |1 - Invalid path
;                  |2 - Invalid mask
;                  |3 - Not found
; Author(s) ..........: AZJIO
; Remarks ..........: Use function _CorrectMask if it is required correct mask, which is entered by user
; ============================================================================================
Func _FO_FileSearch($sPath, $sMask = '*', $fInclude = True, $iDepth = 125, $iFull = 1, $iArray = 1, $iTypeMask = 1, $sLocale = 0, $vExcludeFolders = '', $iExcludeDepth = -1)
	Local $vFileList
	If $sMask = '|' Then Return SetError(2, 0, '')
	; If Not StringRegExp($sPath, '(?i)^[a-z]:[^/:*?"<>|]*$') Or StringInStr($sPath, '\\') Then Return SetError(1, 0, '')
	If Not FileExists($sPath) Then Return SetError(1, 0, '')
	If StringRight($sPath, 1) <> '\' Then $sPath &= '\'

	If $vExcludeFolders Then
		$vExcludeFolders = StringSplit($vExcludeFolders, '|')
	Else
		Dim $vExcludeFolders[1] = [0]
	EndIf

	If $sMask = '*' Or $sMask = '' Then
		__FO_FileSearchAll($vFileList, $sPath, $iDepth, $vExcludeFolders, $iExcludeDepth)
	Else
		Switch $iTypeMask
			Case 0
				If StringInStr($sMask, '*') Or StringInStr($sMask, '?') Or StringInStr($sMask, '.') Then
					__FO_GetListMask($sPath, $sMask, $fInclude, $iDepth, $vFileList, $sLocale, $vExcludeFolders, $iExcludeDepth)
				Else
					__FO_FileSearchType($vFileList, $sPath, '|' & $sMask & '|', $fInclude, $iDepth, $vExcludeFolders, $iExcludeDepth)
				EndIf
			Case 1
				__FO_GetListMask($sPath, $sMask, $fInclude, $iDepth, $vFileList, $sLocale, $vExcludeFolders, $iExcludeDepth)
			Case Else
				If StringInStr($sMask, '*') Or StringInStr($sMask, '?') Or StringInStr($sMask, '.') Then Return SetError(2, 0, '')
				__FO_FileSearchType($vFileList, $sPath, '|' & $sMask & '|', $fInclude, $iDepth, $vExcludeFolders, $iExcludeDepth)
		EndSwitch
	EndIf

	If Not $vFileList Then Return SetError(3, 0, '')
	Switch $iFull
		Case 0
			$vFileList = StringRegExpReplace($vFileList, '(?m)^[^\v]{' & StringLen($sPath) & '}', '')
		Case 2
			$vFileList = StringRegExpReplace($vFileList, '(?m)^.*\\', '')
		Case 3
			$vFileList = StringRegExpReplace($vFileList, '(?m)^[^\v]+\\', '')
			$vFileList = StringRegExpReplace($vFileList, '(?m)\.[^./:*?"<>|\\\v]+\r?$', @CR)
	EndSwitch
	$vFileList = StringTrimRight($vFileList, 2)
	Switch $iArray
		Case 1
			$vFileList = StringSplit($vFileList, @CRLF, 1)
		Case 2
			$vFileList = StringSplit($vFileList, @CRLF, 3)
	EndSwitch
	Return $vFileList
EndFunc   ;==>_FO_FileSearch

Func __FO_GetListMask($sPath, $sMask, $fInclude, $iDepth, ByRef $sFileList, $sLocale, ByRef $aExcludeFolders, ByRef $iExcludeDepth)
	Local $aFileList, $rgex
	__FO_FileSearchMask($sFileList, $sPath, $iDepth, $aExcludeFolders, $iExcludeDepth)
	$sFileList = StringTrimRight($sFileList, 2)
	$sMask = StringReplace(StringReplace(StringRegExpReplace($sMask, '[][$^.{}()+]', '\\$0'), '?', '.'), '*', '.*?')

	Switch $sLocale
		Case -1
			$rgex = 'i'
		Case 1
		Case 0
			$sLocale = '\x{80}-\x{ffff}'
			ContinueCase
		Case Else
			$rgex = 'i'
			$sMask = __FO_UserLocale($sMask, $sLocale)
	EndSwitch

	If $fInclude Then
		$aFileList = StringRegExp($sFileList, '(?m' & $rgex & ')^([^|]+\|(?:' & $sMask & '))(?:\r|\z)', 3)
		$sFileList = ''
		For $i = 0 To UBound($aFileList) - 1
			$sFileList &= $aFileList[$i] & @CRLF
		Next
	Else
		$sFileList = StringRegExpReplace($sFileList & @CRLF, '(?m' & $rgex & ')^[^|]+\|(' & $sMask & ')\r\n', '')
	EndIf
	$sFileList = StringReplace($sFileList, '|', '')
EndFunc   ;==>__FO_GetListMask

Func __FO_UserLocale($sMask, $sLocale)
	Local $s, $tmp
	$sLocale = StringRegExpReplace($sMask, '[^' & $sLocale & ']', '')
	$tmp = StringLen($sLocale)
	For $i = 1 To $tmp
		$s = StringMid($sLocale, $i, 1)
		If $s Then
			If StringInStr($sLocale, $s, 0, 2, $i) Then
				$sLocale = $s & StringReplace($sLocale, $s, '')
			EndIf
		Else
			ExitLoop
		EndIf
	Next
	If $sLocale Then
		Local $Upper, $Lower
		$tmp = StringSplit($sLocale, '')
		For $i = 1 To $tmp[0]
			$Upper = StringUpper($tmp[$i])
			$Lower = StringLower($tmp[$i])
			If Not ($Upper == $Lower) Then $sMask = StringReplace($sMask, $tmp[$i], '[' & $Upper & $Lower & ']')
		Next
	EndIf
	Return $sMask
EndFunc   ;==>__FO_UserLocale

Func __FO_FileSearchType(ByRef $sFileList, $sPath, $sMask, ByRef $fInclude, ByRef $iDepth, ByRef $aExcludeFolders, ByRef $iExcludeDepth, $iCurD = 0)
	Local $iPos, $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			If $iCurD >= $iDepth Or ($iCurD <= $iExcludeDepth And __ChExcludeFolders($sFile, $aExcludeFolders)) Then ContinueLoop
			__FO_FileSearchType($sFileList, $sPath & $sFile & '\', $sMask, $fInclude, $iDepth, $aExcludeFolders, $iExcludeDepth, $iCurD + 1)
		Else
			$iPos = StringInStr($sFile, ".", 0, -1)
			If $iPos And StringInStr($sMask, '|' & StringTrimLeft($sFile, $iPos) & '|') = $fInclude Then
				$sFileList &= $sPath & $sFile & @CRLF
			ElseIf Not $iPos And Not $fInclude Then
				$sFileList &= $sPath & $sFile & @CRLF
			EndIf
		EndIf
	WEnd
	FileClose($s)
EndFunc   ;==>__FO_FileSearchType

Func __FO_FileSearchMask(ByRef $sFileList, $sPath, ByRef $iDepth, ByRef $aExcludeFolders, ByRef $iExcludeDepth, $iCurD = 0)
	Local $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			If $iCurD >= $iDepth Or ($iCurD <= $iExcludeDepth And __ChExcludeFolders($sFile, $aExcludeFolders)) Then ContinueLoop
			__FO_FileSearchMask($sFileList, $sPath & $sFile & '\', $iDepth, $aExcludeFolders, $iExcludeDepth, $iCurD + 1)
		Else
			$sFileList &= $sPath & '|' & $sFile & @CRLF
		EndIf
	WEnd
	FileClose($s)
EndFunc   ;==>__FO_FileSearchMask

Func __FO_FileSearchAll(ByRef $sFileList, $sPath, ByRef $iDepth, ByRef $aExcludeFolders, ByRef $iExcludeDepth, $iCurD = 0)
	Local $sFile, $s = FileFindFirstFile($sPath & '*')
	If $s = -1 Then Return
	While 1
		$sFile = FileFindNextFile($s)
		If @error Then ExitLoop
		If @extended Then
			If $iCurD >= $iDepth Or ($iCurD <= $iExcludeDepth And __ChExcludeFolders($sFile, $aExcludeFolders)) Then ContinueLoop
			__FO_FileSearchAll($sFileList, $sPath & $sFile & '\', $iDepth, $aExcludeFolders, $iExcludeDepth, $iCurD + 1)
		Else
			$sFileList &= $sPath & $sFile & @CRLF
		EndIf
	WEnd
	FileClose($s)
EndFunc   ;==>__FO_FileSearchAll

Func __ChExcludeFolders(ByRef $sFile, ByRef $aExcludeFolders)
	For $i = 1 To $aExcludeFolders[0]
		If $sFile = $aExcludeFolders[$i] Then Return True
	Next
	Return False
EndFunc   ;==>__ChExcludeFolders
