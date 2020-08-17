#cs	==================================================================================================================
	Title 				:	jnknsAmbiguousError
	AutoIt Version	: 	3.3.14.5
	Language		: 	English
	Description		:	Countermeasue Main process for fixing ambiguous error
	Author				: 	prdedumo
    Version            :    0.1
#ce	==================================================================================================================

#include <File.au3>
#include <Array.au3>
#include <FileConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Excel.au3>
#include <Misc.au3>
#include <AutoItConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>

#include "..\jnknsMainInitialize.au3"
#include "..\ErrHandler\jnknsErrHandler.au3"
#include "..\TraceLog\jnknsProcessLogger.au3"

Global Enum $HAN_GUI, $HAN_TREE, $HAN_BTN, $HAN_BTN2, $HAN_COUNT
Global  $g_iJM_Handles[$HAN_COUNT], _
            $g_iJM_Spider_F5_Class , _							;	Class value depending in the AutoIt v3 Window info
            $g_iJM_Spider_File_Class, _							;	Class value depending in the AutoIt v3 Window info
            $g_iJM_Spider_Software_Path_Class , _		;	Class value depending in the AutoIt v3 Window info
            $g_iJM_Spider_Run_Class							;	Class value depending in the AutoIt v3 Window info
Global  $g_JPL_txtfile = @ScriptDir & '\..\TraceLog\jenkins-trace-log.txt'

Local	$sLogTextFile = @ScriptDir & "\..\Log.txt", _
			$sErrNumber, _
			$sTestSheetFile, _
			$sTprjPath, _
			$sSpider_Ver, _
			$sFix_Value, _
			$sStatus, _
			$sSoftwarePath = "", _
			$sTargetFunction, _
			$sDrive = "", _
			$sDir = "", _
			$sFileName = "", _
			$sExtension = "", _
			$sSpider_Run_Class = "テスト実行中", _
			$sSpider_Local, _
            $sUnitTest_Log_TxtFile, _
            $sSpider_Log_TxtFile, _
            $sTextClasses
Local	$hTextFile
Local	$i
Local	$aArray, _
			$aPathSplit
Local	$iBackUpResult, _
			$iCopyResult, _
			$iRebuildResult, _
			$iErrNumber

; Open log text file
$hTextFile = FileOpen($sLogTextFile, $FO_READ)
If $hTextFile = -1 Then
    _JPL_jnknsCreatelogfile('Ambiguous Error', "", 'Error: Cannot read log file', 'No', "")
    Exit
EndIf
; Initialization of variables required
; ====================================
; Tprj Line number
$sTprjPath = FileReadLine($hTextFile,1)
$sTprjPath = StringTrimLeft($sTprjPath,11)
$g_sJMI_TPRJ_Path = $sTprjPath
; Test Sheet Line number
$sTestSheetFile = FileReadLine($hTextFile,2)
$sTestSheetFile = StringTrimLeft($sTestSheetFile,20)
$g_sJMI_TestDesign_File = $sTestSheetFile
; Error Number Line number
$iErrNumber = FileReadLine($hTextFile,3)
$iErrNumber = StringTrimLeft($iErrNumber,14)
; Spider Version
$sSpider_Ver = FileReadLine($hTextFile,5)
$sSpider_Ver = StringTrimLeft($sSpider_Ver,16)
$g_sJMI_Spider_Version = $sSpider_Ver
; Comment Result Fix
$sFix_Value = FileReadLine($hTextFile,7)
$sFix_Value = StringTrimLeft($sFix_Value,13)
; ====================================
$sStatus = FileReadLine($hTextFile,4)
$sStatus = StringTrimLeft($sStatus,8)
$sSoftwarePath = StringTrimRight($g_sJMI_TPRJ_Path,21)

$iBackUpResult = 0
$iCopyResult = 0
$iRebuildResult = 0

; Initialize FSUnit Title
_JMI_jnknsCallDSpider()
; Gets the information
$sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle($g_sJMI_Spider_Version))
if _JMI_jnknsBuildTree($sTextClasses) Then
EndIf
; Check Error
$sUnitTest_Log_TxtFile = $sSoftwarePath & "\UnitTest\log.txt"
$sSpider_Log_TxtFile = @ScriptDir & '\..\Log.txt'
If _JMI_jnknsReCheckIfError($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile) Then
EndIf
; Log text constraints must satisfy the MCDC condition
if $iErrNumber = 4 And $sStatus = "Pending" Then
	$sTargetFunction = _StringBetween($sTestSheetFile, "【", "】", $STR_ENDNOTSTART)
    ; start logging of countermeasure
    _JPL_jnknsCreatelogfile('Ambiguous Error', $sTestSheetFile, 'Test : Finding instances', 'Yes', "start")
	; Find ambiguous in all software file
	$aArray = _JEH_FindInFile($sFix_Value & '*', $sSoftwarePath, '*.c;*.h') ; Search for the fix value created in the log file
	$i = 0
	if $aArray[0] > 1 Then
		For $i = 1 to $aArray[0] step 1
			; Checks if file is a target unit, skip if yes
			if StringInStr($aArray[$i],$sTargetFunction[0]) Then
			else
				$aPathSplit = _PathSplit($aArray[$i], $sDrive, $sDir, $sFileName, $sExtension)
 				; BackUp file before editing
				Sleep(2000)
				$iBackUpResult = _JEH_BackUpFile($aArray[$i], $sSoftwarePath, $aPathSplit[$PATH_FILENAME], $aPathSplit[$PATH_EXTENSION])
				Sleep(2000)
				if $iBackUpResult Then
					; Edit the software if backup is done
                    _JPL_jnknsCreatelogfile('Ambiguous Error', "", 'Test : Creating Backup', 'Yes', "= Passed")
					$iCopyResult = _JEH_EditFile ($aArray[$i], $sFix_Value)
				EndIf
			EndIf
		Next
	EndIf
     _JPL_jnknsCreatelogfile('Ambiguous Error', "", 'Test : Editing File', 'Yes', "= Passed")
	if $iCopyResult Then
		; Rebuild Test Environment if copying of files is complete
		$iRebuildResult = _JEH_Rebuild_Software ($sSoftwarePath)
        _JPL_jnknsCreatelogfile('Ambiguous Error', "", 'Test : Rebuilding Software', 'Yes', "= Passed")
	EndIf
    Sleep(200)
    ; Refresh FSUnit Settings
    _JEH_RefreshSettings($sSoftwarePath)
	; Re-run the sheet
    _JMI_jnknsPressF5($g_sJMI_Spider_Version)
    ; Rechecks if different error occured
	$sUnitTest_Log_TxtFile = $sSoftwarePath & "\UnitTest\log.txt"
	$sSpider_Log_TxtFile = @ScriptDir & '\..\Log.txt'
    If _JMI_jnknsReCheckIfError($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile) Then
    EndIf
    _JPL_jnknsCreatelogfile('Ambiguous Error', "", 'Exiting countermeasure', 'Yes', 'End')
    Exit
EndIf
FileClose($hTextFile)
Exit
