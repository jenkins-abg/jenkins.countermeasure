#cs	==================================================================================================================
	Title 				:	jnknsCResultError
	AutoIt Version	: 	3.3.14.5
	Language		: 	English
	Description		:	Countermeasure Main process for fixing Comment Result sheet error
	Author				: 	prdedumo
    Version            :    0.1
#ce	==================================================================================================================

#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Excel.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>

#include "..\jnknsMainInitialize.au3"
#include "..\ErrHandler\jnknsErrHandler.au3"
#include "..\TraceLog\jnknsProcessLogger.au3"

Global  $g_JPL_txtfile = @ScriptDir & '\..\TraceLog\jenkins-trace-log.txt'
Global Enum $HAN_GUI, $HAN_TREE, $HAN_BTN, $HAN_BTN2, $HAN_COUNT
Global		$g_iJM_Handles[$HAN_COUNT], _
				$g_iJM_Spider_F5_Class , _							;	Class value depending in the AutoIt v3 Window info
				$g_iJM_Spider_File_Class, _							;	Class value depending in the AutoIt v3 Window info
				$g_iJM_Spider_Software_Path_Class , _		;	Class value depending in the AutoIt v3 Window info
				$g_iJM_Spider_Run_Class							;	Class value depending in the AutoIt v3 Window info

Local	$sLogTextFile = @ScriptDir & '\..\Log.txt', _
			$sTestSheetFile, _
			$sTprjPath, _
			$sSpider_Ver, _
			$sComment_Result, _
			$sStatus, _
			$sSpider_F5_Class = "テスト実行中", _
			$sSpider_Local, _
            $sUnitTest_Log_TxtFile, _
            $sSpider_Log_TxtFile, _
            $sSoftwarePath = "", _
            $sSheetVer
Local	$sRetShowForm, _
			$sTextClasses

Local	$iErrNumber
Local	$hTextFile
Local	$oExcel, _
			$oWorkbook

; Open log text file
$hTextFile = FileOpen($sLogTextFile, $FO_READ)
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
;~ $sComment_Result = FileReadLine($hTextFile,7)
;~ $sComment_Result = StringTrimLeft($sComment_Result,13)
; ====================================
$sSoftwarePath = StringTrimRight($g_sJMI_TPRJ_Path,21)
$sStatus = FileReadLine($hTextFile,4)
$sStatus = StringTrimLeft($sStatus,8)

; Initialize FSUnit Title
_JMI_jnknsCallDSpider()
; Gets the information
$sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle($g_sJMI_Spider_Version))
if _JMI_jnknsBuildTree($sTextClasses) Then
EndIf
; Condition to check log error file
;~ if $iErrNumber = 3 And $sStatus = "Pending" Then
    ; start logging of countermeasure
    _JPL_jnknsCreatelogfile('Comment Result Error', $sTestSheetFile, 'Test : Renaming Test Sheet', 'Yes', "start")
	$oExcel = _Excel_Open()
	If @error Then
        _JPL_jnknsCreatelogfile('Comment Result Error', "", 'Error: There was an error reading the file', 'No', 'Failed')
        Exit
	EndIf
    ; Open an existing workbook and return its object identifier.
	$oWorkbook = _Excel_BookOpen($oExcel, $g_sJMI_TestDesign_File)
    If @error Then
        _JPL_jnknsCreatelogfile('Comment Result Error', "", 'Error: There was an error reading the file', 'No', 'Failed')
        Exit
    EndIf
	; Fix the test sheet name
;~ 	$oWorkbook.Sheets(2).name = $sComment_Result
    $sSheetVer= _Excel_RangeRead($oWorkbook, 1, "A1")
    $sComment_Result = $oWorkbook.Sheets(2).name
	_Excel_BookClose($oWorkbook,True)
    _JPL_jnknsCreatelogfile('Comment_Result Error', "", 'Test : Sheet name to ' & $sComment_Result, 'Yes', "= Passed")
	_Excel_Close($oExcel)
	$oExcel = Null
    ;Check if $sComment_Result string is valid
    _JEH_RefreshSettings($sSoftwarePath & "\","",$sComment_Result)
    _JPL_jnknsCreatelogfile('Comment_Result Error', "", 'Exiting countermeasure', 'Yes', 'End')
    If ($sComment_Result <> "Comment_Result") Or ($sComment_Result <> "Comment_カバレッジ結果") Then
       _JPL_jnknsCreatelogfile('Comment_Result Error', "", 'Invalid Sheetname', 'No', 'Failed ')
    EndIf
    If $sSheetVer =="(Format Rev3.00)" Then
        _JPL_jnknsCreatelogfile('Comment_Result Error', "", 'Format Rev3.00', 'Yes', 'Proceed')
    ElseIF $sSheetVer =="(Format Rev3.00)" Then
       _JPL_jnknsCreatelogfile('Comment_Result Error', "", 'Format Rev2.00', 'Yes', 'Proceed')
    Else
        _JPL_jnknsCreatelogfile('Comment_Result Error', "", 'Invalid Sheet Version', 'No', 'Failed ')
    EndIf

#cs ========================================================
    ;This was commented out since the next pipeline method is PL Error
    ;PL Error must run the NAISEI to detect error
	; Re-run the sheet
    _JMI_jnknsPressF5($g_sJMI_Spider_Version)
    ; Rechecks if different error occured
	$sUnitTest_Log_TxtFile = $sSoftwarePath & "\UnitTest\log.txt"
	$sSpider_Log_TxtFile = @ScriptDir & '\..\Log.txt'
    _JEH_jnknsCreateLogFile("3", "OK", $sLogTextFile)
    If _JMI_jnknsReCheckIfError($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile) Then
        Exit
    EndIf
#ce ========================================================
;~ EndIf
FileClose($hTextFile)
Exit
