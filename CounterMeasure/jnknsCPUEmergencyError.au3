#cs	==================================================================================================================
	Title 				:	jnknsCPUEmergencyError
	AutoIt Version	: 	3.3.14.5
	Language		: 	English
	Description		:	Countermeasure Main process for fixing CPU_Emergency
	Author				: 	prdedumo
    Version            :    0.1
#ce	==================================================================================================================

#include <FileConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Excel.au3>

#include "..\jnknsMainInitialize.au3"
#include "..\ErrHandler\jnknsErrHandler.au3"
#include "..\TraceLog\jnknsProcessLogger.au3"

Local	$sLogTextFile = @ScriptDir & '\..\Log.txt', _
			$sTestSheetFile, _
			$sTprjPath, _
			$sSpider_Ver, _
			$sStatus, _
			$sSpider_Run_Class = "テスト実行中", _
			$sSpider_Local, _
            $sUnitTest_Log_TxtFile, _
            $sSpider_Log_TxtFile, _
            $sSoftwarePath = "", _
            $sFileAttrib
Local	$hTextFile
Local    $aArrayA, _
            $aArrayB, _
            $aArrayC, _
            $aArrayD[11], _
            $aTarget_Source
Local   $iCPU_Count
Local   $bDone

Local   $sCount_txtfile = @ScriptDir & '\..\TraceLog\1.txt'
;   queue file text path
Local   $sQuePath = ""
Global $g_JPL_txtfile = @ScriptDir & '\..\TraceLog\jenkins-trace-log.txt'

; Open log text file
$hTextFile = FileOpen($sLogTextFile, $FO_READ)
If $hTextFile = -1 Then
    _JPL_jnknsCreatelogfile('CPU_Emergency Error', "", 'Error: Cannot read log file', 'No', "")
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
; Spider Version
$sSpider_Ver = FileReadLine($hTextFile,5)
$sSpider_Ver = StringTrimLeft($sSpider_Ver,16)
$g_sJMI_Spider_Version = $sSpider_Ver
; Target Source
$aTarget_Source = _StringBetween($sTestSheetFile,"【", "】")
; ====================================
$sSoftwarePath = StringTrimRight($g_sJMI_TPRJ_Path, 21)
$sStatus = FileReadLine($hTextFile,4)
$sStatus = StringTrimLeft($sStatus,8)

; search the target source path
$aArrayA = _FileListToArray($sSoftwarePath, $aTarget_Source[0], 1, True)
If @error = 1 Then
    _JPL_jnknsCreatelogfile('CPU_Emergency Error', "", 'Error: Path was invalid', 'No', "")
    Exit
EndIf
If @error = 4 Then
    _JPL_jnknsCreatelogfile('CPU_Emergency Error', "", 'Error: No file(s) were found', 'No', "")
    Exit
EndIf

$iCPU_Count = 0
$bDone = False
$aArrayB = FileReadToArray($aArrayA[1])
Local $iLineCountA = @extended
If @error Then
    _JPL_jnknsCreatelogfile('CPU_Emergency Error', "", 'Error: There was an error reading the file', 'No', "")
Exit
Else
    For $i = 1 To $iLineCountA -1
        If StringInStr ($aArrayB[$i], 'CPU_EMERGENCY') Then
            $iCPU_Count = $iCPU_Count + 1   ; CPU counter
        EndIf
    Next
EndIf
; Condition if CPU EMERGENCY is found
If $iCPU_Count > 0 Then
    ; start logging of countermeasure
    _JPL_jnknsCreatelogfile('CPU_Emergency', $sTestSheetFile, 'Test : Editing definition', 'Yes', "start")
    $aArrayC = _JEH_FindInFile('CPU_EMERGENCY', $sSoftwarePath, '*.h')  ; Search definition file of CPU_Emergency
    ; Loop for any declaration found
    For $i = 1 To $aArrayC[0]
        $aArrayB = FileReadToArray($aArrayC[$i])    ;  Array to store all lines in the file found
        $iLineCountB = @extended
        If @error Then
            _JPL_jnknsCreatelogfile('CPU_Emergency Error', "", 'Error: There was an error reading the file', 'No', "Failed")
            Exit
        Else
            For $j= 0 To $iLineCountB -1
                $x = 0
                ReDim $aArrayD[$iLineCountB]
                ReDim $aArrayB[$iLineCountB]
                If StringInStr ($aArrayB[$j], 'CPU_EMERGENCY') And StringInStr ($aArrayB[$j], 'define') Then
                    For $k = ($j +10)  To $j Step - 1
                         If StringInStr ($aArrayB[$k], 'do') Or  (StringInStr ($aArrayB[$k], 'while') And StringInStr ($aArrayB[$k], '0')) Or StringInStr ($aArrayB[$k], 'define') Then
                         Else
                            If $aArrayB[$k] <> ""  Then
                                Sleep(200)
                                $aArrayD[$x] = $k   ; Assign line number to arrays to be deleted
                                $x = $x + 1
                            EndIf
                        EndIf
                    Next
                        ; Checking file if readonly
                        _JEH_SetAttrib($aArrayC[$i])
                        ReDim $aArrayD[$x]
                        For $x = 0 To UBound($aArrayD) - 1
                            If $aArrayD[$x] <> "" Then
                                _ArrayDelete($aArrayB, $aArrayD[$x])    ; Delete line number
                            EndIf
                        Next
                        Sleep(200)
                            _FileWriteFromArray($aArrayC[$i], $aArrayB, 1)   ; Write all lines to the file
                EndIf
            Next
        EndIf
    next
    Sleep(1000)
    $bDone = True
    _JPL_jnknsCreatelogfile('CPU_Emergency Error', "", 'Test : Editing definition', 'Yes', "= Passed")
EndIf

If $bDone Then
    ; Create Text File
    Local $hFileOpen = FileOpen($sCount_txtfile, 2)
        If $hFileOpen = -1 Then
            Exit
        EndIf
    FileWriteLine($hFileOpen,"CPU_Emergency countermeasure applied")
#cs ========================================================
    This was commented out since this is a pre-run countermeasure
    This will be run together after applying/checking the 3 pre-run countermeasures
    Local $hFileOpen = FileOpen($sCount_txtfile)
    ; Rebuild Test Environment
    If _JEH_Rebuild_Software ($sSoftwarePath) Then
        Sleep(20000)
        ; Refresh FSUnit Settings
        _JEH_RefreshSettings($sSoftwarePath & '\')
        ; Re-run the sheet
        _JMI_jnknsCallDSpider()
        WinActivate($g_sJMI_Spider_Version)
        MouseClick("Left",609, 299)
        ; Loop to wait until running of the tool is done
        $sSpider_Local = WinActivate($sSpider_Run_Class)
        While 1
            Sleep(1000)
            $sSpider_Local = WinActivate($sSpider_Run_Class)
            if $sSpider_Local <> 0 Then
            Else
                ExitLoop
            EndIf
        WEnd
        Sleep(2000)
        ; Rechecks if different error occured
        $sUnitTest_Log_TxtFile = $sSoftwarePath & "\UnitTest\log.txt"
        $sSpider_Log_TxtFile = @ScriptDir & '\..\Log.txt'
        If _JMI_jnknsReCheckIfError($sUnitTest_Log_TxtFile, $sSpider_Log_TxtFile) Then
            Exit
        EndIf
    EndIf
#ce ========================================================
    _JPL_jnknsCreatelogfile('CPU_Emergency Error', "", 'Exiting countermeasure', 'Yes', 'End')
EndIf
FileClose($hTextFile)
Exit
