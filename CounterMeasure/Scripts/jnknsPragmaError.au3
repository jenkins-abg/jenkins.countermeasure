#cs	==================================================================================================================
	Title 				:	jnknsPragmaError
	AutoIt Version	: 	3.3.14.5
	Language		: 	English
	Description		:	Countermeasure Main process for fixing Pragma
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
            $sSoftwarePath = ""
Local	$hTextFile
Local    $aArrayA, _
            $aArrayB, _
            $aArrayC, _
            $aArrayD[11], _
            $aTarget_Source
Local   $iPragma_Count, _
            $iLine_Count[10], _
            $iFunction_Definition_Count
Local   $bDone
Local   $sCount_txtfile = @ScriptDir & '\..\TraceLog\2.txt'
Global  $g_JPL_txtfile = @ScriptDir & '\..\TraceLog\jenkins-trace-log.txt'

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
; Spider Version
$sSpider_Ver = FileReadLine($hTextFile,5)
$sSpider_Ver = StringTrimLeft($sSpider_Ver,16)
$g_sJMI_Spider_Version = $sSpider_Ver
; Target Source
$aTarget_Source = _StringBetween($sTestSheetFile,"【", "】")
; ====================================
$sSoftwarePath = StringTrimRight($g_sJMI_TPRJ_Path,21)
$sStatus = FileReadLine($hTextFile,4)
$sStatus = StringTrimLeft($sStatus,8)

; search the target source path
$aArrayA = _FileListToArray($sSoftwarePath, $aTarget_Source[0], Default, True)
If @error = 1 Then
    _JPL_jnknsCreatelogfile('Pragma Error', "", 'Error: Path was invalid', 'No', "")
    Exit
EndIf
If @error = 4 Then
    _JPL_jnknsCreatelogfile('Pragma Error', "", 'Error: No file(s) were found', 'No', "")
    Exit
EndIf

; initialize variable
$x = 0
$iPragma_Count = 0
$bDone = False

$aArrayB = FileReadToArray($aArrayA[1])
Local $iLineCountA = @extended
If @error Then
     _JPL_jnknsCreatelogfile('Pragma Error', "", 'Error: There was an error reading the file', 'No', 'Failed')
     Exit
Else
    For $i = 1 To $iLineCountA -1
        If StringInStr ($aArrayB[$i], '#pragma') Then
            $iPragma_Count = $iPragma_Count + 1   ; Pragma counter
        EndIf
    Next
EndIf

 ; Condition if Pragma is found
If $iPragma_Count > 0 Then
    ; start logging of countermeasure
    _JPL_jnknsCreatelogfile('Pragma Error', $sTestSheetFile, 'Test : Editing definition', 'Yes', "start")
    ReDim $iLine_Count[$iLineCountA]
    ; Getting Line number of pragma
    For $i = 1 To $iLineCountA -1
        If StringInStr ($aArrayB[$i], '#pragma') Then
            $iLine_Count[$x] = $i
            $x = $x + 1
        EndIf
    Next
    $iFunction_Definition_Count = 0
    ReDim $iLine_Count[$x + 1]
    ; Loop for number of pragma found
    For $x = 0 To (UBound($iLine_Count) -1)
        ; loop for checking in-between of two consecutive pragmas
        If $iLine_Count[$x] <> "" Then
            For $k = $iLine_Count[$x] To $iLine_Count[$x+1]
                ; if next line is another pragma, proceed to next iteration of line
                If StringInStr ($aArrayB[$k], '関数名') Then
                    $iFunction_Definition_Count = $iFunction_Definition_Count + 1
                EndIf
            Next
            If $iFunction_Definition_Count = 1 Then
                    _FileWriteToLine($aArrayA[1], $iLine_Count[$x] + 1, '/* ' & $aArrayB[$iLine_Count[$x]]& ' */', 1)
                    _FileWriteToLine($aArrayA[1], $iLine_Count[$x+1] +1, '/* ' & $aArrayB[$iLine_Count[$x + 1]]& ' */', 1)
            EndIf
            $iFunction_Definition_Count = 0
        EndIf
        Sleep(100)
    Next
    Sleep(1000)
    $aArrayB = FileReadToArray($aArrayA[1])
    $iLineCountA = @extended
    ; Loop if there are two pragmas enclosing
    For $x = 0 To (UBound($iLine_Count) -1)
        ; loop for checking in-between of two consecutive pragmas
        If $iLine_Count[$x] <> "" Then
            For $k = $iLine_Count[$x] To $iLine_Count[$x+1]
                ; Checking preceding pragmas
                If ($iLine_Count[$x+1] - $iLine_Count[$x]) = 1 And StringInStr($aArrayB[$k+1], '#pragma') And StringInStr($aArrayB[$k+1], '/*') Then
                    _FileWriteToLine($aArrayA[1], $k+1, '/* ' & $aArrayB[$k]& ' */', 1)
                EndIf
                ; Checking succeeding pragmas
                 If ($iLine_Count[$x+1] - $iLine_Count[$x]) = 1 And StringInStr($aArrayB[$k-1], '#pragma') And StringInStr($aArrayB[$k-1], '/*') Then
                    _FileWriteToLine($aArrayA[1], $k + 1, '/* ' & $aArrayB[$k]& ' */', 1)
                EndIf
            Next
        EndIf
        Sleep(100)
    Next
     _JPL_jnknsCreatelogfile('Pragma Error', '', 'Test : Editing definition', 'Yes', "= Passed")
    $bDone = True
EndIf

If $bDone Then
    ; Create text file
   Local $hFileOpen = FileOpen($sCount_txtfile, 2)
        If $hFileOpen = -1 Then
            Exit
        EndIf
    FileWriteLine($hFileOpen,"Pragma countermeasure applied")
#cs ========================================================
    This was commented out since this is a pre-run countermeasure
    This will be run together after applying/checking the 3 pre-run countermeasures
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
    _JPL_jnknsCreatelogfile('Pragma Error', "", 'Exiting countermeasure', 'Yes', 'End')
EndIf
FileClose($hTextFile)
Exit
