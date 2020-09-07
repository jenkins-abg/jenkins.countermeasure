
#include-once

#cs	==================================================================================================================
	Title 				:	jnknsProcessErrorLogger
	AutoIt Version	: 	3.3.14.5
	Language		: 	English
	Description		:	Log Current Process
	Author				: 	rdbayanado
    Modified by    :    prdedumo
    Version            :    0.1
#ce	==================================================================================================================

;~ #AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w- 7

; #INCLUDES# ===========================================================================================================
    #include <FileConstants.au3>
    #include <StringConstants.au3>
    #include <WinAPIFiles.au3>
    #include <Date.au3>
; ====================================================================================================================

; #GLOBAL VARIABLES# =====================================================================================================
Global  $g_JPL_txtfile   ;Text File Location
; ====================================================================================================================

; #FUNCTION# ===========================================================================================================
; Name					:	_JPL_jnknsCreatelogfile($sCounterMeasure, $sFileName, $sEvent, $isPassed, $sStatus)
; Description		:	Get current window handle of cmd process
; Syntax				:	_JPL_jnknsCreatelogfile()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:	rdbayanado
; Modified			:	prdedumo
;=====================================================================================================================
Func _JPL_jnknsCreatelogfile($sCounterMeasure, $sFileName, $sEvent, $isPassed, $sStatus)
    Local   $runDate, _
            $runTime, _
            $sFileNameCheck = '', _
            $aTxtFileContent, _
            $iLineCountA, _
            $sInitPassed
    ; Initialize Variables
    $runTime = _DateTimeFormat(_NowTime(), 5)
    $runDate = _NowDate()
    $sInitPassed = 'No'

    $aTxtFileContent = FileReadToArray($g_JPL_txtfile)
    $iLineCountA = @extended
    If $isPassed <> "" Then
        $sInitPassed = $sStatus
    EndIf

    ; If file does not exist, create file and log
    If @error Then
        FileOpen($g_JPL_txtfile)
        FileWriteLine($g_JPL_txtfile, '#Software:' & @TAB & 'Jenkins Started')
        FileWriteLine($g_JPL_txtfile, '#Version:' & @TAB & '1.0.0')
        FileWriteLine($g_JPL_txtfile, '#Date:' & @TAB &@TAB  &  $runDate)
    EndIf
    ; initialize if new testsheet design
    If $sFileNameCheck <> $sFileName And $sStatus = "start" Then
        FileWriteLine($g_JPL_txtfile, @CRLF)
        FileWriteLine($g_JPL_txtfile, '#Filename:' & @TAB & $sFileName)
        FileWriteLine( $g_JPL_txtfile,  $runDate & @TAB & $runTime & @TAB &  @TAB &"Information" & @TAB & @TAB & "----------------------- Countermeasure started -----------------------")
    EndIf
    ; countermeasure applied logging
    If $iLineCountA > 0 And $sFileName = "" And $sStatus <> "End" And $sStatus = "= Passed" Then
        FileWriteLine( $g_JPL_txtfile,  $runDate & @TAB & $runTime & @TAB &  @TAB &"Information" & @TAB & @TAB & "Countermeasure: " & $sCounterMeasure )
    elseIf $sInitPassed = 'No' Then
	Else
        FileWriteLine( $g_JPL_txtfile,  $runDate & @TAB & $runTime & @TAB &  @TAB &"Information" & @TAB & @TAB & $sEvent  & " " & $sInitPassed )
    EndIf
    ; end of counter measure
    If $sStatus = "End" Then
        FileWriteLine( $g_JPL_txtfile,  $runDate & @TAB & $runTime & @TAB &  @TAB &"Information" & @TAB & @TAB & "----------------------- Countermeasure ended  ------------------------" )
    EndIf
EndFunc
