#cs	==================================================================================================================
	Title 				:	 jnknsMain
	AutoIt Version	: 	3.3.14.5
	Language		: 	English
	Description		:	Main process for setting up the Jenkins environment
	Author				: 	prdedumo
    Version            :    0.1
#ce	==================================================================================================================

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>

#include "jnknsMainInitialize.au3"
#include "TraceLog\jnknsProcessLogger.au3"
#include "ErrHandler\jnknsErrHandler.au3"

Global  $g_JPL_txtfile = @ScriptDir & '\TraceLog\jenkins-trace-log.txt'
Global Enum $HAN_GUI, $HAN_TREE, $HAN_BTN, $HAN_BTN2, $HAN_COUNT
Global		$g_iJM_Handles[$HAN_COUNT], _
			$g_iJM_Spider_F5_Class , _							;	Class value depending in the AutoIt v3 Window info
			$g_iJM_Spider_File_Class, _							;	Class value depending in the AutoIt v3 Window info
			$g_iJM_Spider_Software_Path_Class , _		;	Class value depending in the AutoIt v3 Window info
			$g_iJM_Spider_Run_Class							;	Class value depending in the AutoIt v3 Window info

Local	$sRetShowForm, _
			$sTextClasses, _
            $sSoftwarePath, _
            $spider_UnitLog_TxtFile
Local   $initStatus
Local   $retBuild
Local   $iFileCounter

Local   $sLogTextFile = @ScriptDir & '\Log.txt'
Local   $aTextFiles[] = [@ScriptDir &"\TraceLog\1.txt", @ScriptDir &"\TraceLog\2.txt", @ScriptDir &"\TraceLog\3.txt", @ScriptDir & "\TraceLog\jenkins-trace-log.txt"]
Global $g_fileDeleter = 0

; Initialize FSUnit Title
While 1
    If _JMI_jnknsCallDSpider() Then
        ExitLoop
    EndIf
WEnd
; Gets the information
$sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle($g_sJMI_Spider_Version))
if _JMI_jnknsBuildTree($sTextClasses) Then
    $iFileCounter = 0
    $initStatus = 0
    ; Creating first log
    _JPL_jnknsCreatelogfile("", "", "", "", "")
    ; Log text for pre-countermeasure
    ;$initStatus = _JMI_jnksEnvironmentLog()
    _JMI_jnknsSpiderSettings()
    $sSoftwarePath = StringTrimRight($g_sJMI_TPRJ_Path, 21)
	$spider_UnitLog_TxtFile = $sSoftwarePath & "\UnitTest\log.txt"

    For $i = 0 To UBound($aTextFiles) - 1
        If FileExists($aTextFiles[$i]) Then
            FileDelete($aTextFiles[$i])
        EndIf
        Sleep(100)
    Next

EndIf
Exit
