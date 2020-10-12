#cs	==================================================================================================================
	Title 				:	jnknsByteError
	AutoIt Version	: 	3.3.14.5
	Language		: 	English
	Description		:	Fix Error Regarding StartUp Address
	Author	/s			: 	cjhernandez
                            :   rdbayanado
#ce	==================================================================================================================

#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <StringConstants.au3>
#include <Array.au3>
#include <Excel.au3>
#include <WinAPIFiles.au3>
#include <File.au3>
#include<FileConstants.au3>
#include <WinAPI.au3>
#Include <Misc.au3>
#include <Date.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>

#include "..\jnknsMainInitialize.au3"
#include "..\Tracelog\jnknsProcessLogger.au3"
#include "..\ErrHandler\jnknsErrHandler.au3"

Global  $sLogTextFile= @ScriptDir & "\..\Log.txt"
Global  $sSoftwarePathStartUp = ""
Global  $g_JPL_txtfile = @ScriptDir & '\..\TraceLog\jenkins-trace-log.txt'
Global Enum $HAN_GUI, $HAN_TREE, $HAN_BTN, $HAN_BTN2, $HAN_COUNT
Global		$g_iJM_Handles[$HAN_COUNT], _
				$g_iJM_Spider_F5_Class , _							;	Class value depending in the AutoIt v3 Window info
				$g_iJM_Spider_File_Class, _							;	Class value depending in the AutoIt v3 Window info
				$g_iJM_Spider_Software_Path_Class , _		;	Class value depending in the AutoIt v3 Window info
				$g_iJM_Spider_Run_Class

Local   $sRet, _
            $sTextClasses
Local   $sTPRJPATH, _
            $sDrive, _
            $sFile, _
			$sSoftwarePath
Local   $oExcel
Local   $hTextFile
Local   $sTestSheetFile

; Open log text file
$hTextFile = FileOpen($sLogTextFile, $FO_READ)

; Open testsheet
$sTPRJPATH = FileReadLine($hTextFile, 1)
$sTPRJPATH = StringTrimLeft($sTPRJPATH, 11)
$g_sJMI_TPRJ_Path = $sTPRJPATH
$sSoftwarePath = StringTrimRight($g_sJMI_TPRJ_Path,21)
;$sSoftwarePathStartUp = StringTrimRight($sTPRJPATH,21)    ; added by prdedumo
$sTPRJPATH=StringTrimRight($sTPRJPATH,20)
$sTestSheetFile = FileReadLine($hTextFile,2)
$sTestSheetFile = StringTrimLeft($sTestSheetFile,20)

; Initializa Environment
_JMI_jnknsCallDSpider()
Sleep(2000)
_JMI_jnknsSpiderSettings()

; Gets the information
$sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle($g_sJMI_Spider_Version))
if _JMI_jnknsBuildTree($sTextClasses) Then
EndIf

; Write countermeasure to log file
_JPL_jnknsCreatelogfile('Setting Start-Up address', $sTestSheetFile, 'Computing address', 'Yes', "start")			; start logging of countermeasure
$sRet = _STRE_jnkns_CheckCurrentAddress($sTPRJPATH)
If $sRet = 1 Then
    _JEH_RefreshSettings($sSoftwarePath & "\")
EndIf

; Write countermeasure to log file
_JPL_jnknsCreatelogfile('Setting Start-Up address', "", 'Exiting countermeasure', 'Yes', 'End')
FileClose($hTextFile)

Exit

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name					:   _STRE_jnkns_GetXcellSheetPath
; Description		:	Return Test Sheet Path
; Syntax				:	_STRE_jnkns_GetXcellSheetPath()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:	rdbayanado
; Modified			:	None
;=====================================================================================================================
Func _STRE_jnkns_GetXcellSheetPath()
    Local $sTestSheetFile, _
            $sDrive, _
            $sDir, _
            $sFileName, _
            $sExtension, _
            $sFile
    Local $oExcel
    Local $hTextFile
    Local $fTestDesign

    ; Open log text file
    $hTextFile = FileOpen($sLogTextFile, $FO_READ)
    ; Open testsheet
    $sTestSheetFile = FileReadLine($hTextFile, 2)
    $sTestSheetFile = StringTrimLeft($sTestSheetFile, 20)
    $oExcel = _Excel_Open(False)

    Local $aTestDesign[] = [$sTestSheetFile]

    For $i = 0 To UBound($aTestDesign, 1) - 1
        _PathSplit($aTestDesign[$i], $sDrive, $sDir, $sFileName, $sExtension)
    Next
    $sFile = $sDrive & $sDir & $sFileName & $sExtension
    Return $sFile
EndFunc ;==>_STRE_jnkns_GetXcellSheetPath

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name					:   STRE_jnknsGetStartUp
; Description		:	Get correct start up address
; Syntax				:	STRE_jnknsGetStartUp()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:	rdbayanado
; Modified			:	None
;=====================================================================================================================
Func STRE_jnkns_GetStartUp($fStartupFile)
    Local $aSrtArray[35]
    Local $hFile
    Local $iSrtLine, _
            $iSLine
    Local $sStartAddress, _
            $sEndAddress, _
            $sResultAddress

    $hFile = FileOpen($fStartupFile&"startup.lst",$FO_READ)
    $iSrtLine=0
    $srt_sLine=""
    If $hFile = -1 Then ;Validation for File Handling
        MsgBox(0,'ERROR','Unable to open file for reading.')
        Exit 1
    EndIf
    ; find the line that has the search strin g
    While 1
        $iSrtLine += 1
        Local $srt_sLine = FileReadLine($hFile) ;Store into array
        If @error = -1 Then ExitLoop
        If StringInStr($srt_sLine, "StartUp_Main")  Then ;Find string StartUp_Main as a reference for position
            For $i = $iSrtLine+1 To $iSrtLine+34 ; Setup range to be search for the string  from the startpoin
                $aSrtArray[$i-$iSrtLine] = FileReadLine($hFile, $i)
                If StringInStr( $aSrtArray[$i-$iSrtLine], "MOV		___ghsend_bss_startup_stk")  Then ;Search for the given string as next for the startpoint
                    $sStartAddress= $aSrtArray[$i-$iSrtLine] ;Set value of $srt_SrtAddress to the find string if matched
                    $sStartAddress=StringLeft($sStartAddress,8) ;Perform string sanitationSS
                    $sStartAddress = Dec($sStartAddress, $NUMBER_AUTO) ;Convert value into Decimal )
                ElseIf StringInStr(  $aSrtArray[$i-$iSrtLine], "MOV		0x000080A0, r7")  Then ;Find string next to the startpoint as reference for the end of start up address
                    $aSrtArray[$i-$iSrtLine] = FileReadLine($hFile, $i)
                    $sEndAddress= $aSrtArray[$i-$iSrtLine] ;Set value of $srt_EndAddress as the end startup  address
                    $sEndAddress=StringLeft($sEndAddress,8) ;Perform string sanitation
                    $sEndAddress = Dec($sEndAddress, $NUMBER_AUTO) ;Convert value into Decimal )
                EndIf
            Next
            ExitLoop
        EndIf
    WEnd
    $sResultAddress =  $sEndAddress-$sStartAddress ;Set Value for the $srt_ResultAddress
    $sResultAddress="0x" & Hex($sResultAddress,2) ;  ;Returns set of val
    FileClose($hFile) ;Close application Handler
    Return $sResultAddress ;Return Result address
EndFunc ;==>STRE_jnkns_GetStartUp

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name					:   STRE_jnknsCheckCurrentAddress
; Description		:	Get current Address from the startup file
; Syntax				:	STRE_jnknsCheckCurrentAddress()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:	rdbayanado
; Modified			:	None
;=====================================================================================================================
Func _STRE_jnkns_CheckCurrentAddress($txtfile) ;Compare current start up address into generated address
    Local $hFile
    Local $iSLine, _
            $iLine = 0, _
            $iSCorrect
    Local $sCurrentAddress, _
            $sCorrectAddress = STRE_jnkns_GetStartUp($txtfile);Set value from return of correct StartUp Address

    $txtfile=$txtfile&"UnitTestProject.tprj"

    $hFile=FileOpen($txtfile)
    If $hFile = -1 Then ;File Error Handler
        ;MsgBox(0,'ERROR','Unable to open file for reading.')
        Exit 1
    EndIf;===> ;File Error Handler
    While 1 ;Start of While Loop
        $iLine += 1
        $iSLine = FileReadLine($hFile)
        If @error = -1 Then
            ExitLoop
        EndIf
        If StringInStr($iSLine, "InitOffset")  Then
        $sCurrentAddress=FileReadLine($hFile,$iLine)

        $sCurrentAddress=StringRight($sCurrentAddress,4)
            If  ($sCurrentAddress<>$sCorrectAddress ) Then
                $iSCorrect=_ReplaceStringInFile($txtfile, $sCurrentAddress, $sCorrectAddress)
                If ($iSCorrect ==1) Then; Check if Process Succesful
                Else
                EndIf
            Else
            EndIf
        EndIf
	WEnd ;End Of While Loop

	_JPL_jnknsCreatelogfile('PL Error', '', 'Test : Calculating StartUp Addres', 'Yes', "= Passed")
	_JPL_jnknsCreatelogfile('PL Error', '', 'StartUp Address set to: ' & $sCorrectAddress, 'Yes', @CRLF & @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")

    FileClose($hFile) ; Close File Handler Object
    Return 1
    
EndFunc ;==>_STRE_jnkns_CheckCurrentAddress
