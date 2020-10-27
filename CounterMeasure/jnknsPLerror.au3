#cs	==================================================================================================================
	Title 				:	 jnknsPLerror
	AutoIt Version	: 	3.3.14.5
	Language		: 	AutoIt
	Description		:	Counter Measure Correct Wrong PL File
	Author				: 	rdbayanado
#ce	==================================================================================================================

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
#include <Date.au3>

#include "..\TraceLog\jnknsProcessLogger.au3"

Global $sLogTextFile = @ScriptDir & "\..\Log.txt"
Global  $g_JPL_txtfile = @ScriptDir & '\..\TraceLog\jenkins-trace-log.txt'

$sTextFile = FileOpen($sLogTextFile, $FO_READ)
; Initialization of variables required
; ====================================
; Tprj Line number
$sTprjPath = FileReadLine($sTextFile,1)
$sTprjPath = StringTrimLeft($sTprjPath,11)
; Test Sheet Line number
$sTestSheetFile = FileReadLine($sTextFile,2)
$sTestSheetFile = StringTrimLeft($sTestSheetFile,20)
; Error Number Line number
$iErrNumber = FileReadLine($sTextFile,3)
$iErrNumber = StringTrimLeft($iErrNumber,14)
; Spider Version
$sSpider_Ver = FileReadLine($sTextFile,5)
$sSpider_Ver = StringTrimLeft($sSpider_Ver,16)
; Comment Result Fix
$sFix_Value = FileReadLine($sTextFile,7)
$sFix_Value = StringTrimLeft($sFix_Value,13)
; ====================================
$sStatus = FileReadLine($sTextFile,4)
$sStatus = StringTrimLeft($sStatus,8)

_JPL_jnknsCreatelogfile('PL Error', $sTestSheetFile, 'Test : PL check...', 'Yes', "start")		; start logging of countermeasure

If _JPE_jnkns_processPLfile() Then
Else
    _JPL_jnknsCreatelogfile('PL Error', '', 'Test : Checking PL File', 'Yes', "= Passed")
    _JPL_jnknsCreatelogfile('PL Error', '', 'PL file used is correct', 'Yes', @CRLF & @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
EndIf

_JPL_jnknsCreatelogfile('PL Error', "", 'Exiting countermeasure', 'Yes', 'End')			 ; end of logging

FileClose($sTextFile)
Exit

; #FUNCTION# ===========================================================================================================
; Name					:	_JPE_jnkns_BinaryCompare
; Description		:	Get System Type
; Syntax				:	_JPE_jnkns_getSystemtype()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:   rdbayanado
; Modified			:	None
;=====================================================================================================================
Func _JPE_jnkns_BinaryCompare($sFilePath_1, $sFilePath_2, $iPercentageRead = 100)

    Local $hFileOpen = FileOpen($sFilePath_1, BitOR($FO_READ, $FO_BINARY))

    $iPercentageRead = Int($iPercentageRead)
    If $iPercentageRead > 100 Or $iPercentageRead < 0 Then
        $iPercentageRead = 100
    EndIf
    
    If $hFileOpen = -1 Then Return SetError(1, 0, False)

    Local $bBinary_1 = FileRead($hFileOpen, ($iPercentageRead / 100) * FileGetSize($sFilePath_1))
    FileClose($hFileOpen)

    $hFileOpen = FileOpen($sFilePath_2, BitOR($FO_READ, $FO_BINARY))
    If $hFileOpen = -1 Then Return SetError(2, 0, False)
    Local $bBinary_2 = FileRead($hFileOpen, ($iPercentageRead / 100) * FileGetSize($sFilePath_2))

    FileClose($hFileOpen)
    Return $bBinary_1 = $bBinary_2

EndFunc   ;==>_JPE_jnkns_BinaryCompare

; #FUNCTION# ===========================================================================================================
; Name				    :	_JPE_jnkns_getSystemtype
; Description		:	Get System Type
; Syntax				:	_JPE_jnkns_getSystemtype()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values	    : 	None
; Author				:   rdbayanado
; Modified			:	None
;=====================================================================================================================
Func _JPE_jnkns_getSystemtype()
    Local $temp_pathstring, $system_type

	$temp_pathstring = $sTprjPath
	$temp_pathstring&="Tools\SetWinAmsSpmcCode.pl"
	$temp_pathstring = StringSplit($temp_pathstring,"\")
	$system_type = $temp_pathstring[$temp_pathstring[0]-2]
	$system_type =StringLeft($system_type,4)

    If $system_type <> "" Then
    Else
        _JPL_jnknsCreatelogfile('PL Error', "", 'Test : Error no PL type for ' & $system_type , 'Yes', "= Failed")
    EndIf
    Return $system_type

EndFunc ;==>_JPE_jnkns_getSystemtype

; #FUNCTION# ===========================================================================================================
; Name				    :	_JPE_jnkns_getSystmkfile()
; Description		:	Get System Type
; Syntax				:	_JPE_jnkns_getSystmkfile()()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values	    : 	None
; Author				:   rdbayanado
; Modified			:	None
;=====================================================================================================================
Func _JPE_jnkns_getSystmkfile()

    Local $sTmpath, _
            $sSystm, _
            $hFile , _
            $aSystemArray[10]

    ;System Array List Definitions
    $aSystemArray[0] ="a2e"
    $aSystemArray[1] ="SA2E"
    $aSystemArray[2] ="a2iR"
    $aSystemArray[3] ="FA2I"
    $aSystemArray[4] ="TG27"
    $aSystemArray[5] ="a3H_"
    $aSystemArray[6] ="a3I_"
    $aSystemArray[7] ="G3T_"

    $sTmpath = $sTprjPath
    $sTmpath=StringTrimRight($sTmpath,20)
    $hFile = FileOpen($sTmpath&"makefile",$FO_READ)
    $sSystm = FileReadLine($hFile,26) ;Store into array
    ;Additional Checking if matches into folder file name

    For $i =0 To 7;Loop in the elements of system array list
        If StringInStr($sSystm, $aSystemArray[$i] ) Then
            _JPL_jnknsCreatelogfile('PL Error', "", 'Test :Checked MakeFile :,'& 'Yes', "= Failed ")
            Return $aSystemArray[$i]
        Else
            _JPL_jnknsCreatelogfile('PL Error', "", 'Test :Checked MakeFile : Not Found', 'Yes', "= Failed ")
        EndIf
    Next

EndFunc ;==>_JPE_jnkns_getSystmkfile

; #FUNCTION# ===========================================================================================================
; Name					:	_JPE_jnkns_processPLfile
; Description		:	Change PL files according to target system unit
; Syntax				:	_JPE_jnkns_processPLfile()
; Parameters		:	None
; Requirement(s)	:	v3.3 +
; Return values		: 	None
; Author				:	rdbayanado
; Modified			:	None
;=====================================================================================================================
Func _JPE_jnkns_processPLfile()
    Local $sSystemType, _
            $sSourcePath, _
            $sFileToCopy
    Local $siSMatch, _
            $sRet

    $sRet = 0
    $sSystemType=_JPE_jnkns_GetSystemtype()

    Local $string2="System Type: "&$sSystemType
    $sSourcePath =StringTrimRight($sTprjPath,21)
    $sSourcePath&="\Tools\SetWinAmsSpmcCode.pl"

	If $sSystemType == "a2ei" Or $sSystemType == "SA2E" Then
        $sFileToCopy=@ScriptDir & "\PL_Files\a2-e スズキ(Suzuki)\Tools\SetWinAmsSpmcCode.pl"
        If _JPE_jnkns_BinaryCompare($sSourcePath, $sFileToCopy) == "False" Then
            FileSetAttrib ( $sSourcePath, "-R" )
            FileCopy($sFileToCopy,$sSourcePath,$FC_OVERWRITE)
			_JPL_jnknsCreatelogfile('PL Error', '', 'Test : Updating PL File', 'Yes', "= Passed")
			_JPL_jnknsCreatelogfile('PL Error', '', 'Replaced PL File : a2-e スズキ(Suzuki)', 'Yes', @CRLF & @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
            $sRet = 1
        EndIf
    ElseIf $sSystemType=="a2iR" Then
		$sFileToCopy=@ScriptDir & "\PL_Files\a2i_Rivian\Tools\SetWinAmsSpmcCode.pl"
        If _JPE_jnkns_BinaryCompare($sSourcePath, $sFileToCopy) == "False" Then
            FileSetAttrib ( $sSourcePath, "-R" )
            FileCopy($sFileToCopy,$sSourcePath,$FC_OVERWRITE)
			_JPL_jnknsCreatelogfile('PL Error', '', 'Test : Updating PL File', 'Yes', "= Passed")
			_JPL_jnknsCreatelogfile('PL Error', '', 'Replaced PL File : a2i_Rivian', 'Yes', @CRLF & @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
            $sRet = 1
        EndIf
    ElseIf $sSystemType=="FA2I" Then
		$sFileToCopy=@ScriptDir & "\PL_Files\a2用\Tools\SetWinAmsSpmcCode.pl"
        If _JPE_jnkns_BinaryCompare($sSourcePath, $sFileToCopy) == "False" Then
            FileSetAttrib ( $sSourcePath, "-R" )
            FileCopy($sFileToCopy,$sSourcePath,$FC_OVERWRITE)
			_JPL_jnknsCreatelogfile('PL Error', '', 'Test : Updating PL File', 'Yes', "= Passed")
			_JPL_jnknsCreatelogfile('PL Error', '', 'Replaced PL File : a2用', 'Yes', @CRLF & @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
            $sRet = 1
        EndIf
	ElseIf $sSystemType =="TG27" Or $sSystemType =="TG25" Then
		$sFileToCopy=@ScriptDir & "\PL_Files\G27トヨタ (Toyota)\Tools\SetWinAmsSpmcCode.pl"
        If _JPE_jnkns_BinaryCompare($sSourcePath, $sFileToCopy) == "False" Then
            FileSetAttrib ( $sSourcePath, "-R" )
            FileCopy($sFileToCopy,$sSourcePath,$FC_OVERWRITE)
			_JPL_jnknsCreatelogfile('PL Error', '', 'Test : Updating PL File', 'Yes', "= Passed")
			_JPL_jnknsCreatelogfile('PL Error', '', 'Replaced PL File : G27トヨタ (Toyota)', 'Yes', @CRLF & @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
            $sRet = 1
        EndIf
	ElseIf $sSystemType=="a3H_" Then
		$sFileToCopy=@ScriptDir & "\PL_Files\a3 ホンダ (Honda)\Tools\SetWinAmsSpmcCode.pl"
        If _JPE_jnkns_BinaryCompare($sSourcePath, $sFileToCopy) == "False" Then
            FileSetAttrib ( $sSourcePath, "-R" )
            FileCopy($sFileToCopy,$sSourcePath,$FC_OVERWRITE)
			_JPL_jnknsCreatelogfile('PL Error', '', 'Test : Updating PL File', 'Yes', "= Passed")
			_JPL_jnknsCreatelogfile('PL Error', '', 'Replaced PL File : a3 ホンダ (Honda)', 'Yes', @CRLF & @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
            $sRet = 1
        EndIf
	ElseIf $sSystemType=="a3I_" Then
		$sFileToCopy=@ScriptDir & "\PL_Files\a3(ISUZU)\Tools\SetWinAmsSpmcCode.pl"
        If _JPE_jnkns_BinaryCompare($sSourcePath, $sFileToCopy) == "False" Then
            FileSetAttrib ( $sSourcePath, "-R" )
            FileCopy($sFileToCopy,$sSourcePath,$FC_OVERWRITE)
			_JPL_jnknsCreatelogfile('PL Error', '', 'Test : Updating PL File', 'Yes', "= Passed")
			_JPL_jnknsCreatelogfile('PL Error', '', 'Replaced PL File : a3(ISUZU)', 'Yes', @CRLF & @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
            $sRet = 1
        EndIf
    ElseIf $sSystemType=="G3T_" Then
		$sFileToCopy=@ScriptDir & "\PL_Files\G3 トヨタ (Toyota)\Tools\SetWinAmsSpmcCode.pl"
        If _JPE_jnkns_BinaryCompare($sSourcePath, $sFileToCopy) == "False" Then
            FileSetAttrib ( $sSourcePath, "-R" )
            FileCopy($sFileToCopy,$sSourcePath,$FC_OVERWRITE)
			_JPL_jnknsCreatelogfile('PL Error', '', 'Test : Updating PL File', 'Yes', "= Passed")
			_JPL_jnknsCreatelogfile('PL Error', '', 'Replaced PL File : G3 トヨタ (Toyota)', 'Yes', @CRLF & @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
            $sRet = 1
        EndIf
	Else
        _JPL_jnknsCreatelogfile('PL Error', "", 'Test : Replaced PL File : Unknown System Type', 'Yes', "= Failed")
	EndIf
    Return $sRet

EndFunc ;==>_JPE_jnkns_processPLfile