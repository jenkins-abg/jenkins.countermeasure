#cs	==================================================================================================================
	Title 				:	jnknsAssemblyError
	AutoIt Version	: 	3.3.14.5
	Language		: 	English
	Description		:	Fix errors regarding Assembly codes
	Author				: 	cjhernandez
    Modified by    :    prdedumo
    Version            :    0.1
#ce	==================================================================================================================

#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <File.au3>
#include <Array.au3>
#include <String.au3>
#include <AutoItConstants.au3>
#Include <Misc.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>

#include "..\jnknsMainInitialize.au3"
#include "..\ErrHandler\jnknsErrHandler.au3"
#include "..\TraceLog\jnknsProcessLogger.au3"

Global Enum $HAN_GUI, $HAN_TREE, $HAN_BTN, $HAN_BTN2, $HAN_COUNT
Global	$g_iJM_Handles[$HAN_COUNT], _
            $g_iJM_Spider_F5_Class , _				        ;	Class value depending in the AutoIt v3 Window info
            $g_iJM_Spider_File_Class, _				        ;	Class value depending in the AutoIt v3 Window info
            $g_iJM_Spider_Software_Path_Class , _	;	Class value depending in the AutoIt v3 Window info
            $g_iJM_Spider_Run_Class, _				    ;	Class value depending in the AutoIt v3 Window info
            $g_iJM_Assembly_Error, _
            $g_iJM_Assembly_TPRJ

Global  $g_JPL_txtfile = @ScriptDir & '\..\TraceLog\jenkins-trace-log.txt'
Global  $g_iJM_getWin
Global  $g_iJM_errorlogFile

Local   $sTextClasses, _
            $sSpider_Run_Class = "テスト実行中", _
            $sSpider_Local
Local   $sCount_txtfile = @ScriptDir & '\..\TraceLog\3.txt'

$g_iJM_Assembly_Error  = 0
; Initialize FSUnit Title
_JMI_jnknsCallDSpider()
_JMI_jnknsSpiderSettings()

Sleep(1000)

$sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle($g_sJMI_Spider_Version))
if _JMI_jnknsBuildTree($sTextClasses) Then
	$g_iJM_Assembly_Error = 1
	_AE_jnknsCheckAssembly()
EndIf
Exit

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_AE_jnknsCheckSource
; Description	:	Retrieve source name
; Author			:	cjhernandez
; Remarks			:
; ====================================================================================================================
Func _AE_jnknsCheckSource()
	Local $errorlogFile = @ScriptDir & "\..\Log.txt"
    Local $sSourceName, $sSource

    Local $hTextFile = FileOpen ( $errorlogFile, $FO_READ )
    If $hTextFile = -1 Then
        _JPL_jnknsCreatelogfile('Assembler Error', "", 'Error: Cannot read log file', 'No', "")
        Exit
    EndIf
	Local $errContent = FileReadToArray  ( $errorlogFile )
	Local $lineCounter = @extended
    If @error Then
        _JPL_jnknsCreatelogfile('Assembler Error', "", 'Error in reading log file in _AE_jnknsCheckSource()', 'No', "")
        Exit
    EndIf
    ; Test Sheet Line number
    $sTestSheetFile = FileReadLine($hTextFile,2)
    $sTestSheetFile = StringTrimLeft($sTestSheetFile,20)
    $g_sJMI_TestDesign_File = $sTestSheetFile
	For $i = 0 To $lineCounter - 1
		$sSourceName = _StringBetween($errContent[$i], "【", "】")
		If UBound($sSourceName) > 0 Then
			$sSourceName = _StringBetween($errContent[$i], "【", "】")
			$sSource = String($sSourceName[0])
			ExitLoop
		EndIf
	Next
	return $sSource
EndFunc ;==>_AE_jnknsCheckSource

; #INTERNAL_USE_ONLY# ================================================================================================
; Name				:	_AE_jnknsCheckAssembly
; Description	:	Check and edit assembly code in source file
; Author			:	cjhernandez
; Remarks			:
; ====================================================================================================================
Func _AE_jnknsCheckAssembly()
	Local $filepath = StringTrimRight($g_iJM_Assembly_TPRJ , 21)
	Local $source = _FO_FileSearch($filepath, _AE_jnknsCheckSource(), 1)
	Local $fileArrayA, $fileArrayB
	Local $counter
	Local $sLineStringEdited = ""

    If Not _FileReadToArray( $source[1], $fileArrayA) Then
        _JPL_jnknsCreatelogfile('Assembler Error', "", 'Error in reading log file in _AE_jnknsCheckAssembly()', 'No', "")
        Exit
    EndIf
    $counter = 0
    ; Initialize Countermeasure
    _JPL_jnknsCreatelogfile('Assembler Error', $g_sJMI_TestDesign_File, 'Test : Editing definition', 'Yes', "start")
	For $a = 1 To $fileArrayA[0]
		
		sLineStringEdited = ""
		If StringInStr($fileArrayA[$a], 'AD_SYNCP') Then
			_FileReadToArray( $filepath & "/ad_drv_st.h", $fileArrayB)
			For $b = 1 To $fileArrayB[0]
				FileSetAttrib( $filepath & "/ad_drv_st.h", "-R" )
				$replaceSyncp = StringReplace( FileReadLine($filepath & "/ad_drv_st.h", $b), '"syncp"', '"nop"', 0)
				_FileWriteToLine( $filepath & "/ad_drv_st.h", $b, $replaceSyncp, 1 )
				$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
			Next

			_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
			_JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing AD_SYNCP definition', 'Yes', "= Passed")
			$counter = $counter + 1
		EndIf

		sLineStringEdited = ""
		If StringInStr($fileArrayA[$a], 'ADCHK_SYNCP') Then
			_FileReadToArray( $filepath & "/cpuadc_pmchk_st.h", $fileArrayB)
			For $b = 1 To $fileArrayB[0]
				FileSetAttrib( $filepath & "/cpuadc_pmchk_st.h", "-R" )
				$replaceSyncp = StringReplace( FileReadLine($filepath & "/cpuadc_pmchk_st.h", $b), '"syncp"', '"nop"', 0)
				_FileWriteToLine( $filepath & "/cpuadc_pmchk_st.h", $b, $replaceSyncp, 1 )
				$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
			Next
			_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
			_JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing ADCHK_SYNCP definition', 'Yes', "= Passed")
			$counter = $counter + 1
		EndIf

		sLineStringEdited = ""
		If _AE_jnknsCheckSource() = "spih1_drv.c" Then
			If StringInStr($fileArrayA[$a], 'SPI_SYNCP') Then
				_FileReadToArray( $filepath & "/spih1_drv_st.h", $fileArrayB)
				For $b = 1 To $fileArrayB[0]
					FileSetAttrib( $filepath & "/spih1_drv_st.h", "-R" )
					$replaceSyncp = StringReplace( FileReadLine($filepath & "/spih1_drv_st.h", $b), '"syncp"', '"nop"', 0)
					_FileWriteToLine( $filepath & "/spih1_drv_st.h", $b, $replaceSyncp, 1 )
					$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
				Next
				_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
                _JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing SPI_SYNCP definition', 'Yes', "= Passed")
                $counter = $counter + 1
			EndIf
		EndIf

		sLineStringEdited = ""
		If _AE_jnknsCheckSource() = "spih2_drv.c" Then
			If StringInStr($fileArrayA[$a], 'SPI_SYNCP') Then
				_FileReadToArray( $filepath & "/spih2_drv_st.h", $fileArrayB)
				For $b = 1 To $fileArrayB[0]
					FileSetAttrib( $filepath & "/spih2_drv_st.h", "-R" )
					$replaceSyncp = StringReplace( FileReadLine($filepath & "/spih2_drv_st.h", $b), '"syncp"', '"nop"', 0)
					_FileWriteToLine( $filepath & "/spih2_drv_st.h", $b, $replaceSyncp, 1 )
					$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
				Next
				_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
                _JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing SPI_SYNCP definition', 'Yes', "= Passed")
                $counter = $counter + 1
			EndIf
		EndIf

		sLineStringEdited = ""
		If _AE_jnknsCheckSource() = "spih3_drv.c" Then
			If StringInStr($fileArrayA[$a], 'SPI_SYNCP') Then
				_FileReadToArray( $filepath & "/spih3_drv_st.h", $fileArrayB)
				For $b = 1 To $fileArrayB[0]
					FileSetAttrib( $filepath & "/spih3_drv_st.h", "-R" )
					$replaceSyncp = StringReplace( FileReadLine($filepath & "/spih3_drv_st.h", $b), '"syncp"', '"nop"', 0)
					_FileWriteToLine( $filepath & "/spih3_drv_st.h", $b, $replaceSyncp, 1 )
					$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
				Next
				_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
                _JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing SPI_SYNCP definition', 'Yes', "= Passed")
                $counter = $counter + 1
			EndIf
		EndIf

		sLineStringEdited = ""
		If _AE_jnknsCheckSource() = "spih4_drv.c" Then
			If StringInStr($fileArrayA[$a], 'SPI_SYNCP') Then
				_FileReadToArray( $filepath & "/spih4_drv_st.h", $fileArrayB)
				For $b = 1 To $fileArrayB[0]
					FileSetAttrib( $filepath & "/spih4_drv_st.h", "-R" )
					$replaceSyncp = StringReplace( FileReadLine($filepath & "/spih4_drv_st.h", $b), '"syncp"', '"nop"', 0)
					_FileWriteToLine( $filepath & "/spih4_drv_st.h", $b, $replaceSyncp, 1 )
					$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
				Next
				_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
                _JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing SPI_SYNCP definition', 'Yes', "= Passed")
                $counter = $counter + 1
			EndIf
		EndIf

		sLineStringEdited = ""
		If _AE_jnknsCheckSource() = "spi1_drv.c" Then
			If StringInStr($fileArrayA[$a], 'SPI_SYNCP') Then
				_FileReadToArray( $filepath & "/spi1_drv_st.h", $fileArrayB)
				For $b = 1 To $fileArrayB[0]
					FileSetAttrib( $filepath & "/spi1_drv_st.h", "-R" )
					$replaceSyncp = StringReplace( FileReadLine($filepath & "/spi1_drv_st.h", $b), '"syncp"', '"nop"', 0)
					_FileWriteToLine( $filepath & "/spi1_drv_st.h", $b, $replaceSyncp, 1 )
					$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
				Next
				_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
                _JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing SPI_SYNCP definition', 'Yes', "= Passed")
                $counter = $counter + 1
            EndIf
		EndIf

		sLineStringEdited = ""
		If _AE_jnknsCheckSource() = "spi2_drv.c" Then
			If StringInStr($fileArrayA[$a], 'SPI_SYNCP') Then
				_FileReadToArray( $filepath & "/spi2_drv_st.h", $fileArrayB)
				For $b = 1 To $fileArrayB[0]
					FileSetAttrib( $filepath & "/spi2_drv_st.h", "-R" )
					$replaceSyncp = StringReplace( FileReadLine($filepath & "/spi2_drv_st.h", $b), '"syncp"', '"nop"', 0)
					_FileWriteToLine( $filepath & "/spi2_drv_st.h", $b, $replaceSyncp, 1 )
					$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
				Next
				_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
                _JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing SPI_SYNCP definition', 'Yes', "= Passed")
                $counter = $counter + 1
			EndIf
		EndIf

		sLineStringEdited = ""
		If _AE_jnknsCheckSource() = "spi3_drv.c" Then
			If StringInStr($fileArrayA[$a], 'SPI_SYNCP') Then
				_FileReadToArray( $filepath & "/spi3_drv_st.h", $fileArrayB)
				For $b = 1 To $fileArrayB[0]
					FileSetAttrib( $filepath & "/spi3_drv_st.h", "-R" )
					$replaceSyncp = StringReplace( FileReadLine($filepath & "/spi3_drv_st.h", $b), '"syncp"', '"nop"', 0)
					_FileWriteToLine( $filepath & "/spi3_drv_st.h", $b, $replaceSyncp, 1 )
					$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
				Next
				_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
                _JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing SPI_SYNCP definition', 'Yes', "= Passed")
                $counter = $counter + 1
			EndIf
		EndIf

		sLineStringEdited = ""
		If _AE_jnknsCheckSource() = "spi4_drv.c" Then
			If StringInStr($fileArrayA[$a], 'SPI_SYNCP') Then
				_FileReadToArray( $filepath & "/spi4_drv_st.h", $fileArrayB)
				For $b = 1 To $fileArrayB[0]
					FileSetAttrib( $filepath & "/spi4_drv_st.h", "-R" )
					$replaceSyncp = StringReplace( FileReadLine($filepath & "/spi4_drv_st.h", $b), '"syncp"', '"nop"', 0)
					_FileWriteToLine( $filepath & "/spi4_drv_st.h", $b, $replaceSyncp, 1 )
					$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
				Next
				_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
                _JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing SPI_SYNCP definition', 'Yes', "= Passed")
                $counter = $counter + 1
			EndIf
		EndIf

		sLineStringEdited = ""
		If StringInStr($fileArrayA[$a], 'SYNCP') OR StringInStr($fileArrayA[$a], 'SYNCI')  Then
			_FileReadToArray( $filepath & "/system.h", $fileArrayB)
			For $b = 1 To $fileArrayB[0]
				FileSetAttrib( $filepath & "/system.h", "-R" )
				$replaceSyncp = StringReplace( FileReadLine($filepath & "/system.h", $b), '"syncp"', '"nop"', 0)
				_FileWriteToLine( $filepath & "/system.h", $b, $replaceSyncp, 1 )
                $replaceSynci = StringReplace( FileReadLine($filepath & "/system.h", $b), '"synci"', '"nop"', 0)
				_FileWriteToLine( $filepath & "/system.h", $b, $replaceSynci, 1 )
				$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
			Next
			_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
			_JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing SYNCP and/or SYNCI  definition', 'Yes', "= Passed")
			$counter = $counter + 1
		EndIf

		sLineStringEdited = ""
		If StringInStr($fileArrayA[$a], 'SYSDRV_SYNCP') Or StringInStr($fileArrayA[$a], 'SYSDRV_SYNCI')  Then
			_FileReadToArray( $filepath & "/system_drv_gl.h", $fileArrayB)
			For $b = 1 To $fileArrayB[0]
				FileSetAttrib( $filepath & "/system_drv_gl.h", "-R" )
				$replaceSyncp = StringReplace( FileReadLine($filepath & "/system_drv_gl.h", $b), '"syncp"', '"nop"', 0)
				_FileWriteToLine( $filepath & "/system_drv_gl.h", $b, $replaceSyncp, 1 )
                $replaceSynci = StringReplace( FileReadLine($filepath & "/startup_st.h", $b), '"synci"', '"nop"', 0)
				_FileWriteToLine( $filepath & "/system_drv_gl.h", $b, $replaceSynci, 1 )
				$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
			Next
			_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
			_JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing SYSDRV_SYNCP and/or SYSDRV_SYNCI  definition', 'Yes', "= Passed")
			$counter = $counter + 1
		EndIf

		sLineStringEdited = ""
		If StringInStr($fileArrayA[$a], 'TAU_SYNCP') Then
			_FileReadToArray( $filepath & "/timer_drv_st.h", $fileArrayB)
			For $b = 1 To $fileArrayB[0]
				FileSetAttrib( $filepath & "/timer_drv_st.h", "-R" )
				$replaceSyncp = StringReplace( FileReadLine($filepath & "/timer_drv_st.h", $b), '"syncp"', '"nop"', 0)
				_FileWriteToLine( $filepath & "/timer_drv_st.h", $b, $replaceSyncp, 1 )
				$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
			Next
			_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
			_JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing TAU_SYNCP  definition', 'Yes', "= Passed")
			$counter = $counter + 1
		EndIf

		sLineStringEdited = ""
		If StringInStr($fileArrayA[$a], 'ACTIVATE_TASK')  Or StringInStr($fileArrayA[$a], 'TERMINATE_TASK')  Or StringInStr($fileArrayA[$a], 'GET_TASK_ID')  Or StringInStr($fileArrayA[$a], 'GET_TASK_STATE')  Or _
        StringInStr($fileArrayA[$a], 'SET_EVENT')  Or StringInStr($fileArrayA[$a], 'CLEAR_EVENT')  Or StringInStr($fileArrayA[$a], 'GET_EVENT')  Or StringInStr($fileArrayA[$a], 'WAIT_EVENT')  Or StringInStr($fileArrayA[$a], 'SET_REL_ALARM') Then
			_FileReadToArray( $filepath & "/system_drv_gl.h", $fileArrayB)
			For $b = 1 To $fileArrayB[0]
				FileSetAttrib( $filepath & "/system_drv_gl.h", "-R" )
                $replaceSyncp = StringReplace( FileReadLine($filepath & "/system_drv_gl.h", $b), '"SYNCP"', '"NOP"', 0)
                _FileWriteToLine( $filepath & "/system_drv_gl.h", $b, $replaceSyncp, 1 )
                $replaceSynci = StringReplace( FileReadLine($filepath & "/system_drv_gl.h", $b), '"SYNCI"', '"NOP"', 0)
                _FileWriteToLine( $filepath & "/system_drv_gl.h", $b, $replaceSynci, 1 )
                $replaceSyncall = StringReplace( FileReadLine($filepath & "/system_drv_gl.h", $b), '"SYSCALL 0"', '"NOP"', 0)
				_FileWriteToLine( $filepath & "/system_drv_gl.h", $b, $replaceSyncall, 1 )
				$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
			Next
			_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
        	_JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing **_TASK, **_EVENT, SET_REL_ALARM  definitions', 'Yes', "= Passed")
        $counter = $counter + 1
		EndIf

		sLineStringEdited = ""
		If StringInStr($fileArrayA[$a], 'asm(') Then
            If StringInStr($fileArrayA[$a], 'asm("nop")') Or StringInStr($fileArrayA[$a], 'asm("NOP")') Then
            Else
                FileSetAttrib( $source, "-R" )
                $replaceAsmF = StringReplace( FileReadLine( $source, $a ), 'asm(', '/* asm(', 0 )
                _FileWriteToLine( $source, $a, $replaceAsmF, 1 )
                $replaceAsmL = StringReplace( FileReadLine( $source, $a ), ';', '; */', 0 )
				_FileWriteToLine( $source, $a, $replaceAsmL, 1 )
				$sLineStringEdited = $sLineStringEdited & "Edited line code number: " & $b & " in " & $fileArrayA[$a] & @CRLF
                $counter = $counter + 1
			EndIf
			_JPL_jnknsCreatelogfile('Assembler Error', '', $sLineStringEdited, 'Yes', @TAB & @TAB  & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & @TAB & "STATUS : OK")
			_JPL_jnknsCreatelogfile('Assembler Error', "", 'Test : Editing ' & $source & ' definitions', 'Yes', "= Passed")
		EndIf
	Next

    If $counter > 0 Then
        Local $hFileOpen = FileOpen($sCount_txtfile, 2)
            If $hFileOpen = -1 Then
                Exit
            EndIf
        FileWriteLine($hFileOpen,"Assembler countermeasure applied")
	EndIf
	
    _JPL_jnknsCreatelogfile('CPU_Emergency Error', "", 'Exiting countermeasure', 'Yes', 'End')
#cs ========================================================
    This was commented out since this is a pre-run countermeasure
    This will be run together after applying/checking the 3 pre-run countermeasures
    ; Re-run the sheet
	_JMI_jnknsCallDSpider()
	WinActivate($g_sJMI_Spider_Version)
	MouseClick("Left",609, 299)
	; Loop to wait until running of the tool is done
	$sSpider_Local = WinActivate($sSpider_Run_Class)
	While 1
		$sSpider_Local = WinActivate($sSpider_Run_Class)
		if $sSpider_Local <> 0 Then
		Else
			ExitLoop
		EndIf
	WEnd
	Sleep(2000)
#ce ========================================================
EndFunc ;==>_AE_jnknsCheckAssembly
