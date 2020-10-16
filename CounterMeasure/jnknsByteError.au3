#cs	==================================================================================================================
	Title 				:	jnknsByteError
	AutoIt Version	: 	3.3.14.5
	Language		: 	English
	Description		:	Fix errors regarding exceeding function call in target unit
	Author	/s			: 	cjhernandez
                            :   rdbayanado
    Version            :    0.1
#ce	==================================================================================================================

#include <Excel.au3>
#include <Array.au3>
#include <File.au3>
#include <String.au3>
#include <GuiTab.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>

#include "..\Initializer.au3"
#include "..\TraceLog\jnknsProcessLogger.au3"

; Global Variables
Global $sLogTextFile = @ScriptDir & "\..\Log.txt"

; Local Variables
Local	$sErrNumber, _
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
            $sSpider_Log_TxtFile
Local	$sTextFile
Local	$i
Local	$aArray, _
			$aPathSplit
Local	$iBackUpResult, _
			$iCopyResult, _
			$iRebuildResult, _
			$iErrNumber

Global  $g_JPL_txtfile = @ScriptDir & '\..\TraceLog\jenkins-trace-log.txt'
Global $IsByteError
; Open log text file
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


_JMI_jnknsCallDSpider()
$sTextClasses = _JMI_jnknsWinGetClassesByText(WinGetHandle($g_sJMI_Spider_Version))
if _JMI_jnknsBuildTree($sTextClasses) Then
    $sSpider_Software_Path = ControlGetText('ソフトウェア単体テスト自動化ツール D-SPIDER  Ver.1.0.0',"",$g_iJM_Spider_Software_Path_Class)
    $sSpider_Path =  StringTrimRight($sSpider_Software_Path,21)
    $sUnitTest_Log_TxtFile = $sSpider_Path & "\UnitTest\log.txt"
EndIf

Sleep(3000)
_JEH_RefreshSettings($sSpider_Path )
;MsgBox(0,"",$sTestSheetFile)
_JMI_jnknsPressF5($g_sJMI_Spider_Version)
;Send("{F5}")
WinWait("","",10)
$sSpider_Local = WinActivate($sSpider_Run_Class)
While 1
    $sSpider_Local = WinActivate($sSpider_Run_Class)
        if $sSpider_Local <> 0 Then
        Else
            ExitLoop
        EndIf
WEnd
Sleep(3000)
checkError()

If $IsByteError ==1 Then
    _JPL_jnknsCreatelogfile('Byte Error', $sTestSheetFile, 'Test : ByteError check...', 'Yes', "start")

    Sleep(10000)
    _BE_RunIFSheet(_BE_GetXCellPath())
    Sleep(10000)
    _JMI_jnknsCallDSpider()
    Sleep(3000)
    _JPL_jnknsCreatelogfile('Byte Error', "", 'Test : Separating Sheet', 'Yes', "= Passed")
    _JEH_RefreshSettings($sSpider_Path  & '\')
    _JMI_jnknsPressF5($g_sJMI_Spider_Version)
    ;Send("{F5}")
    Sleep(10000)
    ;$sSpider_Local = WinActivate($sSpider_Run_Class)
    Local $spiderHwnd = WinGetHandle($sSpider_Run_Class)
	While 1
		$spiderHwnd = WinGetHandle($sSpider_Run_Class)
            if $sSpider_Local <> 0 Then
            Else
                ExitLoop
            EndIf
    WEnd
    _JPL_jnknsCreatelogfile('Byte Error', "", 'Exiting countermeasure', 'Yes', 'End')
EndIf
FileClose($sTextFile)
Exit

Func checkError()
    Local $sLogfile
    Local $hTextFile
    Local $iLine

    $sLogfile=StringTrimRight($sTprjPath, 20)
    $sLogfile&="UnitTest\log.txt"
    $hTextFile=FileOpen($sLogfile,$FO_READ)
    $iLine=FileReadLine($hTextFile,7)
    If StringInStr($iLine, "スタブ関数の数が最大値(40)を超えました。")  Then
        $IsByteError=1
        _BE_jnknsMain( _BE_GetXCellPath())
        ;  MsgBox(0,"","Test1")
    ElseIf StringInStr($iLine, "スタブ機能のメモリ使用量が最大値(4096 byte)を超えました。")  Then
        $IsByteError=1
        _BE_jnknsMain( _BE_GetXCellPath())
    Else
        _JPL_jnknsCreatelogfile('Byte Error', "", 'Test : '&$iLine &' Error Detected ', 'Yes', "= Failed") ; Insert into log file if different error exist.
        $IsByteError=0
        ;MsgBox(0,"","Test2")
    EndIf
    FileClose($hTextFile)
EndFunc


; #INTERNAL_USE_ONLY# ================================================================================================
; Name...........: _BE_jnknsGetArrayData
; Description ...: Store strings of each worksheet into an array
; Author/s ........: rdbayanado
; Remarks .......:
; ====================================================================================================================
Func _BE_jnknsGetArrayData( $xCellFile,$iSheetnum)
    Local $oExcel, _
            $oWorkbook
    Local $iLastCol, _
            $iRangeLast, _
            $iRows, _
            $iCols, _
            $iWorksheets, _
            $iArrayNumCount
    Local $aRawArray, _
            $aSecArray[1][2000], _
            $aElements[0], _
            $aUniqElements[0], _
            $aTempArrayStack[0][0], _
            $aSheetList[0], _
            $aIsSeparate[0], _
            $sStrTypeValidator, _
            $sStrSubfuncValidator, _
            $aClearArray[1]=[""]
    $oExcel = _Excel_Open(False)
    $oWorkbook = _Excel_BookOpen($oExcel, $xCellFile)
    $iWorksheets=$oWorkbook.Worksheets.Count
    $aSheetList=_Excel_SheetList($oWorkbook)
    $iRangeLast = $oWorkbook.Worksheets($iSheetnum).UsedRange.SpecialCells($xlCellTypeLastCell) ; Get last cel used stored into variable
    $iLastCol=_Excel_ColumnToLetter($iRangeLast.Column)  ;Convert last cell to position coordinates in excel file
    $aRawArray = _Excel_RangeRead( $oWorkbook, $iSheetnum, "E27:"&$iLastCol&"27") ; Stored array into $aRawArray
    $iRows = UBound($aRawArray, $UBOUND_ROWS) ;Get number of rows stored into  $iRows
    $iCols = UBound($aRawArray, $UBOUND_COLUMNS);Get number of column stored into  $iCols

;~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~  Loop on the Extracted Array and Check if consist subfunctions~~
;~  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    For $i = 0 To $iRows - 1 ;Loop on Rows
        For $j = 0 To $iCols - 1;  Loop on column
            $sStrSubfuncValidator= StringInStr($aRawArray[$i][$j],")") ;Validate string as subfunc type
            If ($sStrSubfuncValidator<>0 ) Then ;Check if valid
                $sStrValidator=StringMid($aRawArray[$i][$j], $sStrSubfuncValidator+1,1)
                If $sStrValidator == "@" Then ;String Handler
                    $aSecArray[$i][$j]=StringTrimLeft($aRawArray[$i][$j],$sStrSubfuncValidator +1)
                ElseIf $sStrValidator == "[" Then ;String Handler
                    $aSecArray[$i][$j]=$aRawArray[$i][$j]
                EndIf;===>String Handler
            EndIf;==> ;Check if valid
        Next; ===> Loop on column
    Next;===>;Loop on Rows

;~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~                           Get Unique elements in the Array                    ~~
;~  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    For $i = 0 To $iRows - 1 ;Loop on Rows
        For $j = 0 To $iCols - 1;Loop on Columns
            If $j<>0 Then ;Check Counter of Column
                If  $aSecArray[$i][$j] <> $aSecArray[$i][$j-1] And  $aSecArray[$i][$j] <> " "  Then; Check if element is the same as its preceeding element
                    $sStrValidator= StringInStr($aSecArray[$i][$j],"@") ;Check subfunc attributes
                    If $sStrValidator <>0 Then
                        $aSecArray[$i][$j]=StringTrimLeft( $aSecArray[$i][$j],$sStrValidator)
                        If  $aSecArray[$i][$j] <> $aSecArray[$i][$j-1] And $aSecArray[$i][$j] <> " "  Then
                            If StringLen ($aSecArray[$i][$j]) <> 0 Then
                                _ArrayAdd($aElements, $aSecArray[$i][$j]);Append to $aElements as unique
                            EndIf
                        EndIf
                    Else
                        If StringLen ($aSecArray[$i][$j]) <> 0 Then
                        _ArrayAdd($aElements, $aSecArray[$i][$j]);Append to $aElements as unique
                        EndIf
                    EndIf
                EndIf;===>Check if element is the same as its preceeding element
            EndIf;==> ;Check Counter of Column
        Next;===>;Loop on Columns
    Next;==> ;Loop on Rows

    Local $tmparray[1] = ['']
    For $i = 0 To UBound($aElements) - 1 ;Loop for the current element in the array
        _ArraySearch($tmparray, $aElements[$i]) ;Search for the unique item in the array
        If @error Then _ArrayAdd($tmparray, $aElements[$i]);Append array if unique
    Next;;===Loop for the current element in the array

    For  $i = 0 To UBound($tmparray) -1
        If StringLen($tmparray[$i]) <> 0 Then
            _ArrayAdd($aUniqElements, $tmparray[$i])
        EndIf
    Next
    Return UBound($aUniqElements) -1
EndFunc

; #INTERNAL_USE_ONLY# ================================================================================================
; Name...........: _BE_jnknsSeparateSheet
; Description ...: Separate sheet which exceeds subfunction call
; Author/s ........: rdbayanado
;                ........: cjhernandez
; Remarks .......:
; ====================================================================================================================
Func _BE_jnknsSeparateSheet( $ftestDesign,$sheetindex,$sheetindex2,$aArraySheets,$iSheetnum) ; Separate sheet which exceeds subfunction call
    Local $oExcel, _
            $oFtestDesign, _
            $oBook
    Local  $aTobeSeparated, _
            $aSheetList
    Local  $sSheetname, _
            $sStringCount, _
            $sStringAppend1, _
            $sStringAppend2
    Local $fNewSheet, _
            $fIFSheet

    $oExcel=_Excel_Open(False)
    $oBook = _Excel_BookNew( $oExcel, 1 )
    $oftestDesign = _Excel_BookOpen ( $oExcel, $ftestDesign )

    If StringInStr($fTestDesign, "Rev") Then ;Check String File as validation for Revision String
        If StringInStr($fTestDesign,"No.") Then
            $sStringAppend1 = StringTrimRight( $FtestDesign,17)
            $sStringCount=StringLen($sStringAppend1);Get String Count of filename minus the file extension
            $sStringCount+=4
            $sStringAppend1=$sStringAppend1& "_"
        Else
            $sStringAppend1 = StringTrimRight( $FtestDesign,13)
                $sStringCount=StringLen($sStringAppend1);Get String Count of filename minus the file extension
        EndIf
        $sStringAppend2=StringTrimLeft($FtestDesign,$sStringCount); Set value of string 2 to concatenated string
        $fNewSheet= $sStringAppend1&"No."&$iSheetnum& $sStringAppend2
    Else ; Case Scenario if New Test Design
        If StringInStr($fTestDesign,"No.") Then
            $sStringAppend1 = StringTrimRight( $FtestDesign,17)
            $sStringCount=StringLen($sStringAppend1);Get String Count of filename minus the file extension
            $sStringCount+=4
            $sStringAppend1=$sStringAppend1& "_"
        Else
            $sStringAppend1 = StringTrimRight( $FtestDesign,13)
            $sStringCount=StringLen($sStringAppend1);Get String Count of filename minus the file extension
        EndIf
        $sStringAppend2=StringTrimLeft($FtestDesign,$sStringCount); Set value of string 2 to concatenated string
        $fNewSheet= $sStringAppend1&"_"&"No."&$iSheetnum& $sStringAppend2

    EndIf ;===>  ;Check String File as validation for Revision String

    _Excel_BookSaveAs( $oBook, $fNewSheet, $xlWorkbookDefault, True )
    $fNewSheet = _Excel_BookOpen ( $oExcel, $fNewSheet )
    _Excel_SheetCopyMove ( $oftestDesign, UBound($aArraySheets)-$iSheetnum, $fNewSheet, 1, True )
    If $sheetindex<> 0 Then ; IF argument is set to default then  perform
        _Excel_SheetCopyMove ( $oftestDesign, $sheetindex, $fNewSheet, 1, True )
    EndIf
    _Excel_SheetCopyMove ( $oftestDesign, $sheetindex2, $fNewSheet, 1, True )
    _Excel_SheetCopyMove ( $oftestDesign, 2, $fNewSheet, 1, True )
    _Excel_SheetCopyMove ( $oftestDesign, 1, $fNewSheet, 1, True )
    _Excel_SheetDelete ( $fNewSheet, "Sheet1" )

    $fNewSheet.Theme.ThemeColorScheme.Load ("C:\Program Files (x86)\Microsoft Office\Document Themes 16\Theme Colors\Office 2007 - 2010.xml")
;~     WinActivate
    _Excel_BookClose ( $ftestDesign )
    _Excel_BookClose ( $fNewSheet )
    _Excel_Close($oExcel)
    If $sheetindex >=3 Then
        _BE_jnkns_SheetDelete($ftestDesign,$sheetindex )
    EndIf
    If $sheetindex2 >=3 Then
        _BE_jnkns_SheetDelete($ftestDesign,$sheetindex2 )
    EndIf
EndFunc;==>  ; Separate sheet which exceeds subfunction call

; #INTERNAL_USE_ONLY# ================================================================================================
; Name...........: _BE_jnknsMain
; Description ...: Decides whether the sheet needs to be separated or not
; Author/s ........: rdbayanado
; Remarks .......:
; ====================================================================================================================
Func _BE_jnknsMain($xCellFile);Decides whether the sheet needs to be separated or not
    Local $iWorksheets, _
            $iLoopCount, _
            $iSheet, _
            $iRSCount
    Local $aWorkSheetList
    Local $sSearchString, _
            $sDirection, _
            $sPreviousLoopSheet, _
            $sPrevious, _
            $sNext, _
            $sCurrent

    $oExcel = _Excel_Open(False) ;Instantiatiate Excel application
    $oWorkbook = _Excel_BookOpen($oExcel, $xCellFile, True, False)
    $iWorksheets =$oWorkbook.worksheets.Count ;Worksheets count
    $aWorkSheetList =_Excel_SheetList($oWorkbook) ;Store
    $iLoopCount=0 ;loop counter
    $iRSCount= _BE_jnkns_IsSepratedOrNot($xCellFile)
    
    For $i =3 To ($iWorksheets-1)-$iRSCount
        $iSheet =  _BE_jnknsGetArrayData($xCellFile,$i-$iLoopCount)
        If $iSheet >= 41 And ($iWorksheets-$iRSCount)-1 > 5 Then;Check if Exceeds in  subfunction limit
            $sPrevious = $aWorkSheetList[$i-2][0]
            $sNext = $aWorkSheetList[$i][0]
            $sCurrent =$aWorkSheetList[$i-1][0]
    
            If $iLoopCount == 0 Then ;Check variable value for iteration number
                If StringInStr($sCurrent, "IF") Then ;Check If current sheetname is main Sheet or IF sheet
                    $sSearchString= StringLeft($sCurrent,5)
                    $sDirection=1
                Else
                    $sSearchString=$sCurrent
                    $sDirection=0
                EndIf;==>;Check If current sheetname is main Sheet or IF sheet
    
                If $sDirection== 1 Then ;Condition   varies set to left direction set course
                    If StringInStr($sPrevious, $sSearchString) Then ;
                        _BE_jnknsSeparateSheet($xCellFile,$i, $i-1,$aWorkSheetList, $iLoopCount+1)
                        $iLoopCount +=1;===>
                        ;MsgBox(0,"","Case1"&"Sheet"&$i&$i-1)
                    Else
                        _BE_jnknsSeparateSheet($xCellFile,0, $i-1,$aWorkSheetList, $iLoopCount+1)
                        $iLoopCount +=1;===>
                        ;MsgBox(0,"","Case2"&"Sheet"&$i-1)
                    EndIf

                ElseIf $sDirection==0 Then;Condition   varies set to right  direction
                    If StringInStr($sNext, $sSearchString) Then
                        _BE_jnknsSeparateSheet($xCellFile,$i+1,$i,$aWorkSheetList, $iLoopCount+1)
                        $iLoopCount +=1
                        ;MsgBox(0,"","Case3"&"Sheet"&$i+1&$i)
                    Else
                        _BE_jnknsSeparateSheet($xCellFile,0,$i,$aWorkSheetList, $iLoopCount+1)
                        $iLoopCount +=1
                        ; MsgBox(0,"","Case4"&"Sheet"&$i)
                    EndIf ;===>Condition   varies set to left direction set course
                EndIf
            Else
            ;Subtract Loop Count in the Sheet count since it affect the placing of the sheets.
                If  StringInStr( _BE_jnkns_ChkIFSheet($sCurrent), _BE_jnkns_ChkIFSheet($sPreviousLoopSheet)) Then ;Check if Current sheet was identical to the previous sheet
                Else
                    If StringInStr($sCurrent, "IF") Then ;
                        $sSearchString= StringLeft($sCurrent,5)
                        $sDirection=1
                    Else
                        $sSearchString=$sCurrent ;Set to default
                        $sDirection=0 ;Set Direction parameter to Right  direction
                    EndIf;==>;Check If current sheetname is main Sheet or IF sheet

                    If $sDirection== 1 Then ;Condition   varies set to left direction set course
                        If StringInStr($sPrevious, $sSearchString) Then
                            _BE_jnknsSeparateSheet($xCellFile,$i-$iLoopCount, ($i-1)-$iLoopCount,$aWorkSheetList, $iLoopCount+1)
                            $iLoopCount +=1;===>
                            ;MsgBox(0,"","Case5"&"Sheet"&$i-$iLoopCount&($i-1)-$iLoopCount)
                        Else ;IF it deals with no IF Sheet only main sheet will be separate
                            _BE_jnknsSeparateSheet($xCellFile,0, ($i-1)-$iLoopCount,$aWorkSheetList, $iLoopCount+1)
                            $iLoopCount +=1;===>Check if Previous string match within the search string
                            ; MsgBox(0,"","Case6"&"Sheet"&($i-1)-$iLoopCount)
                        EndIf

                    ElseIf $sDirection==0 Then;Condition   varies set to right  direction
                        If StringInStr($sNext, $sSearchString) Then
                            _BE_jnknsSeparateSheet($xCellFile,($i+1)-$iLoopCount,$i-$iLoopCount,$aWorkSheetList, $iLoopCount+1)
                            ;MsgBox(0,"","Case7"&"Sheet"&($i+1)-$iLoopCount&$i-$iLoopCount)
                            $iLoopCount +=1
                        Else ;IF it deals with no IF Sheet only main sheet will be separate
                                _BE_jnknsSeparateSheet($xCellFile,0,$i-$iLoopCount,$aWorkSheetList, $iLoopCount+1)
                                ; MsgBox(0,"","Case8"&"Sheet"&$i-$iLoopCount)
                                $iLoopCount +=1
                        EndIf
                    EndIf ;===>Condition   varies set to left direction set course
                EndIf
            EndIf;==>;Check variable value for iteration number
        ElseIf $iSheet >= 41 And ($iWorksheets-$iRSCount)-1 ==3 Then ;if Single file only
            jnkns_BE_DeleteVar($xCellFile,3)
        ElseIf  $iSheet >= 41 And ($iWorksheets-$iRSCount)-1 == 4 Then ;IF Separated sheet needs additional countermeasure
            If StringInStr($aWorkSheetList[3][0],$aWorkSheetList[2][0])  Then
                _BE_jnknsSeparateSheet($xCellFile,0,$i,$aWorkSheetList, $iLoopCount+1)
            EndIf
        EndIf;==>;Check if Exceeds in  subfunction limit
        $sPreviousLoopSheet =$sCurrent ;Set value to current sheet in  loop
    Next;==>$iWorksheets
EndFunc;==>;Decides whether the sheet needs to be separated or not

; #INTERNAL_USE_ONLY# ================================================================================================
; Name...........: _BE_jnknsSeparateSheet
; Description ...:Delete VAR on the excel cells that exceeds subfunctions limit
; Author/s ........: rdbayanado
;                ........: cjhernandez
; Remarks .......:
; ================================================================================================================
Func jnkns_BE_DeleteVar($xCellFile,$iSheetnum) ; Delete Var Of Excess Functions
    Local $oExcel, _
            $oWorkbook
    Local $iLastCol, _
            $iRangeLast, _
            $iRows, _
            $iCols, _
            $iWorksheets, _
            $iArrayNumCount, _
            $iSubfuncCount, _
            $aRawArray
    Local $aSecArray[1][1000], _
            $aElements[0], _
            $aUniqElements[0], _
            $aTempArrayStack[0][0], _
            $aSheetList[0], _
            $aIsSeparate[0], _
            $aExcessFunc[1]
    Local $sStrTypeValidator, _
            $sStrSubfuncValidator

    $oExcel = _Excel_Open(False)
    $oWorkbook = _Excel_BookOpen($oExcel, $xCellFile)
    $iWorksheets=$oWorkbook.Worksheets.Count
    $aSheetList=_Excel_SheetList($oWorkbook)
    $iRangeLast = $oWorkbook.Worksheets($iSheetnum).UsedRange.SpecialCells($xlCellTypeLastCell)
    $iLastCol=_Excel_ColumnToLetter($iRangeLast.Column)
    $aRawArray = _Excel_RangeRead( $oWorkbook, $iSheetnum, "E27:"&$iLastCol&"27")
    $iRows = UBound($aRawArray, $UBOUND_ROWS)
    $iCols = UBound($aRawArray, $UBOUND_COLUMNS)

;~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;~  Loop on the Extracted Array and Check if consist subfunctions~~
;~  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        For $i = 0 To $iRows - 1 ;Loop on Rows
            For $j = 0 To $iCols - 1;  Loop on column
                $sStrSubfuncValidator= StringInStr($aRawArray[$i][$j],")")
                If ($sStrSubfuncValidator<>0 ) Then ;Validate array element if argument or subfunction type
                    $sStrValidator=StringMid($aRawArray[$i][$j], $sStrSubfuncValidator+1,1)
                    If $sStrValidator == "@" Then;Check subfunction type
                        $aSecArray[$i][$j]=StringTrimLeft($aRawArray[$i][$j],$sStrSubfuncValidator +1)
                    ; Check if its a argument of subfunction
                    ElseIf $sStrValidator == "[" Then
                        $aSecArray[$i][$j]=$aRawArray[$i][$j] ;Append to Array subset
                    EndIf;====> Check subfunction type
                EndIf
            Next ;===Loop on column
        Next;===> Loop on Rows

    ;~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ;~                           Get Unique elements in the Array                    ~~
    ;~  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        For $i = 0 To $iRows - 1
            For $j = 0 To $iCols - 1
                If $j<>0 Then
                    If  $aSecArray[$i][$j] <> $aSecArray[$i][$j-1] And  $aSecArray[$i][$j] <> " "  Then; Check if element is the same as its preceeding element
                        $sStrValidator= StringInStr($aSecArray[$i][$j],"@") ;Check subfunc attributes
                        If $sStrValidator <>0 Then;Check if subfunction attributes
                            $aSecArray[$i][$j]=StringTrimLeft( $aSecArray[$i][$j],$sStrValidator)
                            If  $aSecArray[$i][$j] <> $aSecArray[$i][$j-1] And $aSecArray[$i][$j] <> " "  Then ;Check if same element as the previous one
                                If StringLen ($aSecArray[$i][$j]) <> 0 Then; Check if an empty string
                                    _ArrayAdd($aElements, $aSecArray[$i][$j]);Append to $aElements as unique
                                EndIf;===Check if an empty string
                            EndIf;===>  ;Check if same element as the previous one
                        Else
                            If StringLen ($aSecArray[$i][$j]) <> 0 Then; Check if an empty string
                                _ArrayAdd($aElements, $aSecArray[$i][$j]);Append to $aElements as unique
                            EndIf;===> Check if an empty string
                        EndIf
                    EndIf;===>Check if element is the same as its preceeding element
                EndIf;===>Check if first loop
                Next; ===>;Loop on Columns
        Next;===;Loop on Rows

        Local $tmparray[1] = [''] ;Instantiate an empty container
        For $i = 0 To UBound($aElements) - 1 ;Loop for the current element in the array
            _ArraySearch($tmparray, $aElements[$i]) ;Search for the unique item in the array
            If @error Then _ArrayAdd($tmparray, $aElements[$i]);Append array if unique
        Next
        For  $i = 0 To UBound($tmparray) -1
            If StringLen($tmparray[$i]) <> 0  Then;Remove empty values in the subset array
                _ArrayAdd($aUniqElements, $tmparray[$i]) ;Append Unique values on the main subset
            EndIf
        Next
        For $i =40 To UBound($aUniqElements) -1 ;Loop to Elements of $aUniqElements
            $iSubfuncCount= UBound($aUniqElements) -1
            If $iSubfuncCount  >=40 Then ; Check if it exceeds the byte limit range
                _ArrayAdd( $aExcessFunc, $aUniqElements[$i]); Append values to Excess functions Array
            EndIf

        Next;===>;Loop to Elements of $aUniqElements
        For $i = 0 To $iRows -1 ;Loop through rows
            For $j =0 To $iCols -1 ;Loop through columns
                For $a =0 To  UBound($aExcessFunc) -1 ;Loop within the elements of  $aExcessFunc
                    If StringInStr($aRawArray[$i][$j], $aExcessFunc[$a]) Then  ; Get all the position of indexes  of $aRawArray matched within the  aExcessFunc array
                        Local $test=_Excel_RangeWrite($oWorkbook, $oWorkbook.Worksheets($iSheetnum), "", _Excel_ColumnToLetter($j+5)&"28") ;Convert array indexes of $aRawArray to cell location
                    EndIf;===>Get all the position of indexes  of $aRawArray matched within the  aExcessFunc array
                Next
            Next;===> ;Loop through rows
        Next;===> ;Loop through columns
    _Excel_BookSave ( $oWorkbook )
EndFunc;===>; Delete Var Of Excess Functions

; #INTERNAL_USE_ONLY# ================================================================================================
; Name...........: jnkns_ChkIFSheet
; Description ...: Get string path of excel spreadsheet
; Author/s ........: rdbayanado
; Remarks .......:
; ================================================================================================================
Func _BE_jnkns_ChkIFSheet($xSheet)
    Local $bool_Isvalid
    If StringInStr($xSheet,"IF") Then
;~         MsgBox( 0, "Notifications","Before Append"&$xSheet)
        $xSheet=StringTrimRight($xSheet,5)
    Else
        $xSheet=$xSheet
    EndIf
    Return $xSheet
EndFunc

; #INTERNAL_USE_ONLY# ================================================================================================
; Name...........: _BE_GetXCellPath
; Description ...: Get string path of excel spreadsheet
; Author/s ........: cjhernandez
; Remarks .......:
; ================================================================================================================
Func _BE_GetXCellPath()
    Local $sTestSheetFile, _
            $sDrive, _
            $sDir, _
            $sFileName, _
            $sExtension, _
            $sFile
    Local $oExcel
    Local $hTextFile
    Local $fTestDesign

    $hTextFile = FileOpen($sLogTextFile, $FO_READ)
    $sTestSheetFile = FileReadLine($hTextFile, 2)
    $sTestSheetFile = StringTrimLeft($sTestSheetFile, 20)
    $oExcel = _Excel_Open(False)

    Local $aTestDesign[] = [$sTestSheetFile]
    For $i = 0 To UBound($aTestDesign, 1) - 1
        _PathSplit($aTestDesign[$i], $sDrive, $sDir, $sFileName, $sExtension)
    Next
    $sFile = $sDrive & $sDir & $sFileName & $sExtension
    Return $sFile
EndFunc

Func _BE_jnkns_SheetDelete($xCellFile,$iSheetnum)
    Local $oExcel, _
            $oBook
    $oExcel = _Excel_Open(False)
    $oBook = _Excel_BookOpen($oExcel,$xCellFile)

    _Excel_SheetDelete($oBook, $iSheetnum)
    _Excel_BookClose($oBook)
EndFunc

Func _BE_RunIFSheet($xCellFile)
    Local $oExcel
    Local $fIFSheet, _
            $fTestDesign
    Local $sExcelname
    Local $sDrive= "", _
            $sDir = "", _
            $sFileName = "", _
            $sExtension = ""
    Local  $aPathSplit

    $oExcel=_Excel_Open(True)

    $fIFSheet =_Excel_BookOpen ( $oExcel, @ScriptDir&"\Tools\UT Step1 - IFC Tool.xlsm" )
    Sleep(7000)
    $fTestDesign=_Excel_BookOpen ( $oExcel,$xCellFile )
    $aPathSplit= _PathSplit($xCellFile, $sDrive, $sDir, $sFileName, $sExtension)
    ;MsgBox(0,"",$sFileName&"-Excel")
    ;WinActivate($sFileName&"-Excel")
    Local $excelHwnd = WinGetHandle($sFileName&"-Excel")		; Sleep for 3 seconds
    Sleep(5000)
    ControlSend($excelHwnd,"","","^+s" )
    ;end("^+s")
    Sleep(15000)
    _Excel_BookSave ( $fTestDesign )
    _Excel_Close($oExcel, True, True)
    while 1
        If ProcessExists("EXCEL.EXE") Then ExitLoop
        Sleep(250)
    wend
    Sleep(1000)

EndFunc


Func _BE_jnkns_IsSepratedOrNot($xCellFile)
    Local $iWorksheets, _
            $iLoopCount, _
            $iRSSheet
    Local $aWorkSheetList

    $iRSSheet =0

    $oExcel = _Excel_Open(False) ;Instantiatiate Excel application
    $oWorkbook = _Excel_BookOpen($oExcel, $xCellFile, True, False)
    $iWorksheets =$oWorkbook.worksheets.Count ;Worksheets count
    $aWorkSheetList =_Excel_SheetList($oWorkbook) ;Store
        For $i =0  To $iWorksheets-1
            If StringInStr($aWorkSheetList[$i][0], "Reason_Result") Then
                $iRSSheet+=1
            EndIf
        Next
    Return $iRSSheet
    _Excel_BookClose($oWorkbook)
;~     If ($iWorksheets-$iRSSheet)== 4 Then
;~       For $i=3 To $iWorksheets-2
;~         jnkns_BE_DeleteVar($xCellFile,$i)
;~       Next
;~     EndIf
EndFunc
