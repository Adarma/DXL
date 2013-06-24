#cs ----------------------------------------------------------------------------

Author:
	Adam Cadamally

Description:
	Run DOORS Code

Licence:

	Copyright (c) 2013 Adam Cadamally

	Permission is hereby granted, free of charge, to any person obtaining a copy
	of this software and associated documentation files (the "Software"), to deal
	in the Software without restriction, including without limitation the rights
	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
	copies of the Software, and to permit persons to whom the Software is
	furnished to do so, subject to the following conditions:

	The above copyright notice and this permission notice shall be included in
	all copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
	THE SOFTWARE.

#ce ----------------------------------------------------------------------------

; MAIN PROGRAM

Opt("WinSearchChildren", 1)		;0=no, 1=search children also
Opt("MustDeclareVars", 1)		;0=no, 1=yes
;~ Opt("RunErrorsFatal", 0)		;1=fatal, 0=silent set @error is in the script


Local $oErrorHandler = ObjEvent("AutoIt.Error","ComErrorHandler")    ; Initialize a COM error handler
; This is my custom defined error handler
Func ComErrorHandler($oMyError)
	Return
	Msgbox(0,"AutoItCOM Error","We intercepted a COM Error !"    & @CRLF  & @CRLF & _
             "err.description is: " & @TAB & $oMyError.description  & @CRLF & _
             "err.windescription:"   & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "       & @TAB & hex($oMyError.number,8)  & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @CRLF & _
             "err.source is: "       & @TAB & $oMyError.source       & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile     & @CRLF & _
             "err.helpcontext is: " & @TAB & $oMyError.helpcontext _
            )
Endfunc



; Check if any Arguments were passed
If $CmdLine[0] < 4 Then
	MsgBox(0, "Wrong Parameters", "Wrong Parameters: " & $CmdLine[0])
Else
	; Get Arguments
	Local $IncludeString = $CmdLine[1]
	Local $ModuleFullName = $CmdLine[2]
	Local $OutFile = $CmdLine[3]
	Local $DxlMode = $CmdLine[4]
	
;~ 	$PrintRepurposeCode = _
;~ 	"#include <C:/Documents/DXL/Includes/Debug/Redirection.inc>" & @CRLF & _
;~ 	"Redirection_Start()"& @CRLF
	
	
	Local $EscapedOutFile = StringReplace($OutFile, "\", "\\")
	Local $PrintRepurposeCode = _
		"Stream oSublimeText2Stream = write(""" & $EscapedOutFile & """, CP_UTF8)" & @CRLF & _
		"void print(string s)" & @CRLF & _
		"{" & @CRLF & _
		@TAB & "if(!null(s)) {" & @CRLF & _
		@TAB & @TAB & "oSublimeText2Stream << s" & @CRLF & _
		@TAB & @TAB & "flush(oSublimeText2Stream)" & @CRLF & _
		@TAB & "}" & @CRLF & _
		"}" & @CRLF & _
		'void print(bool b)		{ print(b "\n") }' & @CRLF & _
		'void print(char c)		{ print(c "\n") }' & @CRLF & _
		'void print(Date d)		{ print(d "\n") }' & @CRLF & _
		'void print(int i)		{ print(i "\n") }' & @CRLF & _
		'void print(real r)		{ print(r "\n") }' & @CRLF
	
	Local $SetModuleCode = ""
	If $ModuleFullName <> "" Then
		; Set the current module
		$SetModuleCode = "// Set Module" & @CRLF & _
			'{' & @CRLF & _
			@TAB & 'Item oItem = item("' & $ModuleFullName & '"); ' & @CRLF & _
			@TAB & 'if(!null(oItem)) {' & @CRLF & _
			@TAB & @TAB & 'Module oModule = module(oItem);' & @CRLF & _
			@TAB & @TAB & 'if(!null(oModule)) {' & @CRLF & _
			@TAB & @TAB & @TAB & '(current ModuleRef__) = oModule;' & @CRLF & _
			@TAB & @TAB & '}' & @CRLF & _
			@TAB & '}' & @CRLF & _
			'}' & @CRLF
	EndIf


	Local $Code = $IncludeString
	Switch $DxlMode
		Case 1
			; Check Final Allocations
			; TODO: Overload functions: halt, show, block etc
			$Code = '#include <C:\Documents\DXL\Includes\Debug\Debug.inc>;' & @CRLF
			$Code = $Code & $IncludeString & @CRLF
			$Code = $Code & 'int iCount = Debug_GetAllocatedObjectCount();' & @CRLF
			$Code = $Code & 'print("Final Allocated Object Count : " iCount "\n");' & @CRLF
		Case 2
			; Log Allocations
			; TODO: Pipe output
			$Code = "#include <C:\Documents\DXL\Includes\Debug\Debug.inc>;" & @CRLF
			$Code = $Code & "Debug_Logging(true);" & @CRLF
			$Code = $Code & $IncludeString & @CRLF
			$Code = $Code & "Debug_Logging(false);" & @CRLF
		Case 3
			; Trace DXL
			; TODO: Pipe output
			$Code = 'startDXLTracing_("C:\\DxlVariables.txt");' & @CRLF
			$Code = $Code & $IncludeString & @CRLF
			$Code = $Code & 'stopDXLTracing_();' & @CRLF
	EndSwitch

;~ 	$PostfixCode = @CRLF & _
;~ 	"Redirection_Display()" & @CRLF & _
;~ 	"Redirection_Stop()"
	   
	Local $PostfixCode = @CRLF & _
	   "close(oSublimeText2Stream)" & @CRLF
	
;~ 	Local $FullCode = $PrintRepurposeCode & $SetModuleCode & $IncludeString & $PostfixCode
	Local $FullCode = $PrintRepurposeCode & $SetModuleCode & $Code
	
	Local $DxlInteractionWindow = "DXL Interaction - DOORS"
	Local $DxlOpen = WinExists($DxlInteractionWindow) And BitAnd(WinGetState($DxlInteractionWindow), 2)

	If $DxlOpen Then
		
		; Write the code in DOORS window
		ControlSetText($DxlInteractionWindow, "", "[CLASS:RICHEDIT50W; INSTANCE:2]", $FullCode)

		; Click the Run Button
		;WinActivate($DoorsWindow)
		ControlClick($DxlInteractionWindow, "", "[CLASS:Button; INSTANCE:8]")
		
	Else
		; Run the Code via COM
		Local $ObjDoors = ObjCreate("DOORS.application")
		$ObjDoors.Result = ""
		$ObjDoors.runStr($FullCode)
		$ObjDoors = 0
	EndIf
	
EndIf
