/*
Author:
	Adam Cadamally
	Mathias Mamsch

Contact:
	adamcadamally@gmail.com

Namespaces:
	Debug_... 			// For public external use where this library is included
	Debug_Private_...	// For private internal use in this library

Description:
	This file is included before the dxl being executed when using a debug build.
	It calls a callback function for every line of the program.
	The call back function logs the file & line number and any Object [de]allocations.

Reason for this library:
	To trace code execution
	To find slow code
	To find which line causes DOORS to crash
	To debug Memory Leaks caused by not releasing Allocated Objects

Usage:

	Get the Allocated Object Count at a certain point:
		int Count = Debug_GetAllocatedObjectCount()
		print "Allocated Object Count : " Count "\n"

	Enable Allocated Object Logging - Surround suspect code with:
		Debug_Logging(true, true, true, "", "")
		{CODE}
		Debug_Logging(false, false, false, "", "")

Example:

	Buffer CreateBuffer()
	{
		Buffer NewBuffer = create()
		return NewBuffer
	}

	Debug_AppendToLog("Begin Allocated Object Count Log:\n")

	Debug_Logging(true, true, true, "", "")

	Debug_AppendToLog("Create some objects:\n")
	Array TestArray = create(1, 1)
	Skip TestSkip = create()
	Buffer TestBuffer = CreateBuffer()

	Debug_AppendToLog("Now Delete them:\n")
	delete TestArray
	delete TestSkip
	delete TestBuffer

	Debug_Logging(false, false, false, "", "")

	int Count = Debug_GetAllocatedObjectCount()
	print "Final Allocated Object Count : " Count "\n"

	Debug_Cleanup()

Note:
	Remember to turn Logging Off at the end of the script.
	Remember to turn Logging Off before halt statements.
	Could cause Memory Exception Errors if:
		Not turned off before script ends.
		Runtime Error before it gets turned off.

	Uses: oleGetResult & oleSetResult
		So wont work on scripts which use use them.

	Logging is done to Files rather than prints/Buffers so the logs are available even if DOORS crashes

Licence:

	Copyright (c) 2012 Adam Cadamally

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
*/

pragma encoding, "UTF-8"
XFLAGS_ &= ~AutoDeclare_

const int Debug_Version = 1

//==================================== GLOBAL VARIABLES ===================================\\


// DxlObject to pass Variables between Dxl Contexts
DxlObject Debug_Private_VariablesObject = new()


//=================================== PRIVATE FUNCTIONS ===================================\\


// The Path of this File
string Debug_Private_GetThisDxlFilePath()
{
	string DXLFileAndLine = dxlHere()
	string DXLFile = ""
	Regexp ExtractPath = regexp2("<(.*):[0-9]+>")
	if(ExtractPath DXLFileAndLine) {
		DXLFile = DXLFileAndLine[match 1]
	}
	delete(ExtractPath)
	return(DXLFile)
}

// Increment the Address of a Pointer by Offset
int *Debug_Private_IncrementPtr(int *Ptr, int Offset)
{
	int *IncrementedPtr = Ptr
	IncrementedPtr += Offset
	return IncrementedPtr
}

// Advance a general Linked List Pointer
// The Address of the next List Item is Offest bits on from the Current Item
int *Debug_Private_AdvanceLinkedListPtr(int *Ptr, int Offset)
{
	int *IncrementedPtr = Debug_Private_IncrementPtr(Ptr, Offset)
	int ValueAtIncrementedPtr = *IncrementedPtr
	int *PointerFromValue = addr_(ValueAtIncrementedPtr)
	return PointerFromValue
}

// Advance the Allocated Objects Linked List Pointer
// The Address of the next List Item is 8 bits on
int *Debug_Private_NextAllocatedObjectPtr(int *AllocatedObjectPtr)
{
	return Debug_Private_AdvanceLinkedListPtr(AllocatedObjectPtr, 8)
}

// Get Pointer to the AllocatedObjects LinkedList
// It is 0x74 bits on from the DXL Context
int *Debug_Private_GetFirstAllocatedObjectPtr(int *CurrentDXLContextPtr)
{
	return Debug_Private_AdvanceLinkedListPtr(CurrentDXLContextPtr, 0x74)
}

// Get the Current DXL Context from a DB
// Its address is stored on the heap 48 bits on from the DB
int Debug_Private_GetCurrentDXLContextAddress()
{
	DB TempDB = create("")
	int *TempDBPtr = addr_(TempDB)
	int *CurrentDXLContextPtr = Debug_Private_AdvanceLinkedListPtr(TempDBPtr, 48)
	destroy(TempDB)
	int CurrentDXLContextAddress = (addr_(CurrentDXLContextPtr)) int
	return CurrentDXLContextAddress
}

// Walk the AllocatedObjects LinkedList and count the Items in it
int Debug_Private_GetAllocatedObjectCountForContext(int ContextAddress)
{
	if(null(ContextAddress)) {
		return(-1)
	}

	int *CurrentDXLContextPtr = addr_(ContextAddress)
	int *AllocatedObjectPtr = Debug_Private_GetFirstAllocatedObjectPtr(CurrentDXLContextPtr)
	int AllocatedObjectCount = 0
	while(!null(AllocatedObjectPtr))
	{
		AllocatedObjectPtr = Debug_Private_NextAllocatedObjectPtr(AllocatedObjectPtr)
		AllocatedObjectCount++
	}
	return AllocatedObjectCount
}

// Define A Debug CallBack - Executes before every line
// Executed in it own Context so can't access global variables
void Debug_Private_Callback(string FilePath, int LineNumber)
{
	if(FilePath == "Line")	{ return }
	if(LineNumber == 0)		{ return }

	// Get the address of the DxlObject by oleGetResult
	int iDxlObjectAddress = intOf(realOf(oleGetResult()))
	if(null(iDxlObjectAddress)) {
		setDebugging_(false)
		return
	}

	// Get the DxlObject
	DxlObject oDxlObject = (addr_(iDxlObjectAddress)) DxlObject
	if(null(oDxlObject)) { return }

	// Get ThisDxlFilePath
	string ThisDxlFilePath = (oDxlObject->"ThisDxlFilePath") string
	if(FilePath == ThisDxlFilePath) { return }

	// Append to Trace Log Stream
	bool LogLines = (oDxlObject->"LogLines") bool
	if(LogLines) {

		Stream TimeElapsedStream = (oDxlObject->"TimeElapsedStream") Stream
		if(!null(TimeElapsedStream)) {

			string FileFilterExecution = (oDxlObject->"FileFilterExecution") string
			if(null(FileFilterExecution) or null(FilePath) or FilePath == FileFilterExecution)
			{
				string PreviousExecutionFile = (oDxlObject->"PreviousExecutionFile") string
				int PreviousExecutionLine = (oDxlObject->"PreviousExecutionLine") int
				if(FilePath != PreviousExecutionFile or LineNumber != PreviousExecutionLine) {

					int TimeNow = getTickCount_()
					if(PreviousExecutionLine != -1) {

						int PreviousExecutionTime = (oDxlObject->"PreviousExecutionTime") int
						int ExecutionTime = TimeNow - PreviousExecutionTime

						bool LogInstantLines = (oDxlObject->"LogInstantLines") bool

						if(LogInstantLines or ExecutionTime > 0) {
							TimeElapsedStream << ("<" PreviousExecutionFile ":" PreviousExecutionLine "> " ExecutionTime "\n")
							flush(TimeElapsedStream)
						}
					}

					// Store Current File, Line and Time
					oDxlObject->"PreviousExecutionFile" = FilePath
					oDxlObject->"PreviousExecutionLine" = LineNumber
					oDxlObject->"PreviousExecutionTime" = TimeNow
				}
			}
		}
	}

	// Append to Allocations Log Stream
	bool LogAllocations = (oDxlObject->"LogAllocations") bool
	if(LogAllocations) {

		Stream AllocationsStream = (oDxlObject->"AllocationsStream") Stream
		if(!null(AllocationsStream)) {

			string FileFilterAllocation = (oDxlObject->"FileFilterAllocation") string
			if(null(FileFilterAllocation) or null(FilePath) or FilePath == FileFilterAllocation) {

				int ContextAddress = (oDxlObject->"CurrentDXLContext") int
				int TotalAllocatedObjects = Debug_Private_GetAllocatedObjectCountForContext(ContextAddress)
				int InitialAllocations = (oDxlObject->"InitialAllocations") int
				int CurrentAllocatedObjectCount = TotalAllocatedObjects - InitialAllocations

				int PreviousAllocationCount = (oDxlObject->"PreviousAllocationCount") int
				if((PreviousAllocationCount) != -1) {

					int AllocatedObjectChange = CurrentAllocatedObjectCount - PreviousAllocationCount
					if(AllocatedObjectChange != 0) {

						string PreviousAllocationFile = (oDxlObject->"PreviousAllocationFile") string
						int PreviousAllocationLine = (oDxlObject->"PreviousAllocationLine") int

						AllocationsStream << ("<" PreviousAllocationFile ":" PreviousAllocationLine "> (" )
						if(AllocatedObjectChange > 0) {
							AllocationsStream << ("+")
						}
						AllocationsStream << (AllocatedObjectChange ") " CurrentAllocatedObjectCount "\n")
						flush(AllocationsStream)
					}
				}

				// Store Current File, Line and Count
				oDxlObject->"PreviousAllocationFile" = FilePath
				oDxlObject->"PreviousAllocationLine" = LineNumber
				oDxlObject->"PreviousAllocationCount" = CurrentAllocatedObjectCount
			}
		}
	}
}

void Debug_Private_Initialize(DxlObject oDxlObject)
{
	// Make sure Debugging is off
	setDebugging_(false)

	// Default Log Files
	string CallTraceFilePath = "c:\\DxlCallTrace.log"
	string AllocationsFilePath = "c:\\DxlAllocations.log"

	// Create and Clear the Log Files
	Stream TimeElapsedStream = write(CallTraceFilePath)
	Stream AllocationsStream = write(AllocationsFilePath)

	int CurrentDXLContext = Debug_Private_GetCurrentDXLContextAddress()

	// Store the variables in the DxlObject
	oDxlObject->"ThisDxlFilePath" = Debug_Private_GetThisDxlFilePath()
	oDxlObject->"CurrentDXLContext" = CurrentDXLContext
	oDxlObject->"InitialAllocations" = Debug_Private_GetAllocatedObjectCountForContext(CurrentDXLContext)
	oDxlObject->"TimeElapsedStream" = TimeElapsedStream
	oDxlObject->"AllocationsStream" = AllocationsStream
	oDxlObject->"LogLines" = false
	oDxlObject->"LogAllocations" = false
	oDxlObject->"LogInstantLines" = false
	oDxlObject->"PreviousExecutionFile" = ""
	oDxlObject->"PreviousExecutionLine" = -1
	oDxlObject->"PreviousExecutionTime" = getTickCount_()
	oDxlObject->"PreviousAllocationFile" = ""
	oDxlObject->"PreviousAllocationLine" = -1
	oDxlObject->"PreviousAllocationCount" = -1
	oDxlObject->"FileFilterAllocation" = ""
	oDxlObject->"FileFilterExecution" = ""

	// Set the Callback - Executes in a different Context
	setDebugCallback_(Debug_Private_Callback)
}

//=================================== PUBLIC FUNCTIONS ====================================\\

// Get Number of Allocated Objects
int Debug_GetAllocatedObjectCount()
{
	int CurrentDXLContextAddress = (Debug_Private_VariablesObject->"CurrentDXLContext") int
	int InitialAllocations = (Debug_Private_VariablesObject->"InitialAllocations") int
	return Debug_Private_GetAllocatedObjectCountForContext(CurrentDXLContextAddress) - InitialAllocations
}

// Append some text to the logs
void Debug_AppendToLog(string AppendText)
{
	Stream TimeElapsedStream = (Debug_Private_VariablesObject->"TimeElapsedStream") Stream
	Stream AllocationsStream = (Debug_Private_VariablesObject->"AllocationsStream") Stream

	if(!null(TimeElapsedStream)) {
		TimeElapsedStream << AppendText
	}

	if(!null(AllocationsStream)) {
		AllocationsStream << AppendText
	}
}

// Turn Logging on/off
void Debug_Logging(bool LogLines, LogAllocations, LogInstantLines, string FileFilterAllocation, FileFilterExecution)
{
	// Make sure Logging is Off
	setDebugging_(false)

	bool PreLogLines = (Debug_Private_VariablesObject->"LogLines") bool
	bool PreLogAllocations = (Debug_Private_VariablesObject->"LogAllocations") bool

	if(LogLines or LogAllocations) {

		// Reset the PreviousExecutionTime Object Count
		if(!PreLogLines) {
			Debug_Private_VariablesObject->"PreviousExecutionTime" = -1
		}

		// Reset the PreviousAllocated Object Count
		if(!PreLogAllocations) {
			Debug_Private_VariablesObject->"PreviousAllocationCount" = -1
		}

		// The the Debug_Private_Callback executes in a different context so can't access variables in this Context
		// Use oleResult to store the Address of the Debug_Private_VariablesObject so it can be reteved in Debug_Private_Callback
		oleSetResult(((addr_(Debug_Private_VariablesObject)) int) "")

		// Debug_AppendToLog("Logging Enabled\n")
	} else {

		// Call Callback to force last allocation to be logged
		Debug_Private_Callback("", -1)

		// Clear the Debug_Private_VariablesObject Address
		oleSetResult("0")

		// Debug_AppendToLog("Logging Disabled\n")
	}

	Debug_Private_VariablesObject->"LogLines" = LogLines
	Debug_Private_VariablesObject->"LogAllocations" = LogAllocations
	Debug_Private_VariablesObject->"LogInstantLines" = LogInstantLines

	Debug_Private_VariablesObject->"FileFilterAllocation" = FileFilterAllocation
	Debug_Private_VariablesObject->"FileFilterExecution" = FileFilterExecution

	// Turn Logging On or Off
	setDebugging_(LogLines or LogAllocations)
}

/*	Debug_Cleanup: Destroy the Global DxlObject and Streams
	Any use of the library after calling this will crash.
	It may be called when the library will not be used again,
	but is unnecessary since the variables will be destroyed
	when the dxl finishes anyway.
	Recommend not to Call */
void Debug_Cleanup()
{
	// Make sure it is off
	Debug_Logging(false, false, false, "", "")

	if(!null(Debug_Private_VariablesObject)) {
		// Clear the Debug_Private_VariablesObject Address
		oleSetResult("0")

		Stream TimeElapsedStream = (Debug_Private_VariablesObject->"TimeElapsedStream") Stream
		Stream AllocationsStream = (Debug_Private_VariablesObject->"AllocationsStream") Stream

		// Cleanup Streams
		close(TimeElapsedStream)
		TimeElapsedStream = null
		close(AllocationsStream)
		AllocationsStream = null

		// Cleanup DxlObject
		delete(Debug_Private_VariablesObject)
		Debug_Private_VariablesObject = null
	}
}


//================================= REPURPOSED FUNCTIONS =================================\\


// Originals
void Debug_Private_OriginalHalt() {
	halt
}
void Debug_Private_OriginalShow(DB oDB) {
	show(oDB)
}
void Debug_Private_OriginalDestroy(DB oDB) {
	destroy(oDB)
}

// Repurposed
void halt() {
	Debug_Cleanup()
	Debug_Private_OriginalHalt()
}
void show(DB oDB) {
	int InitialShowDB = (Debug_Private_VariablesObject->"InitialShowDB") int
	if(null(InitialShowDB)) {
		int Address = ((addr_(oDB)) int)
		Debug_Private_VariablesObject->"InitialShowDB" = Address
	}
	Debug_Private_OriginalShow(oDB)
}
void destroy(DB oDB) {
	int Address = ((addr_(oDB)) int)
	int StoredAddress = (Debug_Private_VariablesObject->"InitialShowDB") int
	if(StoredAddress == Address) {
		Debug_Cleanup()
	}
	Debug_Private_OriginalDestroy(oDB)
}


//===================================== INITIALIZATION ====================================\\


Debug_Private_Initialize(Debug_Private_VariablesObject)
