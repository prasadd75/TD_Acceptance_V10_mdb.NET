Option Strict Off
Option Explicit On
Friend Class clsIniFile
	' ___________________________________________________________________________
	' ***************************************************************************
	
	
	'   ** Windows API calls **
	'
	'   ** For Private INIs
	Private Declare Function GetPrivateProfileInt Lib "kernel32"  Alias "GetPrivateProfileIntA"(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Private Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Private Declare Function GetPrivateProfileSection Lib "kernel32"  Alias "GetPrivateProfileSectionA"(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Private Declare Function WritePrivateProfileString Lib "kernel32"  Alias "WritePrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
	Private Declare Function WritePrivateProfileSection Lib "kernel32"  Alias "WritePrivateProfileSectionA"(ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	'   ** For Windows INIs
	Private Declare Function GetProfileInt Lib "kernel32"  Alias "GetProfileIntA"(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer) As Integer
	Private Declare Function GetProfileString Lib "kernel32"  Alias "GetProfileStringA"(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
	Private Declare Function GetProfileSection Lib "kernel32"  Alias "GetProfileSectionA"(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
	Private Declare Function WriteProfileString Lib "kernel32"  Alias "WriteProfileStringA"(ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Integer
	Private Declare Function WriteProfileSection Lib "kernel32"  Alias "WriteProfileSectionA"(ByVal lpAppName As String, ByVal lpString As String) As Integer
	'   ** Fow Windows Communications
	Private Declare Function GetWindowsDirectory Lib "kernel32"  Alias "GetWindowsDirectoryA"(ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	Private Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Integer, ByVal wmsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer
	
	
	'   ** Constants used to size buffers
	
	'   Set the max size of the string to be returned in bytes.
	'   *** Use factors of 1024 to a MAX of 32767. ***
	Private Const Max_SectionBuffer As Short = 4096 'bytes
	'   Set the max size of an entry in bytes.
	Private Const Max_EntryBuffer As Short = 255 'bytes
	
	'   ** Special values to alert other apps of Win.Ini changes
	Private Const HWND_BROADCAST As Short = &HFFFFs
	Private Const WM_WININICHANGE As Short = &H1As
	
	'   ** Enum variables for ExtractPath routine.
	Public Enum ExtractFileName_Constants
		exnFullName = 1
		exnOnlyExtn = 2
	End Enum
	
	'   Module level variables to hold property values
	Private mstrSectionName As String 'local copy
	Private mstrFileName As String 'local copy
	Private mstrFilePath As String 'local copy
	Private mstrFullPath As String 'local copy
	Private mstrPrivPath As String 'local copy
	Private mstrFileExt As String 'local copy
	Private mbooInitialized As Boolean 'local copy
	Private mbooUseWinINI As Boolean 'local copy
	
	Private ReadOnly Property PrivPath() As String
		Get
			'The private path is a read only storage area.
			'The private path is set during initialization
			'   but can be changed by setting the FullPath property.
			'   It exists to allow switching to/from win.ini.
			PrivPath = mstrPrivPath
			
		End Get
	End Property
	
	
	Private Property UseWinINI() As Boolean
		Get
			
			UseWinINI = mbooUseWinINI
			
		End Get
		Set(ByVal Value As Boolean)
			Dim mvarFilName As Object
			
			Dim strBuff As String
			Dim intRet As Short
			Dim intDot As Short
			
			If mbooUseWinINI = Value Then Exit Property Else mbooUseWinINI = Value
			
			If Value Then 'Win INI is to be used.
				'Find Win.Ini
				strBuff = New String(Chr(0), Max_EntryBuffer)
				intRet = GetWindowsDirectory(strBuff, Max_EntryBuffer)
				mstrFullPath = Left(strBuff, intRet) & "\WIN.INI"
				mstrFilePath = ExtractPath(mstrFullPath)
				mstrFileName = ExtractName(mstrFullPath, ExtractFileName_Constants.exnFullName)
				intDot = InStr(mstrFileName, ".")
				'UPGRADE_WARNING: オブジェクト mvarFilName の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
				If intDot Then mstrFileExt = Right(mvarFilName, Len(mstrFileName) - intDot)
			Else
				FullPath = mstrPrivPath
			End If
			
		End Set
	End Property
	
	
	Public Property FullPath() As String
		Get
			
			'The full path property can be read to find out what
			'  file is currently in use.
			FullPath = mstrFullPath
			
		End Get
		Set(ByVal Value As String)
			
			'Setting the FullPath property is used for using a file other than the
			'  default or win.ini files. Assign the Private Path String. Adjust other
			'  properties to suit. It's redundant to set this property to the Win.INI
			'  file  - simply set the UseWinINI to true.
			
			'Make sure you're using a valid path BEFORE setting this property!
			'No internal path checking is done!
			
			mstrPrivPath = Value
			mstrFullPath = mstrPrivPath
			mstrFilePath = ExtractPath(mstrFullPath)
			mstrFileName = ExtractName(mstrFullPath, ExtractFileName_Constants.exnFullName)
			mstrFileExt = ExtractName(mstrFullPath, ExtractFileName_Constants.exnOnlyExtn)
			If Trim(mstrSectionName) = "" Then mbooInitialized = False Else mbooInitialized = True
			mbooUseWinINI = False
			
		End Set
	End Property
	
	
	Public Property SectionName() As String
		Get
			
			'Retrieves the current section name.
			SectionName = mstrSectionName
			
		End Get
		Set(ByVal Value As String)
			
			'The section name is set prior to data being read or
			'  written to the selected INI file.
			mstrSectionName = Value
			
			If Trim(Value) = "" Then
				mbooInitialized = False
			Else
				mbooInitialized = True
			End If
			
		End Set
	End Property
	
	Private ReadOnly Property FilePath() As String
		Get
			
			'The file path property is read only and contains only the path
			'  of the current INI file.
			FilePath = mstrFilePath
			
		End Get
	End Property
	
	Private ReadOnly Property FileExt() As String
		Get
			
			'The file extension property is read only and contains only the extension
			'  of the current INI file.
			FileExt = mstrFileExt
			
		End Get
	End Property
	
	Private ReadOnly Property FileName() As String
		Get
			
			FileName = mstrFileName
			
		End Get
	End Property
	
	Private Function ExtractName(ByVal PathIn As String, ByVal RetChoice As ExtractFileName_Constants) As String
		
		Dim intCount As Short
		Dim intDot As Short
		Dim strSpecOut As String
		
		On Error Resume Next
		
		'The Extract Name function is used internally by the class but I
		'  have left it as Public so that it is available for othe purposes.
		'If you use it for other purposes - use it properly.
		
		If Len(PathIn) = 0 Then PathIn = mstrFullPath
		
		'Extract and return the full file name.
		If InStr(PathIn, "\") Then
			For intCount = Len(PathIn) To 1 Step -1
				If Mid(PathIn, intCount, 1) = "\" Then
					strSpecOut = Mid(PathIn, intCount + 1)
					Exit For
				End If
			Next intCount
		ElseIf InStr(PathIn, ":") = 2 Then 
			strSpecOut = Mid(PathIn, 3)
		Else
			strSpecOut = PathIn
		End If
		
		intDot = InStr(strSpecOut, ".")
		
		'Returns only the base of the file name.
		If intDot And RetChoice = ExtractFileName_Constants.exnFullName Then strSpecOut = Left(strSpecOut, intDot - 1)
		
		'Returns only the extension of the filename.
		If intDot And RetChoice = ExtractFileName_Constants.exnOnlyExtn Then strSpecOut = Right(strSpecOut, Len(strSpecOut) - intDot)
		
		ExtractName = strSpecOut
		
	End Function
	
	Private Function ExtractPath(ByVal PathIn As String) As String
		
		Dim intCount As Short
		Dim strSpecOut As String
		
		On Error Resume Next
		
		'The Extract Path function is used internally by the class but I
		'  have left it as Public so that it is available for othe purposes.
		'If you use it for other purposes - use it properly.
		
		If Len(PathIn) = 0 Then PathIn = mstrFullPath
		
		If InStr(PathIn, "\") Then
			For intCount = Len(PathIn) To 1 Step -1
				If Mid(PathIn, intCount, 1) = "\" Then
					strSpecOut = Left(PathIn, intCount - 1) 'Reduced length of strSpecOut by one - 99/07/17.
					Exit For
				End If
			Next intCount
		ElseIf InStr(PathIn, ":") = 2 Then 
			strSpecOut = CurDir(PathIn)
			If Len(strSpecOut) = 0 Then strSpecOut = CurDir()
		Else
			strSpecOut = CurDir()
		End If
		
		If Right(strSpecOut, 1) = "\" Then
			strSpecOut = Left(strSpecOut, Len(strSpecOut) - 1)
		End If
		
		ExtractPath = strSpecOut
		
	End Function
	
	Public Sub DeleteStrEntry(ByVal EntryName As String)
		
		'Bail if not initialized
		If Not mbooInitialized Then
			WarnAuthor()
			Exit Sub
		End If
		
		'Sets a specific entry to Nothing or Blank.
		Dim lngRetVal As Integer
		If mbooUseWinINI Then
			lngRetVal = WriteProfileString(mstrSectionName, EntryName, "")
			WinIniChanged()
		Else
			lngRetVal = WritePrivateProfileString(mstrSectionName, EntryName, "", mstrFullPath)
		End If
		
	End Sub
	
	Private Sub DeleteNumEntry(ByVal EntryName As String)
		
		'Bail if not initialized.
		If Not mbooInitialized Then
			WarnAuthor()
			Exit Sub
		End If
		
		'Deletes a specific entry.
		Dim lngRetVal As Integer
		If mbooUseWinINI Then
			lngRetVal = WriteProfileString(mstrSectionName, EntryName, CStr(0))
			WinIniChanged()
		Else
			lngRetVal = WritePrivateProfileString(mstrSectionName, EntryName, 0, mstrFullPath)
		End If
		
	End Sub
	
	Public Sub DeleteSection()
		
		'Bail if not initialized
		If Not mbooInitialized Then
			WarnAuthor()
			Exit Sub
		End If
		
		'Deletes all of the current [Section]s Entries.
		Dim lngRetVal As Integer
		If mbooUseWinINI Then
			lngRetVal = WriteProfileSection(mstrSectionName, "")
			WinIniChanged()
		Else
			lngRetVal = WritePrivateProfileSection(mstrSectionName, "", mstrFullPath) 'WritePrivateProfileString(mstrSectionName, 0&, 0&, mstrFullPath)
		End If
		
		mstrSectionName = ""
		mbooInitialized = False
		
		
	End Sub
	
	Function Exists(ByVal filname As String) As Short
		Dim myERROR As Object
		
		
		'Set up default error handling
		On Error GoTo Err_Exists
		
		'******** Coding starts here ********
		' returns true if the file "filname" exist
		
		If InStr(filname, ".") < 1 Then
			filname = filname & ".*"
		End If
		
		'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
		Exists = Dir(filname) > ""
		
		Exit Function
		
		'********* Coding ends here *********
		Exit Function
		
		'Default error handler
Err_Exists: 
		
		Select Case Err.Number
			Case Else
				'Call the default error handler
				'UPGRADE_WARNING: オブジェクト myERROR.Handler の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
				myERROR.Handler("WINPARTS.BAS", "Exists", Err.Source, Err.Description, Err.Number)
				'NB Default error handling ends procedure
				Resume End_Exists
		End Select
		
End_Exists: 
		
	End Function
	
	Private Function GetInt(ByVal EntryName As String, ByVal DefaultInt As Short) As Short
		
		'Bail if not initialized
		If Not mbooInitialized Then
			WarnAuthor()
			Exit Function
		End If
		
		'Retrieves an Integer value, range: 0-32767
		If mbooUseWinINI Then
			GetInt = GetProfileInt(mstrSectionName, EntryName, DefaultInt)
		Else
			GetInt = GetPrivateProfileInt(mstrSectionName, EntryName, DefaultInt, mstrFullPath)
		End If
		
	End Function
	
	Private Function GetSectEntries() As String
		
		'Bail if not initialized.
		If Not mbooInitialized Then
			WarnAuthor()
			Exit Function
		End If
		
		'Retrieves all Entries in a [Section]
		'Returnes a string of Null delineated entries "EntryName=Value&vbNull&...." with the
		'  last entry double-terminated.
		
		Dim strTemp As New VB6.FixedLengthString(Max_SectionBuffer)
		Dim lngRetVal As Integer
		If mbooUseWinINI Then
			lngRetVal = GetProfileSection(mstrSectionName, strTemp.Value, Len(strTemp.Value))
		Else
			lngRetVal = GetPrivateProfileSection(mstrSectionName, strTemp.Value, Len(strTemp.Value), mstrFullPath)
		End If
		
		GetSectEntries = Left(strTemp.Value, lngRetVal + 1)
		
	End Function
	
	Private Function GetSectEntriesEx(ByRef DataArry() As String) As Short
		Dim GetSectionsEntriesEx As Object
		
		'Bail if not initialized.
		If Not mbooInitialized Then
			WarnAuthor()
			Exit Function
		End If
		
		'GetSectionEntriesEx retrieves all of the entries in the current [Section] and
		'  returns tha values in the two dimensional DataArry(1, Entries) array.
		'  The 0 Row holding the Entry Name and the 1 holding the Values.
		'  There will be as many Columns as there are Entries.
		
		On Error Resume Next
		
		'Get "normal" null terminated string of all [Section] Entries.
		Dim strTemp As String
		strTemp = GetSectEntries()
		If Len(strTemp) = 1 Then
			'UPGRADE_WARNING: オブジェクト GetSectionsEntriesEx の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			GetSectionsEntriesEx = 0
			Exit Function
		End If
		
		'Parse null terminated string of [Section] Entries into table.
		Dim intEntries As Short
		Dim intNull As Short
		Dim intLoc As Short
		Do While Asc(strTemp)
			ReDim Preserve DataArry(1, intEntries)
			intNull = InStr(strTemp, Chr(0))
			intLoc = InStr(Left(strTemp, intNull - 1), "=")
			DataArry(0, intEntries) = Left(strTemp, intLoc - 1)
			DataArry(1, intEntries) = Mid(strTemp, intLoc + 1, intNull - intLoc - 1)
			strTemp = Mid(strTemp, intNull + 1)
			intEntries = intEntries + 1
			If strTemp = "" Then Exit Do
		Loop 
		
		'Make function assignment
		GetSectEntriesEx = intEntries
		
	End Function
	
	Private Function GetSections() As String
		
		'Bail if not initialized
		If Not mbooInitialized Then
			WarnAuthor()
			Exit Function
		End If
		
		'The GetSections function is used to return a Null delineated string of all
		'  the [Section]s in current file.
		
		'Setup some variables
		Dim strRet As String
		Dim strBuff As String
		Dim intFileHandle As Short
		
		'Extract all [Section] lines
		intFileHandle = FreeFile
		FileOpen(intFileHandle, mstrFullPath, OpenMode.Input)
		Do While Not EOF(intFileHandle)
			strBuff = LineInput(intFileHandle)
			strBuff = StripComment(strBuff)
			If InStr(strBuff, "[") = 1 And InStr(strBuff, "]") = Len(strBuff) Then
				strRet = strRet & Mid(strBuff, 2, Len(strBuff) - 2) & vbNullChar
			End If
		Loop 
		FileClose(intFileHandle)
		
		'Assign return value
		If Len(strRet) Then
			GetSections = strRet & Chr(0)
		Else
			GetSections = New String(Chr(0), 2)
		End If
		
	End Function
	
	Private Function GetSectionsEx(ByRef DataArry() As String) As Short
		
		'The GetSectionsEx function is used to return an array of all the [Section]s
		'  in current file.
		
		
		'Get "normal" list of all [Section]'s
		Dim strSect As String
		strSect = GetSections()
		If Len(strSect) = 0 Then
			GetSectionsEx = 0
			Exit Function
		End If
		
		'Parse [Section]'s into table
		Dim intEntries As Short
		Dim intNull As Short
		Do While Asc(strSect)
			ReDim Preserve DataArry(intEntries)
			intNull = InStr(strSect, vbNullChar)
			DataArry(intEntries) = Left(strSect, intNull - 1)
			strSect = Mid(strSect, intNull + 1)
			intEntries = intEntries + 1
		Loop 
		
		'Make function assignment (number of [Sections]s added to table).
		GetSectionsEx = intEntries
		
	End Function
	
	Public Function GetString(ByVal EntryName As String) As String
		
		Dim DefaultStr As String
		'Bail if not initialized
		If Not mbooInitialized Then
			WarnAuthor()
			Exit Function
		End If
		
		'Retrieves Specific string Entry from INI.
		Dim strTemp As New VB6.FixedLengthString(Max_EntryBuffer)
		Dim lngRetVal As Integer
		If mbooUseWinINI Then
			lngRetVal = GetProfileString(mstrSectionName, EntryName, DefaultStr, strTemp.Value, Len(strTemp.Value))
		Else
			lngRetVal = GetPrivateProfileString(mstrSectionName, EntryName, DefaultStr, strTemp.Value, Len(strTemp.Value), mstrFullPath)
		End If
		
		If lngRetVal Then
			GetString = Left(strTemp.Value, lngRetVal)
		End If
		
	End Function
	
	Private Function GetTF(ByVal EntryName As String, ByVal DefaultInt As Short) As Boolean
		
		'Retrieves Specific Entry as either True/False.
		'local vars
		Dim strTF As String
		Dim strDefault As String
		
		'get string value from INI
		If DefaultInt Then
			strDefault = "True"
		Else
			strDefault = "False"
		End If
		
		strTF = GetString(EntryName)
		
		'interpret return string and translate to T/F.
		Select Case Trim(UCase(strTF))
			Case "YES", "Y", "TRUE", "T", "ON", "1", "-1"
				GetTF = True
			Case "NO", "N", "FALSE", "F", "OFF", "0"
				GetTF = False
			Case Else
				GetTF = False
		End Select
		
	End Function
	
	Public Sub FlushCache()
		
		'Bail if not initialized
		If Not mbooInitialized Then
			WarnAuthor()
			Exit Sub
		End If
		
		'To improve performance, Windows keeps a cached version of the most-recently
		'accessed initialization file. If that filename is specified and the other
		'three parameters are NULL, Windows flushes the cache
		Dim lngRetVal As Integer
		If mbooUseWinINI Then
			lngRetVal = WriteProfileString(CStr(0), CStr(0), CStr(0))
		Else
			lngRetVal = WritePrivateProfileString(CStr(0), 0, 0, mstrFullPath)
		End If
		
	End Sub
	
	Private Sub WarnAuthor()
		
		'Warn *PROGRAMMER* that there's a logic error!
		'MsgBox "[Section] and FileName Not Registered in Private.Ini!", vbInformation + vbOKOnly, "IniFile Logic Error"
		
	End Sub
	
	Public Sub IniRead(ByVal SectionName As String, ByVal EntryName As String, ByRef ReturnStr As String, ByVal FullPath As String)
		
		'One-shot read from Ini, more *work* than it's worth
		'It does not use any of the claas properties (except the UseWinIni value) - only what you pass to it!
		
		Dim lngRetVal As Integer
		Dim RetStr As New VB6.FixedLengthString(Max_EntryBuffer) 'Create an empty string to be filled
		Dim DefaultInt As Short
		Dim DefaultStr As String
		Dim ReadNumber As Boolean
		
		'DefaultStr = "KeyNull"
		'If ReadNumber Then     'we are looking for integer input
		'    If mbooUseWinINI Then
		'        ReadNumber = GetProfileInt(SectionName, EntryName, DefaultInt)
		'    Else
		'        ReadNumber = GetPrivateProfileInt(SectionName, EntryName, DefaultInt, FullPath)
		'    End If
		'Else
		'    If mbooUseWinINI Then
		'        lngRetVal = GetProfileString(SectionName, EntryName, DefaultStr, strTemp, Len(strTemp))
		'    Else
		lngRetVal = GetPrivateProfileString(SectionName, EntryName, DefaultStr, RetStr.Value, Len(RetStr.Value), FullPath)
		'    End If
		If lngRetVal Then
			ReturnStr = Left(RetStr.Value, lngRetVal)
		End If
		'End If
		
	End Sub
	Private Sub WinIniChanged() 'Orig.
		
		'Notify all other applications that Win.Ini has been changed
		Dim rtn As Integer
		'Rtn = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0&, ByVal mstrSectionName)
		
	End Sub
	
	Public Sub IniWrite(ByVal SectionName As String, ByVal EntryName As String, ByVal NewVal As String, ByVal FullPath As String)
		
		Dim lngRetVal As Integer
		
		'One-shot write to Private.Ini, more *work* than it's worth.
		'It does not use any of the claas properties (except the UseWinIni value) - only what you pass to it!
		
		'If mbooUseWinINI Then
		'    lngRetVal = WriteProfileString(SectionName, EntryName, NewVal)
		'Else
		lngRetVal = WritePrivateProfileString(SectionName, EntryName, NewVal, FullPath)
		'End If
		
		
	End Sub
	
	Private Function PutInt(ByVal EntryName As String, ByVal IntValue As Short) As Short
		
		'Bail if not initialized
		If Not mbooInitialized Then
			WarnAuthor()
			Exit Function
		End If
		
		'Write an integer to INI
		If mbooUseWinINI Then
			PutInt = WriteProfileString(mstrSectionName, EntryName, CStr(IntValue))
			WinIniChanged()
		Else
			PutInt = WritePrivateProfileString(mstrSectionName, EntryName, CStr(IntValue), mstrFullPath)
		End If
		
	End Function
	
	
	Public Function PutString(ByVal EntryName As String, ByVal StrValue As String) As Short
		
		'Bail if not initialized
		If Not mbooInitialized Then
			WarnAuthor()
			Exit Function
		End If
		
		'Write a string to INI
		If mbooUseWinINI Then
			PutString = WriteProfileString(mstrSectionName, EntryName, StrValue)
			WinIniChanged()
		Else
			PutString = WritePrivateProfileString(mstrSectionName, EntryName, StrValue, mstrFullPath)
		End If
		
	End Function
	
	Private Function PutTF(ByVal EntryName As String, ByVal IntValue As Short) As Boolean
		
		'Set an entry in .Ini to True/False
		'local vars
		Dim strTF As String
		
		'translate the value  to a string.
		If IntValue Then
			strTF = "True"
		Else
			strTF = "False"
		End If
		
		'enter the value in the INI
		PutTF = PutString(EntryName, strTF)
		
		If mbooUseWinINI Then WinIniChanged()
		
	End Function
	
	Private Function SectExist(ByVal SectionName As String) As Boolean
		
		'Retrieve list of all [Section]'s
		Dim strSect As String
		strSect = GetSections()
		If Len(strSect) = 0 Then
			SectExist = False
			Exit Function
		End If
		
		'Check for existence registered [Section]
		strSect = Chr(0) & UCase(strSect)
		If InStr(strSect, Chr(0) & UCase(SectionName) & Chr(0)) Then
			SectExist = True
		Else
			SectExist = False
		End If
		
	End Function
	
	Private Function StripComment(ByVal StrIn As String) As String 'orig
		Dim intRet As Short
		'Check for comment
		intRet = InStr(StrIn, ";")
		
		'Remove it if present
		If intRet = 1 Then
			'Whole string is a comment
			StripComment = ""
			Exit Function
		ElseIf intRet > 1 Then 
			'Strip comment
			StrIn = Left(StrIn, intRet - 1)
		End If
		
		'Trim any trailing space
		StripComment = Trim(StrIn)
		
	End Function
	
	'UPGRADE_NOTE: Class_Initialize は Class_Initialize_Renamed にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
	Private Sub Class_Initialize_Renamed() 'orig
		
#If DebugMode Then
		'UPGRADE_NOTE: 式 DebugMode が True に評価されなかったか、またはまったく評価されなかったため、#If #EndIf ブロックはアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"' をクリックしてください。
		'get the next available class ID, and print out
		'that the class was created successfully
		mlClassDebugID = GetNextClassDebugID()
		Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
#End If
		
		'UPGRADE_WARNING: App プロパティ App.EXEName には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
		If Trim(My.Application.Info.AssemblyName) = "" Then Exit Sub
		
		'Written to use long path names instead of the short path name returned by the App Object.
		'If the class is initialised early in the application the long path name will be supplied.
		'UPGRADE_WARNING: App プロパティ App.EXEName には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
		FullPath = GetLongPathName(My.Application.Info.DirectoryPath) & "\" & StrConv(My.Application.Info.AssemblyName, VbStrConv.ProperCase) & ".ini"
		
		mstrSectionName = "Default"
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate は Class_Terminate_Renamed にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
	Private Sub Class_Terminate_Renamed() 'Orig
		
		'the class is being destroyed
#If DebugMode Then
		'UPGRADE_NOTE: 式 DebugMode が True に評価されなかったか、またはまったく評価されなかったため、#If #EndIf ブロックはアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"' をクリックしてください。
		Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
#End If
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Function GetLongPathName(ByVal strShortName As String) As String 'Orig
		
		Dim strLongName As String
		Dim strTemp As String
		Dim intSlashPos As Short
		
		If Len(strShortName) < 1 Then Exit Function
		
		'Check to see if the root directory has been passed.
		If Right(strShortName, 2) = ":\" Then
			GetLongPathName = strShortName
			Exit Function
		End If
		
		'Clip off the trailing back-slash
		If Right(strShortName, 1) <> "\" Then
			strShortName = strShortName & "\" 'Left$(strShortName, Len(strShortName) - 1)
		End If
		
		intSlashPos = InStr(4, strShortName, "\")
		
		While intSlashPos
			'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
			strTemp = Dir(Left(strShortName, intSlashPos - 1), FileAttribute.Normal + FileAttribute.Hidden + FileAttribute.System + FileAttribute.Directory)
			If strTemp = "" Then
				GetLongPathName = ""
				Exit Function
			End If
			strLongName = strLongName & "\" & strTemp
			intSlashPos = InStr(intSlashPos + 1, strShortName, "\")
		End While
		
		GetLongPathName = Left(strShortName, 2) & strLongName
		
	End Function
	
	Public Function TokenCount(ByRef SearchStr As String, ByRef SepCh As Short) As Short
		TokenCount = 0
		If Len(SearchStr) = 0 Then Exit Function
		If SepCh < 0 Or SepCh > 255 Then Exit Function
		TokenCount = 1
		Dim sP1, sP0, WorkInWord As Integer
		Dim c As String
		c = Chr(SepCh)
		sP0 = 0 'the first seperator position=0
		WorkInWord = 1 'work in first one word that seperated by SepCh
		Do 
			sP1 = InStr(sP0 + 1, SearchStr, c)
			If sP1 = 0 Then 'no else find
				Exit Do
			Else
				sP0 = sP1
				WorkInWord = WorkInWord + 1
			End If
		Loop 
		TokenCount = WorkInWord
	End Function
	
	Public Function WhichToken(ByRef SearchStr As String, ByRef SepCh As Short, ByRef sFind As String, Optional ByRef lStart As Short = 0) As Short
		'to find the sfind is locate in which token number, 0 mean no find
		WhichToken = 0 'default no find
		If Len(SearchStr) = 0 Then Exit Function
		If Len(sFind) = 0 Then Exit Function
		If SepCh < 0 Or SepCh > 255 Then Exit Function
		Dim sP0 As Integer
		Dim c As String
		'UPGRADE_NOTE: IsMissing() は IsNothing() に変更されました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"' をクリックしてください。
		If IsNothing(lStart) Then lStart = 1
		If lStart < 1 Then lStart = 1
		If lStart > Len(SearchStr) Then Exit Function
		sP0 = InStr(lStart, SearchStr, sFind)
		If sP0 = 0 Then Exit Function
		c = Left(SearchStr, sP0 - 1)
		WhichToken = TokenCount(c, SepCh)
		If WhichToken = 0 Then WhichToken = 1
	End Function
	
	Public Function GetToken(ByRef SearchStr As String, ByRef SepCh As Short, ByRef WordIndex As Short) As String
		GetToken = ""
		If Len(SearchStr) = 0 Then Exit Function
		If WordIndex <= 0 Then Exit Function
		If SepCh < 0 Or SepCh > 255 Then Exit Function
		Dim sP1, sP0, WorkInWord As Short
		Dim c As String
		c = Chr(SepCh)
		sP0 = 0 'the first seperator position=0
		WorkInWord = 1 'work in first one word that seperated by SepCh
		Do 
			sP1 = InStr(sP0 + 1, SearchStr, c)
			If sP1 = 0 Then 'no else find
				sP1 = Len(SearchStr) + 1
				If sP0 = 0 Then 'no find seperator in input string
					GetToken = SearchStr
					Exit Function
				End If
				If WorkInWord <> WordIndex Then
					GetToken = ""
					Exit Function
				End If
				Exit Do
			Else
				If WorkInWord = WordIndex Then Exit Do
				sP0 = sP1
				WorkInWord = WorkInWord + 1
			End If
		Loop 
		GetToken = Mid(SearchStr, sP0 + 1, sP1 - sP0 - 1)
	End Function
	
	Public Sub ReplaceToken(ByRef SearchStr As String, ByRef SepCh As Short, ByRef WordIndex As Short, ByRef sNewToken As String)
		Dim NewToken As Object
		Dim sP1, sP0, WorkInWord As Short
		Dim c As String
		Dim OriString As String
		OriString = SearchStr
		If Len(SearchStr) = 0 Then Exit Sub
		If WordIndex <= 0 Then Exit Sub
		If WordIndex > TokenCount(SearchStr, SepCh) Then Exit Sub
		If SepCh < 0 Or SepCh > 255 Then Exit Sub
		c = Chr(SepCh)
		sP0 = 0 'the first seperator position=0
		WorkInWord = 1 'work in first one word that seperated by SepCh
		Do 
			sP1 = InStr(sP0 + 1, OriString, c)
			If sP1 = 0 Then 'no else find
				sP1 = Len(OriString) + 1
				If sP0 = 0 Then 'no find seperator in input string
					'UPGRADE_WARNING: オブジェクト NewToken の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
					SearchStr = NewToken
					Exit Sub
				End If
				If WorkInWord <> WordIndex Then Exit Sub
				Exit Do
			Else
				If WorkInWord = WordIndex Then Exit Do
				sP0 = sP1
				WorkInWord = WorkInWord + 1
			End If
		Loop 
		SearchStr = Left(OriString, sP0) & sNewToken & Right(OriString, Len(OriString) - sP1 + 1)
	End Sub
	
	Public Sub DeleteToken(ByRef SearchStr As String, ByRef SepCh As Short, ByRef WordIndex As Short)
		Dim sP1, sP0, WorkInWord As Short
		Dim c As String
		Dim OriString As String
		OriString = SearchStr
		If Len(SearchStr) = 0 Then Exit Sub
		If WordIndex <= 0 Then Exit Sub
		If WordIndex > TokenCount(SearchStr, SepCh) Then Exit Sub
		If SepCh < 0 Or SepCh > 255 Then Exit Sub
		c = Chr(SepCh)
		sP0 = 0 'the first seperator position=0
		WorkInWord = 1 'work in first one word that seperated by SepCh
		Do 
			sP1 = InStr(sP0 + 1, OriString, c)
			If sP1 = 0 Then 'no else find
				sP1 = Len(OriString) + 1
				If sP0 = 0 Then 'no find seperator in input string
					SearchStr = ""
					Exit Sub
				End If
				If WorkInWord <> WordIndex Then Exit Sub
				Exit Do
			Else
				If WorkInWord = WordIndex Then Exit Do
				sP0 = sP1
				WorkInWord = WorkInWord + 1
			End If
		Loop 
		If sP0 = 0 Then
			SearchStr = Mid(OriString, sP1 + 1)
		Else
			If Len(OriString) < sP1 Then
				sP1 = Len(OriString)
				sP0 = sP0 - 1
				If sP0 < 0 Then sP0 = 0
			End If
			SearchStr = Left(OriString, sP0) & Right(OriString, Len(OriString) - sP1)
		End If
	End Sub
End Class