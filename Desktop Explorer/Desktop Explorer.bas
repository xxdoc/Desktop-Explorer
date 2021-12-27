Attribute VB_Name = "DesktopExplorerModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants used by this program:
Public Const SW_SHOWNA As Long = &H8&
Public Const SW_RESTORE As Long = &H9&
Private Const CCHDEVICENAME As Long = &H20&
Private Const CCHFORMNAME As Long = &H20&
Private Const DF_ALLOWOTHERACCOUNTHOOK As Long = &H1&
Private Const ERROR_SUCCESS As Long = &H0&
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const GENERIC_ALL As Long = &H10000000
Private Const PROCESS_QUERY_INFORMATION As Long = &H400&
Private Const WINSTA_ALL_ACCESS As Long = &H37F&

'The Microsoft Windows API structures used by this program:
Private Type DEVMODE
   dmDeviceName As String * CCHDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCHFORMNAME
   dmUnusedPadding As Integer
   dmBitsPerPel As Long
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

'The Microsoft Windows API functions used by this program:
Public Declare Function BringWindowToTop Lib "User32.dll" (ByVal hwnd As Long) As Long
Public Declare Function EnableWindow Lib "User32.dll" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function IsIconic Lib "User32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "User32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CloseDesktop Lib "User32.dll" (ByVal hDesktop As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CloseWindowStation Lib "User32.dll" (ByVal hWinSta As Long) As Long
Private Declare Function CreateDesktopA Lib "User32.dll" (ByVal lpszDesktop As String, ByVal lpszDevice As String, pDevmode As DEVMODE, ByVal dwFlags As Long, ByVal dwDesiredAccess As Long, lpsa As SECURITY_ATTRIBUTES) As Long
Private Declare Function CreateProcessA Lib "Kernel32.dll" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function EnumDesktopsA Lib "User32.dll" (ByVal hWinSta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumDesktopWindows Lib "User32.dll" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function EnumProcesses Lib "Psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function EnumWindowStationsA Lib "User32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetCurrentProcessId Lib "Kernel32.dll" () As Long
Private Declare Function GetThreadDesktop Lib "User32.dll" (ByVal dwThread As Long) As Long
Private Declare Function GetUserNameA Lib "Advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32.dll" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function lstrcpynA Lib "Kernel32.dll" (ByVal lpString1 As Any, ByVal lpString2 As Any, ByVal iMaxLength As Long) As Long
Private Declare Function lstrlenA Lib "Kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Function OpenDesktopA Lib "User32.dll" (ByVal lpszDesktop As String, ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function OpenWindowStationA Lib "User32.dll" (ByVal lpszWinSta As String, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Private Declare Function QueryFullProcessImageNameW Lib "Kernel32.dll" (ByVal ProcessHandle As Long, ByVal ProcessFlags As Long, ByVal EXEName As Long, ByRef BufferLength As Long) As Long
Private Declare Function SwitchDesktop Lib "User32.dll" (ByVal hDesktop As Long) As Long


'The constants and structures used by this program:
Public Type ProcessStr
   Id As Long             'Defines a process' id.
   Path As String         'Defines a process' path.
   WindowHandle As Long   'Defines a process' main window handle.
End Type

Public Const NO_HANDLE As Long = 0         'Defines a null handle.
Public Const NO_INDEX As Long = -1         'Defines the absence of a selection.
Private Const NO_ID As Long = 0            'Defines the absence of an identification number.
Private Const MAX_PATH As Long = 260       'Defines the maximum number of characters allowed for a file path.
Private Const MAX_STRING As Long = 65535   'Defines the maximum length allowed for a string.

'This procedure activates the specified desktop.
Public Sub ActivateDesktop(Desktop As String)
On Error GoTo ErrorTrap
Dim CurrentDesktopH As Long
Dim DesktopH As Long
Dim ProcessId As Long

      If Dir$(GetFullProgramPath(), vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = vbNullString Then
         MsgBox "Cannot find """ & GetFullProgramPath() & """.", vbExclamation
      Else
         CurrentDesktopH = CheckForError(GetThreadDesktop(App.ThreadID))
         DesktopH = CheckForError(OpenDesktopA(Desktop, CLng(0), CLng(False), GENERIC_ALL))
        
         If Not DesktopH = NO_HANDLE Then
            DesktopProcessList , , , InitializeList:=True
            CheckForError EnumDesktopWindows(DesktopH, AddressOf WindowHandler, CLng(0))
      
            If SearchProcessList(GetProcessPath(CheckForError(GetCurrentProcessId()))) = NO_INDEX Then
               ProcessId = StartProcess(GetFullProgramPath(), Desktop).dwProcessId
               If Not ProcessId = NO_ID Then
                  CheckForError SwitchDesktop(DesktopH)
                  WaitForProcess ProcessId
               End If
            End If
         
            CheckForError CloseDesktop(DesktopH)
            CheckForError SwitchDesktop(CurrentDesktopH)
         End If
      End If
      
EndProcedure:
      Exit Sub
      
ErrorTrap:
      If HandleError() = vbRetry Then Resume
      If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub



'This procedure creates a new desktop with the specified name.
Public Sub AddDesktop(NewDesktop As String)
On Error GoTo ErrorTrap
Dim DeviceMode As DEVMODE
Dim Security As SECURITY_ATTRIBUTES

   CheckForError CreateDesktopA(NewDesktop, vbNullString, DeviceMode, DF_ALLOWOTHERACCOUNTHOOK, GENERIC_ALL, Security)
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub



'This procedure checks whether an error has occurred during the most recent API call.
Public Function CheckForError(ReturnValue As Long) As Long
On Error GoTo ErrorTrap
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String

   ErrorCode = Err.LastDllError
   Err.Clear
   
   If Not IgnoreAPIErrors() Then
      If Not ErrorCode = ERROR_SUCCESS Then
         Description = String$(MAX_STRING, vbNullChar)
         Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
         If Length = 0 Then
            Description = "No description."
         ElseIf Length > 0 Then
            Description = Left$(Description, Length - 1)
         End If
        
         Message = "API error code: " & CStr(ErrorCode) & " - " & Description
         Message = Message & "Return value: " & CStr(ReturnValue)
   
         MsgBox Message, vbExclamation
      End If
   End If
   
EndProcedure:
   CheckForError = ReturnValue
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function



'This procedure handles any instances of desktops found.
Private Function DesktopHandler(ByVal lpszDesktop As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorTrap
Dim Desktop As String

   Desktop = String$(CheckForError(lstrlenA(lpszDesktop)), vbNullChar)
   CheckForError lstrcpynA(Desktop, lpszDesktop, Len(Desktop) + 1)
   DesktopList NewDesktop:=Desktop
    
EndProcedure:
   DesktopHandler = CLng(True)
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function


'This procedure manages the list of processes with windows on a specific desktop.
Public Function DesktopProcessList(Optional NewPath As String = vbNullString, Optional NewId As Long = NO_ID, Optional NewHandle As Long, Optional Index As Long = NO_INDEX, Optional InitializeList As Boolean = False) As ProcessStr
On Error GoTo ErrorTrap
Dim ProcessPath As ProcessStr
Static ProcessPaths() As ProcessStr

   With ProcessPath
      .Id = NO_ID
      .Path = vbNullString
      .WindowHandle = NO_HANDLE
   End With
   
   If InitializeList Then
      ReDim ProcessPaths(0 To 0) As ProcessStr
   ElseIf Not (NewPath = vbNullString Or NewId = NO_ID) Then
      If SearchProcessList(, Id:=NewId) = NO_INDEX Then
         With ProcessPaths(UBound(ProcessPaths()))
            .Id = NewId
            .Path = NewPath
            .WindowHandle = NewHandle
         End With
         ReDim Preserve ProcessPaths(LBound(ProcessPaths()) To UBound(ProcessPaths()) + 1) As ProcessStr
      End If
   ElseIf Not Index = NO_INDEX Then
      If Index >= LBound(ProcessPaths()) And Index <= UBound(ProcessPaths()) Then ProcessPath = ProcessPaths(Index)
   End If
EndProcedure:
   DesktopProcessList = ProcessPath
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function

'This procedure displays the processes that have windows on the specified desktop.
Public Sub DisplayDesktopProcesses(Desktop As String, List As Object)
On Error GoTo ErrorTrap
Dim Index As Long
Dim Process As ProcessStr

   GetDesktopProcesses Desktop
   
   Index = 0
   List.Clear
   Process = DesktopProcessList(, , , Index)
   Do Until Process.Path = vbNullString
      List.AddItem "(" & Space$(5 - Len(CStr(Process.Id))) & CStr(Process.Id) & ") " & Process.Path
      Index = Index + 1
      Process = DesktopProcessList(, , , Index)
   Loop
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub



'This procedure fills the specified list with a list of the desktops attached to a window station.
Public Sub DisplayDesktops(WindowStation As String, List As Object)
On Error GoTo ErrorTrap
Dim Desktop As String
Dim Index As Long

   GetDesktopList WindowStation
   
   Index = 0
   Desktop = DesktopList(, Index)
   List.Clear
   Do Until Desktop = vbNullString
      List.AddItem Desktop
      DoEvents
      Index = Index + 1
      Desktop = DesktopList(, Index)
   Loop
   
   If List.ListCount > 0 Then List.ListIndex = 0
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub

'This procedure adds the names of the active window stations to the specified list.
Public Sub DisplayWindowStations(List As Object)
On Error GoTo ErrorTrap
Dim Index As Long
Dim WindowStation As String

   WindowStationList , , InitializeList:=True
   CheckForError EnumWindowStationsA(AddressOf WindowStationHandler, CLng(0))
   
   Index = 0
   WindowStation = WindowStationList(, Index)
   List.Clear
   Do Until WindowStation = vbNullString
      List.AddItem WindowStation
      DoEvents
      Index = Index + 1
      WindowStation = WindowStationList(, Index)
   Loop
   
   If List.ListCount > 0 Then List.ListIndex = 0
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub

'This procedure manages the list of desktops.
Public Function DesktopList(Optional NewDesktop As String = vbNullString, Optional Index As Long = NO_INDEX, Optional InitializeList As Boolean = False) As String
On Error GoTo ErrorTrap
Dim Desktop As String
Static Desktops() As String

   Desktop = vbNullString
   
   If InitializeList Then
      ReDim Desktops(0 To 0) As String
   Else
      If Not NewDesktop = vbNullString Then
         Desktops(UBound(Desktops())) = NewDesktop
         ReDim Preserve Desktops(LBound(Desktops()) To UBound(Desktops()) + 1) As String
      ElseIf Not Index = NO_INDEX Then
         If Index >= LBound(Desktops()) And Index <= UBound(Desktops()) Then Desktop = Desktops(Index)
      End If
   End If
EndProcedure:
   DesktopList = Desktop
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function



'This procedure creates a list of desktops attached to the specified window station.
Private Sub GetDesktopList(WindowStation As String)
On Error GoTo ErrorTrap
Dim WindowStationH As Long

   DesktopList , , InitializeList:=True
   WindowStationH = CheckForError(OpenWindowStationA(WindowStation, CLng(False), WINSTA_ALL_ACCESS))
   If Not WindowStationH = NO_HANDLE Then
      CheckForError EnumDesktopsA(WindowStationH, AddressOf DesktopHandler, CLng(0))
      CheckForError CloseWindowStation(WindowStationH)
   End If
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub


'This procedure gets the processes that have windows on the specified desktop.
Public Sub GetDesktopProcesses(Desktop As String)
On Error GoTo ErrorTrap
Dim DesktopH As Long

   DesktopProcessList , , , , InitializeList:=True
   
   DesktopH = CheckForError(OpenDesktopA(Desktop, CLng(0), CLng(False), GENERIC_ALL))
   If Not DesktopH = NO_HANDLE Then
      CheckForError EnumDesktopWindows(DesktopH, AddressOf WindowHandler, CLng(0))
      CheckForError CloseDesktop(DesktopH)
   End If
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub


'This procedure returns the full path of this program's executable.
Private Function GetFullProgramPath() As String
On Error GoTo ErrorTrap
Dim Path As String

   Path = App.Path
   If Not Right$(Path, 1) = "\" Then Path = Path & "\"
   Path = Path & App.EXEName
   Path = LCase$(Path)
   If Not Right$(Path, 4) = ".exe" Then Path = Path & ".exe"
   
EndProcedure:
   GetFullProgramPath = Path
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function

'This procedure returns a list of ids for all running processes.
Private Function GetProcessIds() As Long()
On Error GoTo ErrorTrap
Dim ProcessIds() As Long
Dim Size As Long
Dim SizeUsed As Long

   ReDim ProcessIds(0 To 0) As Long
   Do
      Size = (UBound(ProcessIds()) + 1) * Len(ProcessIds(LBound(ProcessIds())))
      CheckForError EnumProcesses(ProcessIds(0), Size, SizeUsed)
      If Size > SizeUsed Then Exit Do
      ReDim Preserve ProcessIds(LBound(ProcessIds()) To UBound(ProcessIds()) + 1) As Long
   Loop
   
EndProcedure:
   GetProcessIds = ProcessIds()
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function



'This procedure returns the path of the process with the specified id.
Private Function GetProcessPath(ProcessId As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim Path As String
Dim ProcessH As Long

   Path = vbNullString
   ProcessH = CheckForError(OpenProcess(PROCESS_QUERY_INFORMATION, CLng(False), ProcessId))
   If Not ProcessH = NO_HANDLE Then
      Path = String$(MAX_PATH, vbNullChar)
      Length = Len(Path)
      CheckForError QueryFullProcessImageNameW(ProcessH, CLng(0), StrPtr(Path), Length)
      Path = Left$(Path, Length)
      CheckForError CloseHandle(ProcessH)
   End If
   
EndProcedure:
   GetProcessPath = Path
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function

'This procedure returns the name for the account under which this program is being executed.
Public Function GetUser() As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim UserName As String

   UserName = String$(MAX_STRING, vbNullChar)
   Length = Len(UserName)
   CheckForError GetUserNameA(UserName, Length)
   If Length > 0 Then UserName = Left$(UserName, Length - 1) Else UserName = vbNullString
   
EndProcedure:
   GetUser = UserName
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function


'This procedure handles any errors that occur.
Public Function HandleError(Optional DoNotAsk As Boolean = False) As Long
Dim Message As String
Static Choice As Long

   If Not DoNotAsk Then
      Message = "Error: " & CStr(Err.Number) & vbCr
      Message = Message & Err.Description
      
      On Error Resume Next
      
      Choice = MsgBox(Message, vbAbortRetryIgnore Or vbDefaultButton2 Or vbExclamation)
      If Choice = vbAbort Then End
   End If
   
   HandleError = Choice
End Function



'This procedure sets/returns whether API errors are ignored or not.
Public Function IgnoreAPIErrors(Optional NewIgnoreStatus As Variant) As Boolean
On Error GoTo ErrorTrap
Static CurrentIgnoreStatus As Boolean
   
   If Not IsMissing(NewIgnoreStatus) Then CurrentIgnoreStatus = NewIgnoreStatus
   
EndProcedure:
   IgnoreAPIErrors = CurrentIgnoreStatus
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function



'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   IgnoreAPIErrors NewIgnoreStatus:=True
   
   DesktopExplorerWindow.Show
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub



'This procedure returns the index of the specified process path.
Public Function SearchProcessList(Optional Path As String, Optional Id As Long = NO_ID) As Long
On Error GoTo ErrorTrap
Dim FoundAt As Long
Dim Index As Long
Dim Process As ProcessStr

   FoundAt = NO_INDEX
   Index = 0
   Do
      Process = DesktopProcessList(, , , Index)
      If Process.Path = vbNullString Then Exit Do

      If ((Not Path = vbNullString) And LCase$(Trim$(Process.Path)) = LCase$(Trim$(Path))) Or ((Not Id = NO_ID) And Process.Id = Id) Then
         FoundAt = Index
         Exit Do
      End If

      Index = Index + 1
   Loop
   
EndProcedure:
   SearchProcessList = FoundAt
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function

'This procedure executes the specified application on the specified desktop.
Private Function StartProcess(Path As String, Desktop As String) As PROCESS_INFORMATION
On Error GoTo ErrorTrap
Dim Process As PROCESS_INFORMATION
Dim Security As SECURITY_ATTRIBUTES
Dim StartUp As STARTUPINFO

   With Security
      .bInheritHandle = CLng(True)
      .lpSecurityDescriptor = CLng(0)
      .nLength = Len(Security)
   End With
   
   With StartUp
      .cb = Len(StartUp)
      .lpDesktop = Desktop
   End With
   
   CheckForError CreateProcessA(vbNullString, Path, Security, Security, CLng(True), CLng(0), CLng(0), vbNullString, StartUp, Process)
   
EndProcedure:
   StartProcess = Process
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function


'This procedure waits until the specified process is terminated.
Private Sub WaitForProcess(ProcessId As Long)
On Error GoTo ErrorTrap
Dim Index As Long
Dim ProcessIds() As Long
Dim Wait As Boolean

   Do While DoEvents() > 0
      ProcessIds() = GetProcessIds()
      Wait = False
      For Index = LBound(ProcessIds()) To UBound(ProcessIds())
         If ProcessIds(Index) = ProcessId Then
            Wait = True
            Exit For
         End If
      Next Index
      If Not Wait Then Exit Do
   Loop
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub



'This procedure handles any active windows on a specific desktop.
Private Function WindowHandler(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorTrap
Dim ProcessId As Long

   CheckForError GetWindowThreadProcessId(hwnd, ProcessId)
   If Not ProcessId = NO_ID Then DesktopProcessList NewPath:=GetProcessPath(ProcessId), NewId:=ProcessId, NewHandle:=hwnd
   
EndProcedure:
   WindowHandler = CLng(True)
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function

'This procedure handles any instances of window stations found.
Private Function WindowStationHandler(ByVal lpszWindowStation As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorTrap
Dim WindowStation As String

   WindowStation = String$(CheckForError(lstrlenA(lpszWindowStation)), vbNullChar)
   CheckForError lstrcpynA(WindowStation, lpszWindowStation, Len(WindowStation) + 1)
   WindowStationList NewWindowStation:=WindowStation
   
EndProcedure:
   WindowStationHandler = CLng(True)
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function


'This procedure manages the list of active window stations.
Public Function WindowStationList(Optional NewWindowStation As String = vbNullString, Optional Index As Long = NO_INDEX, Optional InitializeList As Boolean = False) As String
On Error GoTo ErrorTrap
Dim WindowStation As String
Static WindowStations() As String

   WindowStation = vbNullString
   
   If InitializeList Then
      ReDim WindowStations(0 To 0) As String
   Else
      If Not NewWindowStation = vbNullString Then
         WindowStations(UBound(WindowStations())) = NewWindowStation
         ReDim Preserve WindowStations(LBound(WindowStations()) To UBound(WindowStations()) + 1) As String
      ElseIf Not Index = NO_INDEX Then
         If Index >= LBound(WindowStations()) And Index <= UBound(WindowStations()) Then WindowStation = WindowStations(Index)
      End If
   End If
EndProcedure:
   WindowStationList = WindowStation
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Function

