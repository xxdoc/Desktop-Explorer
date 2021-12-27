VERSION 5.00
Begin VB.Form DesktopExplorerWindow 
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "Desktop Explorer.frx":0000
   ScaleHeight     =   13.938
   ScaleMode       =   4  'Character
   ScaleWidth      =   39
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox DesktopProcessListBox 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Click on an item to bring the process' window to the foreground."
      Top             =   1920
      Width           =   4335
   End
   Begin VB.ListBox DesktopListBox 
      Height          =   1230
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "Click on an item to display the associated processes."
      Top             =   360
      Width           =   2055
   End
   Begin VB.ListBox WindowStationListBox 
      Height          =   1230
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Click on an item to display the associated desktops."
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label DesktopProcessesLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop Processes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label DesktopsLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Desktops:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label WindowStationsLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Window Stations:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu ProgramMainMenu 
      Caption         =   "&Program"
      Begin VB.Menu IgnoreAPIErrorsMenu 
         Caption         =   "&Ignore API Errors"
         Shortcut        =   ^E
      End
      Begin VB.Menu ProgramMainMenuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu ExecuteMenu 
         Caption         =   "&Execute"
         Shortcut        =   ^R
      End
      Begin VB.Menu ProgramMainMenuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu ProgramMainMenuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu QuitMenu 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu DesktopsMainMenu 
      Caption         =   "&Desktops"
      Begin VB.Menu CreateDesktopMenu 
         Caption         =   "&Create Desktop"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "DesktopExplorerWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's interface.
Option Explicit


'This procedure gives the command to display a list of desktops for the selected window station.
Private Sub RefreshDesktopList()
On Error GoTo ErrorTrap
   With Me.WindowStationListBox
      If Not .ListIndex = NO_INDEX Then DisplayDesktops WindowStationList(, CLng(.ListIndex)), Me.DesktopListBox
   End With
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub


'This procedure gives the command to create a new desktop with the name specified by the user.
Private Sub CreateDesktopMenu_Click()
On Error GoTo ErrorTrap
Dim NewDesktop As String
   
   NewDesktop = InputBox$("New desktop name:")
   If Not NewDesktop = vbNullString Then
      AddDesktop NewDesktop
      RefreshDesktopList
   End If
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub


'This procedure gives the command to display the statistics for the selected desktop.
Private Sub DesktopListBox_Click()
On Error GoTo ErrorTrap
   DisplayDesktopProcesses DesktopList(, DesktopListBox.ListIndex), DesktopProcessListBox
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub

'This procedure gives the command to activate the selected desktop.
Private Sub DesktopListBox_DblClick()
On Error GoTo ErrorTrap
   ActivateDesktop DesktopList(, DesktopListBox.ListIndex)
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub




'This procedure attempts to bring the process' main window to the foreground.
Private Sub DesktopProcessListBox_DblClick()
On Error GoTo ErrorTrap
Dim EndPosition As Long
Dim Index As Long
Dim StartPosition As Long
Dim ProcessId As Long
Dim WindowHandle As Long

   With DesktopProcessListBox
      StartPosition = InStr(.List(.ListIndex), "(")
      If StartPosition > 0 Then
         EndPosition = InStr(StartPosition + 1, .List(.ListIndex), ")")
         If EndPosition > 0 Then
            ProcessId = CLng(Val(Trim$(Mid$(.List(.ListIndex), StartPosition + 1, (EndPosition - StartPosition) - 1))))
            Index = SearchProcessList(, Id:=ProcessId)
            If Not Index = NO_INDEX Then
               WindowHandle = DesktopProcessList(, , , Index).WindowHandle
               If WindowHandle = NO_HANDLE Then
                  MsgBox "This process does not have a window.", vbInformation
               Else
                  CheckForError EnableWindow(WindowHandle, CLng(True))
                  CheckForError ShowWindow(WindowHandle, SW_SHOWNA)
                  CheckForError BringWindowToTop(WindowHandle)
                  If CBool(CheckForError(IsIconic(WindowHandle))) Then CheckForError ShowWindow(WindowHandle, SW_RESTORE)
               End If
            End If
         End If
      End If
   End With
  
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub



'This procedure requests the user to specify an application to be executed.
Private Sub ExecuteMenu_Click()
On Error GoTo ErrorTrap
Dim Path As String
   
   Path = InputBox$("Path:")
   If Not Path = vbNullString Then
      Shell Path, vbNormalFocus
      DisplayDesktopProcesses DesktopList(, DesktopListBox.ListIndex), DesktopProcessListBox
   End If
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub



'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2
   
   With App
      Me.Caption = .Title & ", v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName & " - User: " & GetUser()
   End With
   
   IgnoreAPIErrorsMenu.Checked = IgnoreAPIErrors()
   
   DisplayWindowStations Me.WindowStationListBox
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub

'This procedure adjusts this window's controls and objects to its new size.
Private Sub Form_Resize()
On Error Resume Next
   With Me
      .WindowStationsLabel.Left = 1
      .WindowStationListBox.Left = .WindowStationsLabel.Left
      .WindowStationListBox.Width = (.ScaleWidth / 2) - 2
      .WindowStationListBox.Height = (.ScaleHeight - 2) / 2
   
      .DesktopsLabel.Left = (.ScaleWidth / 2) + 1
      .DesktopListBox.Left = .DesktopsLabel.Left
      .DesktopListBox.Width = .WindowStationListBox.Width
      .DesktopListBox.Height = .WindowStationListBox.Height
   
      .DesktopProcessesLabel.Top = .WindowStationListBox.Top + .WindowStationListBox.Height + 0.5
      .DesktopProcessListBox.Width = .ScaleWidth - 2
      .DesktopProcessListBox.Height = .WindowStationListBox.Height - 1
      .DesktopProcessListBox.Top = (.ScaleHeight - 1) - .DesktopProcessListBox.Height
   End With
End Sub



'This procedure closes this program when this window is closed.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   Unload Me
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub


'This procedure toggles the ignore API errors setting.
Private Sub IgnoreAPIErrorsMenu_Click()
On Error GoTo ErrorTrap
   IgnoreAPIErrors NewIgnoreStatus:=(Not IgnoreAPIErrors())
   IgnoreAPIErrorsMenu.Checked = IgnoreAPIErrors()
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub


'This procedure displays information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
   MsgBox App.Comments, vbInformation
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub


'This procedure closes this window.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap
   Unload Me
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub

'This procedure gives the command to refresh the list of desktops for the selected window station.
Private Sub WindowStationListBox_Click()
On Error GoTo ErrorTrap
   RefreshDesktopList
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) Then Resume EndProcedure
End Sub


