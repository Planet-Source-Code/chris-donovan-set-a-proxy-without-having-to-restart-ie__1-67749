VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proxy Changer"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   705
   ClientWidth     =   4215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4290
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5293
            MinWidth        =   5293
            Text            =   "Current Proxy: Disabled"
            TextSave        =   "Current Proxy: Disabled"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2118
            MinWidth        =   2118
            Text            =   "Proxies: 0"
            TextSave        =   "Proxies: 0"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstView 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IP Address"
         Object.Width           =   4127
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Port"
         Object.Width           =   2381
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgDisabled 
      Height          =   480
      Left            =   0
      Picture         =   "frmMain.frx":000C
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEnabled 
      Height          =   480
      Left            =   0
      Picture         =   "frmMain.frx":08D6
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAddProxy 
         Caption         =   "Add Proxy"
      End
      Begin VB.Menu mnuLoadProxy 
         Caption         =   "Load Proxy List"
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisableProxy1 
         Caption         =   "Turn Proxy Off"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuMinimized 
         Caption         =   "Start Proxy Changer Minimized"
      End
   End
   Begin VB.Menu mnuListview 
      Caption         =   "Listview"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuUseProxy 
         Caption         =   "Use Selected Proxy"
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveProxy 
         Caption         =   "Remove Selected Proxies"
      End
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "Remove All Proxies"
      End
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu mnuSetProxy 
         Caption         =   "Set Proxy"
      End
      Begin VB.Menu mnuSetRandom 
         Caption         =   "Set Random Proxy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisableProxy2 
         Caption         =   "Turn Proxy Off"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit1 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bLoadMinimized As Boolean

'=================================
' General
'=================================

Private Sub Form_Load()
    SetIcon Me.hWnd, "AAA" 'Set 32 bit app icon (compile app to see the icon in captionbar).
    SetHook 'Hook for centering CommonDialog window.

    Call ShellTrayIconAdd(Me.hWnd, imgDisabled.Picture, "Proxy Changer - Disabled") 'Add icon to systemtray.

    bLoadMinimized = GetSetting(App.EXEName, "Options", "LoadMinimized", "False") 'Read setting from registry.
    mnuMinimized.Checked = bLoadMinimized
    If bLoadMinimized = True Then 'Load app minimized or not.
       Me.Left = (Screen.Width - Me.Width) / 2
       Me.Top = (Screen.Height - Me.Height) / 2
       WindowState = vbMinimized
    Else
       Me.Show
       Me.Refresh
    End If

    Call DisableConnectionProxy("") 'Make sure proxy is disabled.
    
    Call ReadProxyList(App.Path & "\ProxyListSaved.txt", lstView) 'Read the saved proxy list if it exists.
    
    If lstView.ListItems.Count > 0 Then mnuSetRandom.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseHook 'Unhook.
    
    Call DisableConnectionProxy("") 'Make sure proxy is disabled.
    
    Call SaveProxyList(App.Path & "\ProxyListSaved.txt", lstView) 'Save Listview with proxies to text file.
  
    If defWindowProc <> 0 Then
        Call ShellTrayIconRemove(Me.hWnd) 'Remove the icon from the systemtray.
        UnSubClass Me.hWnd 'Remove subclassing.
    End If
       
    Cancel = False 'Ensure unloading proceeds.
End Sub

'Remove app from taskbar when minimized.
Private Sub Form_Resize()
    If Me.WindowState = 1 Then
        Me.Hide
    End If
End Sub

Private Sub lstView_DblClick()
    mnuUseProxy_Click 'Call the mnuUseProxy sub, because that one contains the code to set
                      'the proxy, so we don't have to add all the code in this sub again.
End Sub

Private Sub lstView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim itmClicked As ListItem
    
    If Button = vbRightButton Then 'Only if right mouse button hit.
        Set itmClicked = lstView.HitTest(x, y) 'Find the list item that matches the mouse coordinates.
        If Not (itmClicked Is Nothing) Then 'Ensure that a valid object was clicked.
            itmClicked.Selected = True 'Select it.
            PopupMenu mnuListview 'Show listview popupmenu.
        End If
    End If

    Set itmClicked = Nothing 'Clean up.
End Sub

'=================================
' Application menu
'=================================

Private Sub mnuAddProxy_Click()
    frmAddProxy.Show vbModal 'Show the frmAddProxy form.
End Sub

'Load a list with proxies.
Private Sub mnuLoadProxy_Click()
Dim sProxy As String
Dim arrPart() As String
Dim lvItem As ListItem
Dim i As Integer
Dim j As Integer
  
  'Load the proxy list into the ListView.
  With CommonDialog
      .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNFileMustExist
      .Filter = "Proxy List (*.txt)|*.txt"
      .DialogTitle = "Select proxy list"
      .InitDir = App.Path
      .ShowOpen
        If .FileName <> "" Then
            Open .FileName For Input As #1
                Do While Not EOF(1)
                    Line Input #1, sProxy
                      If ValidProxy(sProxy) Then 'Only add the proxy if it has a valid IP and port number.
                        arrPart = Split(sProxy, ":") 'Split the IP address and port number on the ':' delimiter.
                        Set lvItem = lstView.ListItems.Add(, , arrPart(0)) 'Add IP address to the first column.
                        lvItem.ListSubItems.Add , , arrPart(1) 'Add port number to the second column.
                        mnuSetRandom.Enabled = True
                      End If
                Loop
            Close #1
        End If
  End With
  
  'Remove duplicates.
  With lstView
      For i = 1 To .ListItems.Count
          For j = .ListItems.Count To (i + 1) Step -1 'Loop through the listview backwards.
              If .ListItems(j) = .ListItems(i) Then 'If item (IP address) already exists....
                  .ListItems.Remove j '... then remove it from the listview.
              End If
          Next
      Next
  End With

  StatusBar.Panels(2).Text = "Proxies: " & lstView.ListItems.Count
End Sub

'Disable proxy.
Private Sub mnuDisableProxy1_Click()
    mnuDisableProxy2_Click 'Call the other mnuDisableProxy sub, because that one contains the code to
                           'disable the proxy, so we don't have to add all the code in this sub again.
End Sub

'Close app.
Private Sub mnuExit_Click()
  Unload Me
End Sub

'Check or uncheck to load minimized on next startup.
Private Sub mnuMinimized_Click()
  If mnuMinimized.Checked = False Then
    mnuMinimized.Checked = True
    SaveSetting App.EXEName, "Options", "LoadMinimized", "True"
  ElseIf mnuMinimized.Checked = True Then
    mnuMinimized.Checked = False
    SaveSetting App.EXEName, "Options", "LoadMinimized", "False"
  End If
End Sub

'=================================
' Listview popup menu
'=================================

'Set the proxy
Private Sub mnuUseProxy_Click()
Dim i As Integer
Dim sFullAddress As String
    
    i = lstView.SelectedItem.Index
    sFullAddress = lstView.ListItems.Item(i) & ":" & lstView.ListItems.Item(i).SubItems(1) 'Get IP:Port from Listview.
    
    If SetConnectionOptions("", sFullAddress) Then
        'Update statusbar.
        StatusBar.Panels(1).Text = "Current Proxy: " & sFullAddress
        
        'Enable the menu options.
        mnuDisableProxy1.Enabled = True
        mnuDisableProxy2.Enabled = True
        
        'Change systemtray icon to 'Enabled' icon.
        Call ShellTrayIconModify(Me.hWnd, imgEnabled.Picture, "Proxy - " & sFullAddress)
        
        'Show BalloonTip.
        Call ShellTrayBalloonTipShow(Me.hWnd, 1, "Proxy has been set", "Proxy - " & sFullAddress)
        sToolTip = "Proxy - " & sFullAddress
    Else
        'Show BalloonTip.
        Call ShellTrayBalloonTipShow(Me.hWnd, 3, "Error", "Proxy could not be set!")
    End If
End Sub

'Remove selected proxies
Private Sub mnuRemoveProxy_Click()
On Error Resume Next
Dim i As Long
    With lstView
        For i = .ListItems.Count To 1 Step -1 'Loop through the listview backwards.
            If .ListItems(i).Selected = True Then 'If the item is selected....
                .ListItems.Remove i '... then remove it from the listview.
            End If
        Next
    End With
    If lstView.ListItems.Count < 1 Then mnuSetRandom.Enabled = False
    StatusBar.Panels(2).Text = "Proxies: " & lstView.ListItems.Count
End Sub

'Remove all proxies.
Private Sub mnuRemoveAll_Click()
    If MsgBox("All proxies will be removed from the list. " & vbCrLf & vbCrLf & _
    "Are you sure?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
        lstView.ListItems.Clear
        StatusBar.Panels(2).Text = "Proxies: " & lstView.ListItems.Count
        mnuSetRandom.Enabled = False
    End If
End Sub

'=================================
' Systemtray menu
'=================================

'Bring app back to the desktop, so user can select a proxy.
Private Sub mnuSetProxy_Click()
    If mnuSetProxy.Checked Then
        WindowState = 1
        Me.Hide
    Else
        WindowState = 0
        Me.Show
    End If
End Sub

'Set a random proxy from the list
Private Sub mnuSetRandom_Click()
Dim i As Long
Dim sFullAddress As String
    
    i = lstView.ListItems.Count - 1
    
    Randomize
    i = CInt(Rnd() * i + 1)
    
    sFullAddress = lstView.ListItems.Item(i) & ":" & lstView.ListItems.Item(i).SubItems(1) 'Get IP:Port from Listview.
    
    If SetConnectionOptions("", sFullAddress) Then
        'Update statusbar.
        StatusBar.Panels(1).Text = "Current Proxy: " & sFullAddress
        
        'Enable the menu options.
        mnuDisableProxy1.Enabled = True
        mnuDisableProxy2.Enabled = True
        
        'Change systemtray icon to 'Enabled' icon.
        Call ShellTrayIconModify(Me.hWnd, imgEnabled.Picture, "Proxy - " & sFullAddress)
        
        'Show BalloonTip.
        Call ShellTrayBalloonTipShow(Me.hWnd, 1, "Proxy has been set", "Proxy - " & sFullAddress)
        sToolTip = "Proxy - " & sFullAddress
    Else
        'Show BalloonTip.
        Call ShellTrayBalloonTipShow(Me.hWnd, 3, "Error", "Proxy could not be set!")
    End If
End Sub

'Disable the proxy.
Private Sub mnuDisableProxy2_Click()
    
    If DisableConnectionProxy("") Then
        'Update statusbar.
        StatusBar.Panels(1).Text = "Current Proxy: Disabled"
        
        'Disable the menu options.
        mnuDisableProxy1.Enabled = False
        mnuDisableProxy2.Enabled = False
        
        'Change systemtray icon to 'Disabled' icon.
        Call ShellTrayIconModify(Me.hWnd, imgDisabled.Picture, "Proxy Changer - Disabled")
        
        'Show BalloonTip.
        Call ShellTrayBalloonTipShow(Me.hWnd, 1, "Proxy Disabled", "Proxy - Disabled")
        sToolTip = "Proxy - Disabled"
    Else
        'Show BalloonTip.
        Call ShellTrayBalloonTipShow(Me.hWnd, 3, "Error", "Proxy could not be disabled!")
    End If
    
End Sub

Private Sub mnuExit1_Click()
  Unload Me
End Sub




