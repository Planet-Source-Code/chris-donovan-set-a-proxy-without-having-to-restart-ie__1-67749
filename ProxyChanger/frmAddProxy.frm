VERSION 5.00
Begin VB.Form frmAddProxy 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Proxy"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddProxy 
      Caption         =   "Ok"
      Height          =   325
      Left            =   1440
      TabIndex        =   4
      Top             =   1260
      Width           =   1095
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   780
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   400
      Width           =   975
   End
End
Attribute VB_Name = "frmAddProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Add a proxy to the Listview.
Private Sub cmdAddProxy_Click()
Dim lvItem As ListItem
    
    If ValidProxy(txtIP.Text & ":" & txtPort.Text) Then
        Set lvItem = frmMain.lstView.ListItems.Add(, , txtIP.Text) 'Add IP address to the first column.
        lvItem.ListSubItems.Add , , txtPort.Text 'Add port number to the second column.
        frmMain.StatusBar.Panels(2).Text = "Proxies: " & frmMain.lstView.ListItems.Count 'Update statusbar
        frmMain.mnuSetRandom.Enabled = True
        Unload Me
    Else
        MsgBox "Invalid Proxy! ", vbExclamation, "Invalid"
    End If
    
End Sub
