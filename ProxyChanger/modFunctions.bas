Attribute VB_Name = "modFunctions"
Option Explicit

Private Sub Main()
    
    'If instance of the app is already running then let user know and close app.
    If App.PrevInstance Then
        MsgBox "You're already running an instance of this app. ", vbExclamation, "Multiple instances"
        End 'Unload Me doesn't work in a module, so we use End instead.
    Else
        frmMain.Show
    End If
    
End Sub

'Read proxy list.
Public Function ReadProxyList(sFileName As String, oListview As Object)
    Dim sProxy As String
    Dim arrPart() As String
    Dim lvItem As ListItem
    
    On Error Resume Next
    If Dir(sFileName, vbNormal) <> "" Then
        Open sFileName For Input As #1
            Do While Not EOF(1)
                Line Input #1, sProxy
                  If ValidProxy(sProxy) Then 'Only add the proxy if it has a valid IP and port number.
                    arrPart = Split(sProxy, ":") 'Split the IP address and port number on the ':' delimiter.
                    Set lvItem = oListview.ListItems.Add(, , arrPart(0)) 'Add IP address to the first column.
                    lvItem.ListSubItems.Add , , arrPart(1) 'Add port number to the second column.
                  End If
            Loop
        Close #1
    End If
    
    frmMain.StatusBar.Panels(2).Text = "Proxies: " & oListview.ListItems.Count
End Function

'Save the proxies in the Listview to a text file.
Public Function SaveProxyList(sFileName As String, oListview As Object)
    Dim i As Integer
      Open sFileName For Output As #1
          For i = 1 To oListview.ListItems.Count 'Loop through Listview.
              Print #1, oListview.ListItems.Item(i) & ":" & oListview.ListItems.Item(i).SubItems(1) 'Write line to text file.
          Next i
      Close #1
End Function

'Check if a full proxy address is valid - eg. "178.25.22.239:8080"
Public Function ValidProxy(sProxyAddress As String) As Boolean
    Dim arrAddress() As String, arrIP() As String

    arrAddress = Split(sProxyAddress, ":")
    arrIP = Split(arrAddress(0), ".")

    On Error GoTo ErrHandler

    If Not IsNumeric(arrAddress(1)) Then
        ValidProxy = False
        Exit Function
    End If
    If CLng(arrAddress(1)) < 1 Or CLng(arrAddress(1)) > 65535 Then
        ValidProxy = False
        Exit Function
    End If
    If Not IsNumeric(arrIP(0)) Or Not IsNumeric(arrIP(1)) Or Not IsNumeric(arrIP(2)) Or Not IsNumeric(arrIP(3)) Then
        ValidProxy = False
        Exit Function
    End If
    If CInt(arrIP(0)) > 255 Or CInt(arrIP(1)) > 255 Or CInt(arrIP(2)) > 255 Or CInt(arrIP(3)) > 255 Then
        ValidProxy = False
        Exit Function
    End If
    If CInt(arrIP(0)) < 1 Or CInt(arrIP(1)) < 1 Or CInt(arrIP(2)) < 1 Or CInt(arrIP(3)) < 1 Then
        ValidProxy = False
        Exit Function
    End If

    ValidProxy = True

Exit Function
ErrHandler:
    ValidProxy = False
End Function
