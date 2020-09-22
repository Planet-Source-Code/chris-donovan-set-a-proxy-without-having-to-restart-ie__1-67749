Attribute VB_Name = "modProxy"
Option Explicit

'======================================================================================================
'
' Description : This module is used to set or disable a proxy without having to restart IE.
'
'               It has been converted from the C++ code on the page below:
'               http://www.codeproject.com/internet/changeproxy1.asp
'
'======================================================================================================
'
' Usage :
'
' conn_name       : active connection name. (LAN = "")
' proxy_full_addr : eg "210.78.22.87:8000"
'
'
' Set proxy       : SetConnectionOptions(conn_name, proxy_full_addr)
' Disable proxy   : DisableConnectionProxy(conn_name)
'
'                   Returns 'True' on success and 'False' on failure.
'
'======================================================================================================
'
' Author : Unknown
'
'          I modified this code slightly, because the original code had a flaw in releasing the
'          memory. This would crash the entire app the second time the proxy was set or disabled.
'
'======================================================================================================

Private Type INTERNET_PER_CONN_OPTION
    dwOption As Long
    dwValue1 As Long
    dwValue2 As Long
End Type
Private Type INTERNET_PER_CONN_OPTION_LIST
    dwSize As Long
    pszConnection As Long
    dwOptionCount As Long
    dwOptionError As Long
    pOptions As Long
End Type
Private Const INTERNET_PER_CONN_FLAGS As Long = 1
Private Const INTERNET_PER_CONN_PROXY_SERVER As Long = 2
Private Const INTERNET_PER_CONN_PROXY_BYPASS As Long = 3
Private Const PROXY_TYPE_DIRECT As Long = &H1
Private Const PROXY_TYPE_PROXY As Long = &H2
Private Const INTERNET_OPTION_REFRESH As Long = 37
Private Const INTERNET_OPTION_SETTINGS_CHANGED As Long = 39
Private Const INTERNET_OPTION_PER_CONNECTION_OPTION As Long = 75
Private Declare Function InternetSetOption _
        Lib "wininet.dll" Alias "InternetSetOptionA" ( _
        ByVal hInternet As Long, ByVal dwOption As Long, _
        lpBuffer As Any, ByVal dwBufferLength As Long) As Long

' Set Proxy

Public Function SetConnectionOptions(ByVal conn_name As String, ByVal proxy_full_addr As String) As Boolean
' conn_name: active connection name. (LAN = "")
' proxy_full_addr : eg "193.28.73.241:8080"
Dim list As INTERNET_PER_CONN_OPTION_LIST
Dim bReturn As Boolean
Dim dwBufSize As Long
Dim options(0 To 2) As INTERNET_PER_CONN_OPTION
Dim abConnName() As Byte
Dim abProxyServer() As Byte
Dim abProxyBypass() As Byte
    
    dwBufSize = Len(list)
    
    ' Fill out list struct.
    list.dwSize = Len(list)
    
    ' NULL == LAN, otherwise connection name.
    abConnName() = StrConv(conn_name & vbNullChar, vbFromUnicode)
    list.pszConnection = VarPtr(abConnName(0))
    
    ' Set three options.
    list.dwOptionCount = 3

    ' Set flags.
    options(0).dwOption = INTERNET_PER_CONN_FLAGS
    options(0).dwValue1 = PROXY_TYPE_DIRECT Or PROXY_TYPE_PROXY

    ' Set proxy name.
    options(1).dwOption = INTERNET_PER_CONN_PROXY_SERVER
    abProxyServer() = StrConv(proxy_full_addr & vbNullChar, vbFromUnicode)
    options(1).dwValue1 = VarPtr(abProxyServer(0))  '//"http://proxy:80"

    ' Set proxy override.
    options(2).dwOption = INTERNET_PER_CONN_PROXY_BYPASS
    abProxyBypass() = StrConv("local" & vbNullChar, vbFromUnicode)
    options(2).dwValue1 = VarPtr(abProxyBypass(0))

    list.pOptions = VarPtr(options(0))
    ' Make sure the memory was allocated.
    If (0& = list.pOptions) Then
        ' Return FALSE if the memory wasn't allocated.
        Debug.Print "Failed to allocate memory in SetConnectionOptions()"
        SetConnectionOptions = 0
    End If

    ' Set the options on the connection.
    bReturn = InternetSetOption(0, INTERNET_OPTION_PER_CONNECTION_OPTION, list, dwBufSize)

    ' Free the allocated memory.
    Erase options
    Erase abConnName
    Erase abProxyServer
    Erase abProxyBypass
    dwBufSize = 0
    list.dwOptionCount = 0
    list.dwSize = 0
    list.pOptions = 0
    list.pszConnection = 0
    Call InternetSetOption(0, INTERNET_OPTION_SETTINGS_CHANGED, ByVal 0&, 0)
    Call InternetSetOption(0, INTERNET_OPTION_REFRESH, ByVal 0&, 0)
    SetConnectionOptions = bReturn
End Function


' Disable Proxy

Public Function DisableConnectionProxy(ByVal conn_name As String) As Boolean
' conn_name: active connection name. (LAN = "")
Dim list As INTERNET_PER_CONN_OPTION_LIST
Dim bReturn As Boolean
Dim dwBufSize As Long
Dim options(0) As INTERNET_PER_CONN_OPTION
Dim abConnName() As Byte
    
    dwBufSize = Len(list)
    
    ' Fill out list struct.
    list.dwSize = Len(list)
    
    ' NULL == LAN, otherwise connectoid name.
    abConnName() = StrConv(conn_name & vbNullChar, vbFromUnicode)
    list.pszConnection = VarPtr(abConnName(0))
    
    ' Set three options.
    list.dwOptionCount = 1

    ' Set flags.
    options(0).dwOption = INTERNET_PER_CONN_FLAGS
    options(0).dwValue1 = PROXY_TYPE_DIRECT

    list.pOptions = VarPtr(options(0))
    ' Make sure the memory was allocated.
    If (0 = list.pOptions) Then
        ' Return FALSE if the memory wasn't allocated.
        Debug.Print "Failed to allocate memory in DisableConnectionProxy()"
        DisableConnectionProxy = 0
    End If

    ' Set the options on the connection.
    bReturn = InternetSetOption(0, INTERNET_OPTION_PER_CONNECTION_OPTION, list, dwBufSize)
    
    ' Free the allocated memory.
    Erase options
    Erase abConnName
    dwBufSize = 0
    list.dwOptionCount = 0
    list.dwSize = 0
    list.pOptions = 0
    list.pszConnection = 0
    Call InternetSetOption(0, INTERNET_OPTION_SETTINGS_CHANGED, ByVal 0&, 0)
    Call InternetSetOption(0, INTERNET_OPTION_REFRESH, ByVal 0&, 0)
    DisableConnectionProxy = bReturn
End Function



