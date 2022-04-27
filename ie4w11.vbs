'************************************************************************************
'*
'* Author:      Curtis Wong
'* File:        ie4w11.vbs
'* Created on:  27 April 2022
'* Modified on: 27 April 2022
'* Version:     1.0.0
'* Aim:         Open native Internet Explorer on Windows 11
'*
'* Copyright (C) 2022 Curtis Wong
'*
'************************************************************************************
Option Explicit
'On Error Resume Next
'********************************
' Define constants and variables
'********************************
' Const HKEY_CLASSES_ROOT = &H80000000
' Const HKEY_CURRENT_USER = &H80000001
' Const HKEY_LOCAL_MACHINE = &H80000002
' Const HKEY_USERS = &H80000003
' Const HKEY_PERFORMANCE_DATA = &H80000004
' Const HKEY_CURRENT_CONFIG = &H80000005
' Const HKEY_DYN_DATA = &H80000006

' Const HKEY_CLASSES_ROOT = 2147483648
Const HKEY_CURRENT_USER = 2147483649
' Const HKEY_LOCAL_MACHINE = 2147483650
' Const HKEY_USERS = 2147483651
' Const HKEY_PERFORMANCE_DATA = 2147483652
' Const HKEY_CURRENT_CONFIG = 2147483653
' Const HKEY_DYN_DATA = 2147483654

' Const REG_NONE = 0
' Const REG_SZ = 1
' Const REG_EXPAND_SZ = 2
' Const REG_BINARY = 3
' Const REG_DWORD = 4
' Const REG_DWORD_LITTLE_ENDIAN = 4
' Const REG_DWORD_BIG_ENDIAN = 5
' Const REG_LINK = 6
' Const REG_MULTI_SZ = 7
' Const REG_RESOURCE_LIST = 8

Const navOpenInNewWindow = 1
Const navNoHistory = 2
Const navNoReadFromCache = 4
Const navNoWriteToCache = 8
Const navAllowAutosearch = 16
Const navBrowserBar = 32
Const navHyperlink = 64
Const navEnforceRestricted = 128
Const navNewWindowsManaged = 256
Const navUntrustedForDownload = 512
Const navTrustedForActiveX = 1024
Const navOpenInNewTab = 2048
Const navOpenInBackgroundTab = 4096
Const navKeepWordWheelText = 8192
Const navVirtualTab = 16384
Const navBlockRedirectsXDomain = 32768
Const navOpenNewForegroundTab = 65536

Const WAIT_TIMEOUT = 100
Const ERR_TIMEOUT = 100
Const READYSTATE_COMPLETE = 4

Dim url 'as String
Dim i 'as Integer
Dim objArg 'as Object
'Dim e 'as Object
'Dim strArg 'as String
Dim objProcess 'as Object
Dim Count1 'as Integer
Dim Count2 'as Integer

'*************************
' Define a registry class
'*************************
Class RegistryClass
    Private strKey, strValue

    Public Property Let KeyPath(ByVal k)
        strKey = k
    End Property

    Public Property Let ValueName(ByVal v)
        strValue = v
    End Property
    
    Public Property Get KeyPath
        KeyPath = strKey
    End Property

    Public Property Get ValueName
        ValueName = strValue
    End Property
End Class

'***********************************
' Store the arguments in a variable
'***********************************
Set objArg = WScript.Arguments

' For Each e In objArg
'     strArg = strArg & " " & CStr(e)
' Next

'************************************************************
' Get the first argument as the Internet Explorer start page
'************************************************************
If objArg.Count > 0 Then
    url = objArg.Item(0)
Else
    'Create new array list object
    Dim arrReg
    Set arrReg = CreateObject("System.Collections.ArrayList")

    'Add registry items to array
    arrReg.Add RegPath("SOFTWARE\Policies\Microsoft\Internet Explorer\Main", "Start Page")
    arrReg.Add RegPath("SOFTWARE\Microsoft\Internet Explorer\Main", "Start Page")
    arrReg.Add RegPath("SOFTWARE\Policies\Microsoft\Edge", "HomepageLocation")
    arrReg.Add RegPath("SOFTWARE\Policies\Microsoft\Edge\Recommended", "HomepageLocation")
    arrReg.Add RegPath("SOFTWARE\Policies\Microsoft\MicrosoftEdge\Internet Settings", "HomeButtonURL")

	' *********************************
	' Get Internet Explorer Start Page
	' *********************************
	Dim objReg, RegTree, ErrorCode
    RegTree = HKEY_CURRENT_USER
	Set objReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
	
    'Note: CreateObject("System.Collections.ArrayList") does not support LBound and UBound
    For i = 0 To arrReg.Count - 1
        ErrorCode = objReg.GetStringValue(RegTree, arrReg(i).KeyPath, arrReg(i).ValueName, url)
        If ErrorCode = 0 Then Exit For
    Next
    
	'If url is null, set url to about:blank
	If ErrorCode <> 0 Then url = "about:blank"
    
    Set objReg = Nothing
    Set arrReg = Nothing
End If
Set objArg = Nothing

'***********************************
' Kill all hidden Internet Explorer
'***********************************
For Each objProcess In CreateObject("Shell.Application").Windows
    If Instr(1, LCase(objProcess.FullName), "iexplore.exe") > 0 Then
        'If it is a IE process, the hidden IE can be killed by its built-in "Quit" property
        If objProcess.Visible = False Then
            objProcess.Quit
            WScript.Sleep WAIT_TIMEOUT
        End If
    End If
Next

'************************************************************
' Kill ielowutil.exe if no Internet Explorer window is found
'************************************************************
Count1 = GetObject("winmgmts:\\.\root\cimv2").ExecQuery _
            ("Select * from Win32_Process where Name = 'ielowutil.exe' or Name = 'iexplore.exe'").Count
Count2 = GetObject("winmgmts:\\.\root\cimv2").ExecQuery _
            ("Select * from Win32_Process where Name = 'ielowutil.exe'").Count

If Count1 = Count2 Then
    KillProcess "ielowutil.exe"
    WScript.Sleep WAIT_TIMEOUT
End If

'****************************************
' Start a new Internet Explorer instance
'****************************************
With CreateObject("InternetExplorer.Application")
    .Visible = True      'Show Internet Explorer
    .MenuBar = True      'Display menu bar
    .ToolBar = True      'Display tool bar
    .StatusBar = True    'Display status bar
    .AddressBar = True   'Display address bar
'    .FullScreen = False  'Full screen
    .Resizable = True    'Allow resize
'    .Width = 800
'    .Height = 600
'    .Left = 0
'    .Top = 0

    'Bring Internet Explorer Application window to the front
    CreateObject("WScript.Shell").AppActivate "Internet Explorer"

    'Open URL
    .Navigate2 url, _
               navBrowserBar + navNoReadFromCache + navTrustedForActiveX, _
               "_parent"

    'Wait for page to connect
    ' i = 0: Err.Clear
    ' Do While .Busy Or .ReadyState <> READYSTATE_COMPLETE 'Equivalent to .ReadyState <> 4
        ' WScript.Sleep WAIT_TIMEOUT 'Wait 1 second, then check again.
        ' i = i + 1
        ' If i > ERR_TIMEOUT Or Err.number <> 0 Then WScript.Quit
    ' Loop
    
    'Wait for document to load
    ' i = 0: Err.Clear
    ' Do Until LCase(.Document.ReadyState) = LCase("complete")
        ' WScript.Sleep WAIT_TIMEOUT 'Wait 1 second, then check again.
        ' i = i + 1
        ' If i > ERR_TIMEOUT Or Err.number <> 0 Then WScript.Quit
    ' Loop
End With

'*******************
' Private functions
'*******************
Private Function RegPath(ByVal k, ByVal v)
    Set RegPath = New RegistryClass
    RegPath.KeyPath = k
    RegPath.ValueName = v
End Function

Private Sub KillProcess(ByVal myProcess)
    Dim Proc
    For Each Proc In GetObject("winmgmts:\\.\root\cimv2").ExecQuery _
                        ("Select * from Win32_Process where Name = '" & myProcess & "'")
        'Kill the process
        Proc.Terminate()
    Next
End Sub
