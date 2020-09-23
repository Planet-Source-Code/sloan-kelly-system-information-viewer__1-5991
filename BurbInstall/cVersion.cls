VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim boolIsWinNT As Boolean          ' Is this machine running WinNT
Dim strVersionString As String      ' Version of the OS
Dim strRevision As String           ' Revision number
Dim strRegisteredOrg As String      ' Registered Company
Dim strRegisteredUser As String     ' Registered User

Public Property Get Revision() As String ' get the revision number
    Revision = strRevision
End Property

Public Property Get RegisteredOrg() As String ' get the registered company's name
    RegisteredOrg = strRegisteredOrg
End Property

Public Property Get RegisteredUser() As String ' get the registered user's name
    RegisteredUser = strRegisteredUser
End Property

Public Property Get BrowserVersion() As String ' see what version of IE they are running
    Dim oReg As New cRegistry
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        If boolIsWinNT Then
            .SectionKey = "SOFTWARE\Microsoft\Internet Explorer\URL Compatibility"
        Else
            .SectionKey = "Software\Microsoft\Internet Explorer\URL Compatibility"
        End If
        .ValueKey = "Version"
        .ValueType = REG_SZ
        BrowserVersion = Trim(.Value)
        If BrowserVersion = "" Then
            BrowserVersion = "None"
        End If
    End With
    Set oReg = Nothing
End Property

Public Property Get NetscapeVersion() As String ' see what version of Netscape they are running
    Dim oReg As New cRegistry
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        If boolIsWinNT Then
            .SectionKey = "SOFTWARE\Netscape\Netscape Navigator"
        Else
            .SectionKey = "Software\Netscape\Netscape Navigator"
        End If
        .ValueKey = "CurrentVersion"
        .ValueType = REG_SZ
        NetscapeVersion = Trim(.Value)
        If NetscapeVersion = "" Then
            NetscapeVersion = "None"
        End If
    End With
End Property

Public Property Get WindowsVersion() As String ' what version of windows is it? Windows 9X or Windows NT?
    WindowsVersion = strVersionString
End Property

Public Property Get WindowsNT() As Boolean ' check and see if the machine is running NT
    WindowsNT = boolIsWinNT
End Property

Private Sub Class_Initialize()
    ' get the initial values from the registry
    Call GetWindowsVersion
    Call GetWindowsRegistrationDetails
End Sub

Private Sub GetWindowsVersion()
    Dim oReg As New cRegistry
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        .ValueKey = "CurrentVersion"
        .ValueType = REG_SZ
        If .Value <> "" Then
            boolIsWinNT = True
            strVersionString = "NT" & .Value
            .ValueKey = "CurrentBuildNumber"
            strRevision = .Value
        Else
            boolIsWinNT = False
            .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
            .ValueKey = "Version"
            strVersionString = .Value
            .ValueKey = "VersionNumber"
            strRevision = .Value
        End If
    End With
End Sub

Private Sub GetWindowsRegistrationDetails()
    Dim oReg As New cRegistry
    With oReg
        .ClassKey = HKEY_LOCAL_MACHINE
        If boolIsWinNT Then
            .SectionKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        Else
            .SectionKey = "Software\Microsoft\Windows\CurrentVersion"
        End If
        .ValueKey = "RegisteredOrganization"
        strRegisteredOrg = .Value
        .ValueKey = "RegisteredOwner"
        strRegisteredUser = .Value
    End With
End Sub
