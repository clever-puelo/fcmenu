VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'
'
' Win32 NetAPIs.
'
' Win32 APIs to determine OS information.
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                                (lpVersionInformation As OSVERSIONINFO) As Long

' Win32 APIs to determine OS information.
  
Private Declare Function NetGetDCNameA Lib "Radmin32.dll" _
                        (ServerName As Byte, _
                        DomainName As Byte, _
                        lpBuffer As Long) As Long
Private Declare Function NetGetDCName Lib "Netapi32.dll" _
                        (ServerName As Any, _
                        DomainName As Any, _
                        lpBuffer As Long) As Long
                                             
Private Declare Function NetWkstaGetInfoA Lib "Radmin32.dll" _
                                (lpServer As Any, _
                                ByVal Level As Long, _
                                lpBuffer As Any) As Long
Private Declare Function NetWkstaGetInfoW Lib "Netapi32.dll" Alias "NetWkstaGetInfo" _
                (lpServer As Any, _
                ByVal Level As Long, _
                lpBuffer As Any) As Long

' Win32 APIs to determine OS information.
Private Declare Function NetUserGetLocalGroupsA Lib "rlocal32" Alias "LocalNetUseGetInfo" _
                (ServerName As Byte, _
                UserName As Byte, _
                ByVal Level As Long, _
                ByVal Flags As Long, _
                lpBuffer As Long, _
                ByVal MaxLen As Long, _
                lpEntriesRead As Long, _
                lpTotalEntries As Long) As Long
Private Declare Function NetUserGetLocalGroupsW Lib "NetApi32" Alias "LocalNetUseGetInfo" _
                (ServerName As Byte, _
                UserName As Byte, _
                ByVal Level As Long, _
                ByVal Flags As Long, _
                lpBuffer As Long, _
                ByVal MaxLen As Long, _
                lpEntriesRead As Long, _
                lpTotalEntries As Long) As Long

' Win32 APIs to determine OS information.
Private Declare Function NetWkstaUserGetInfoA Lib "Radmin32.dll" _
                                (ByVal reserved As Any, _
                                ByVal Level As Long, lpBuffer As Any) As Long
Private Declare Function NetWkstaUserGetInfo Lib "Netapi32.dll" _
                                (ByVal reserved As Any, _
                                ByVal Level As Long, lpBuffer As Any) As Long
                                
' Win32 APIs to determine OS information.
Private Declare Function NetApiBufferFree Lib "Radmin32.dll" _
                                (ByVal lpBuffer As Long) As Long
Private Declare Function NetApiBufferFreeW Lib "Netapi32.dll" Alias "NetApiBufferFree" _
                                (ByVal lpBuffer As Long) As Long
                
' Win32 APIs to determine OS information.
Private Declare Function NetUserGetInfoA Lib "radmin32" _
                (lpServer As Any, _
                UserName As Byte, _
                ByVal Level As Long, _
                lpBuffer As Long) As Long
Private Declare Function NetUserGetInfo Lib "NetApi32" _
                (lpServer As Any, _
                UserName As Byte, _
                ByVal Level As Long, _
                lpBuffer As Long) As Long
                
' Win32 APIs to determine OS information.
Private Declare Function NetUserGetGroupsA Lib "radmin32" _
                (lpServer As Any, _
                UserName As Byte, _
                ByVal Level As Long, _
                lpBuffer As Long, _
                ByVal PrefMaxLen As Long, _
                lpEntriesRead As Long, _
                lpTotalEntries As Long) As Long
Private Declare Function NetUserGetGroups Lib "NetApi32" _
                (lpServer As Any, _
                UserName As Byte, _
                ByVal Level As Long, _
                lpBuffer As Long, _
                ByVal PrefMaxLen As Long, _
                lpEntriesRead As Long, _
                lpTotalEntries As Long) As Long
                
' Win32 APIs to determine OS information.
Private Declare Function NetRemoteTODA Lib "radmin32" _
                (yServer As Any, _
                pBuffer As Long) As Long
 Private Declare Function NetRemoteTOD Lib "NetApi32" _
                (yServer As Any, _
                pBuffer As Long) As Long
               
' Win32 APIs to determine OS information.
Private Declare Function GetUserNameA Lib "advapi32.dll" _
                (ByVal lpBuffer As String, _
                nSize As Long) As Long
Private Declare Function GetUserNameW Lib "advapi32.dll" _
                (lpBuffer As Byte, _
                nSize As Long) As Long
                
' Win32 APIs to determine OS information.
Private Declare Function GetComputerNameA Lib "kernel32" _
                (ByVal lpBuffer As String, _
                nSize As Long) As Long
Private Declare Function GetComputerNameW Lib "kernel32" _
                (lpBuffer As Any, _
                nSize As Long) As Long
                
' Win32 APIs to determine OS information.
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
                (pTo As Any, _
                uFrom As Any, _
                ByVal lSize As Long)
                                
' Win32 APIs to determine OS information.
Private Declare Function lstrlenA Lib "kernel32" _
                                (ByVal LpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" _
                                (ByVal LpString As Long) As Long
                               
                                
Private Declare Function lstrcpyA Lib "kernel32" _
                                (ByVal LpString1 As String, _
                                ByVal lpString2 As Any) As Long

' Win32 APIs to determine OS information.

Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Private Declare Function StrToPtr Lib "kernel32" Alias "lstrcpyW" (ByVal Ptr As Long, Source As Byte) As Long
Private Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long

Private Declare Function PtrToStr9x Lib "kernel32" Alias "lstrcpyA" (RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function StrToPtr9x Lib "kernel32" Alias "lstrcpyA" (ByVal Ptr As Any, Source As String) As Long
Private Declare Function PtrToInt9x Lib "kernel32" Alias "lstrcpynA" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long

Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long

'
'
' Estructuras
'
'
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Private Type USER_INFO_3_API
   ' Level 0 starts here
   Name As Long
   ' Level 1 starts here
   Password As Long
   PasswordAge As Long
   Privilege As Long
   HomeDir As Long
   Comment As Long
   Flags As Long
   ScriptPath As Long
   ' Level 2 starts here
   AuthFlags As Long
   FullName As Long
   UserComment As Long
   Parms As Long
   Workstations As Long
   LastLogon As Long
   LastLogoff As Long
   AcctExpires As Long
   MaxStorage As Long
   UnitsPerWeek As Long
   LogonHours As Long
   BadPwCount As Long
   NumLogons As Long
   LogonServer As Long
   CountryCode As Long
   CodePage As Long
   ' Level 3 starts here
   UserID As Long
   PrimaryGroupID As Long
   Profile As Long
   HomeDirDrive As Long
   PasswordExpired As Long
End Type

Private Type USER_INFO_3
   ' Level 0 starts here
   Name As String
   ' Level 1 starts here
   Password As String
   PasswordAge As Long
   Privilege As Long
   HomeDir As String
   Comment As String
   Flags As Long
   ScriptPath As String
   ' Level 2 starts here
   AuthFlags As Long
   FullName As String
   UserComment As String
   Parms As String
   Workstations As String
   LastLogon As Long
   LastLogoff As Long
   AcctExpires As Long
   MaxStorage As Long
   UnitsPerWeek As Long
   LogonHours(0 To 20) As Byte
   BadPwCount As Long
   NumLogons As Long
   LogonServer As String
   CountryCode As Long
   CodePage As Long
   ' Level 3 starts here
   UserID As Long
   PrimaryGroupID As Long
   Profile As String
   HomeDirDrive As String
   PasswordExpired As Boolean
End Type

Private Type GROUP_INFO_2_API
   Name As Long
   Comment As Long
   GroupID As Long
   Attributes As Long
End Type

Private Type GROUP_INFO_2
   Name As String
   Comment As String
   GroupID As Long
   Attributes As Long
End Type

Private Type WKSTA_INFO_102
   wki102_platform_id As Long
   wki102_computername As Long
   wki102_langroup As Long
   wki102_ver_major As Long
   wki102_ver_minor As Long
   wki102_lanroot As Long
   wki102_logged_on_users As Long
End Type

Private Type WkstaInfo102
   PlatformId As Long
   ComputerName As String
   LanGroup As String
   VerMajor As Long
   VerMinor As Long
   LanRoot As String
   LoggedOnUsers As Long
End Type

Private Type WKSTA_USER_INFO_1
   wkui1_username As Long
   wkui1_logon_domain As Long
   wkui1_oth_domains As Long
   wkui1_logon_server As Long
End Type

Private Type WkstaUserInfo1
   UserName As String
   LogonDomain As String
   OtherDomains As String
   LogonServer As String
End Type

'
'
'    Constantes
'

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const FILTER_TEMP_DUPLICATE_ACCOUNT As Long = &H1&
Private Const FILTER_NORMAL_ACCOUNT As Long = &H2&
Private Const FILTER_PROXY_ACCOUNT As Long = &H4&
Private Const FILTER_INTERDOMAIN_TRUST_ACCOUNT As Long = &H8&
Private Const FILTER_WORKSTATION_TRUST_ACCOUNT As Long = &H10&
Private Const FILTER_SERVER_TRUST_ACCOUNT As Long = &H20&

Private Const NERR_Success As Long = 0&
Private Const ERROR_MORE_DATA As Long = 234&

Private Const TIMEQ_FOREVER = -1&             '((unsigned long) -1L)
Private Const USER_MAXSTORAGE_UNLIMITED = -1& '((unsigned long) -1L)
Private Const USER_NO_LOGOFF = -1&            '((unsigned long) -1L)
Private Const UNITS_PER_DAY = 24
Private Const UNITS_PER_WEEK = UNITS_PER_DAY * 7

Private Const USER_PRIV_MASK = 3
Private Const USER_PRIV_GUEST = 0
Private Const USER_PRIV_USER = 1
Private Const USER_PRIV_ADMIN = 2

Private Const UNLEN = 256         ' Maximum username length
Private Const GNLEN = UNLEN       ' Maximum groupname length
Private Const CNLEN = 15          ' Maximum computer name length
Private Const MAXCOMMENTSZ = 256  ' Multipurpose comment length
Private Const LG_INCLUDE_INDIRECT As Long = &H1&

' Private Const NERR_Success As Long = 0&
'
' Member variables
'
Private m_Wks As WkstaInfo102
Private m_User As WkstaUserInfo1

Private m_UserInfo As USER_INFO_3
Private m_OS As OSVERSIONINFO
Private m_UserName As String
Private m_Server As String
Private m_Groups() As String
Private m_LocalGroups() As String
Private m_OSVers As Byte

' *********************************************************
'  Initialization
' *********************************************************
Private Sub Class_Initialize()
   Dim os As OSVERSIONINFO
   os.dwOSVersionInfoSize = Len(os)
   
   Call GetVersionEx(os)
   
      m_OSVers = os.dwPlatformId
      
      Server = CurrentMachineName()
      UserName = CurrentUserName()
      
        RefreshWkstn
        Refresh

      
  ' End If
End Sub

' *********************************************************
'  Public Properties
' *********************************************************
Public Property Get OSPlataf() As String
     
  Select Case m_OSVers
    Case Is = 1: OSPlataf = "Windows 95/98"
    Case Is = 2: OSPlataf = "Windows NT"
    Case Else: OSPlataf = "Desconocido"
  End Select

End Property

Public Property Get UserName() As String
   UserName = m_UserInfo.Name
End Property


Public Property Let UserName(NewVal As String)
   m_UserName = NewVal
End Property

Public Property Get Server() As String
   Server = m_Server
End Property

Public Property Let Server(NewVal As String)
   m_Server = NewVal
End Property

Public Property Get GroupCount() As Long
  GroupCount = UBound(m_Groups) + 1
End Property

Public Property Get Group(ByVal Index As Long) As String
   If Index >= LBound(m_Groups) And Index <= UBound(m_Groups) Then
      Group = m_Groups(Index)
   End If
End Property

Public Property Get LocalGroupCount() As Long
   LocalGroupCount = UBound(m_LocalGroups) + 1
End Property

Public Property Get LocalGroup(ByVal Index As Long) As String
   If Index >= LBound(m_LocalGroups) And Index <= UBound(m_LocalGroups) Then
      LocalGroup = m_LocalGroups(Index)
   End If
End Property

Public Property Get Password()
  ' Err.Raise Number:=vbObjectError + 1, _
  '           Source:="CNetUser.Password", _
   '          Description:="Password property is write-only for security."
      Password = m_UserInfo.Password

End Property

Public Property Get PasswordAge() As Long
   PasswordAge = m_UserInfo.PasswordAge
End Property

Public Property Get Privilege() As Long
   Privilege = m_UserInfo.Privilege
End Property

Public Property Get HomeDir() As String
   HomeDir = m_UserInfo.HomeDir
End Property

Public Property Get Comment() As String
   Comment = m_UserInfo.Comment
End Property

Public Property Get Flags() As Long
   Flags = m_UserInfo.Flags
End Property

Public Property Get ScriptPath() As String
   ScriptPath = m_UserInfo.ScriptPath
End Property

Public Property Get AuthFlags() As Long
   AuthFlags = m_UserInfo.AuthFlags
End Property

Public Property Get FullName() As String
   FullName = m_UserInfo.FullName
End Property

Public Property Get UserComment() As String
   UserComment = m_UserInfo.UserComment
End Property

Public Property Get Parms() As String
   Parms = m_UserInfo.Parms
End Property

Public Property Get Workstations() As String
   Workstations = m_UserInfo.Workstations
End Property

Public Property Get LastLogon() As Long
    LastLogon = m_UserInfo.LastLogon
End Property

Public Property Get LastLogoff() As Long
   LastLogoff = m_UserInfo.LastLogoff
End Property

Public Property Get LastLogonDate() As Double
    LastLogonDate = NetTimeToVbTime(m_UserInfo.LastLogon)
End Property

Public Property Get LastLogoffDate() As Double
   LastLogoffDate = NetTimeToVbTime(m_UserInfo.LastLogoff)
End Property

Public Property Get AcctExpires() As Long
   AcctExpires = m_UserInfo.AcctExpires
End Property

Public Property Get AcctExpiresDate() As Long
   AcctExpiresDate = NetTimeToVbTime(m_UserInfo.AcctExpires)
End Property

Public Property Get MaxStorage() As Long
   MaxStorage = m_UserInfo.MaxStorage
End Property

Public Property Get UnitsPerWeek() As Long
   UnitsPerWeek = m_UserInfo.UnitsPerWeek
End Property

Public Property Get LogonHours(ByVal Index As Long) As Byte
   If Index >= 0 And Index <= 20 Then
      LogonHours = m_UserInfo.LogonHours(Index)
   End If
End Property

Public Property Get BadPasswordCount() As Long
   BadPasswordCount = m_UserInfo.BadPwCount
End Property

Public Property Get NumLogons() As Long
   NumLogons = m_UserInfo.NumLogons
End Property

Public Property Get LogonServer() As String
   LogonServer = m_UserInfo.LogonServer
End Property

Public Property Get CountryCode() As Long
   CountryCode = m_UserInfo.CountryCode
End Property

Public Property Get CodePage() As Long
   CodePage = m_UserInfo.CodePage
End Property

Public Property Get UserID() As Long
   UserID = m_UserInfo.UserID
End Property

Public Property Get PrimaryGroupID() As Long
   PrimaryGroupID = m_UserInfo.PrimaryGroupID
End Property

Public Property Get Profile() As String
   Profile = m_UserInfo.Profile
End Property

Public Property Get HomeDirDrive() As String
   HomeDirDrive = m_UserInfo.HomeDirDrive
End Property

Public Property Get PasswordExpired() As Boolean
   PasswordExpired = m_UserInfo.PasswordExpired
End Property

' *********************************************************
'  Public Properties (Workstation)
' *********************************************************
Public Property Get ComputerName() As String
   ComputerName = m_Wks.ComputerName
End Property

Public Property Get Domain() As String
   Domain = m_Wks.LanGroup
End Property

Public Property Get LanRoot() As String
   LanRoot = m_Wks.LanRoot
End Property

Public Property Get LoggedOnUsers() As Long
   LoggedOnUsers = m_Wks.LoggedOnUsers
End Property

Public Property Get PlatformId() As Long
   PlatformId = m_Wks.PlatformId
End Property

Public Property Get VerMajor() As Long
   VerMajor = m_Wks.VerMajor
End Property

Public Property Get VerMinor() As Long
   VerMinor = m_Wks.VerMinor
End Property

' *********************************************************
'  Public Properties (Workstation User)
' *********************************************************
Public Property Get LogonDomain() As String
   LogonDomain = m_User.LogonDomain
End Property

Public Property Get OtherDomains() As String
   OtherDomains = m_User.OtherDomains
End Property

Public Property Get UserNameWks() As String
   UserNameWks = m_User.UserName
End Property

' *********************************************************
'  Public Methods
' *********************************************************

Public Function CurrentUserName() As String
    Dim Buffer As String
    Dim yBuffer() As Byte
    Dim nRet As Long
    Dim nLen As Long
    Const NameLength = UNLEN + 1
                
    Buffer = String(UNLEN, Chr$(0))

    If m_OSVers = VER_PLATFORM_WIN32_NT Then
            nLen = NameLength * 2
            ReDim yBuffer(0 To nLen - 1) As Byte
            If GetUserNameW(yBuffer(0), nLen) Then
                    Buffer = yBuffer
                    CurrentUserName = Left(Buffer, nLen - 1)
                    End If
   Else
            Buffer = String(UNLEN, Chr$(0))
            GetUserNameA Buffer, UNLEN
           CurrentUserName = Left$(Buffer, InStr(Buffer, Chr$(0)) - 1)
End If
   

End Function

Public Function CurrentMachineName() As String
   Dim Buffer As String
   Dim yBuffer() As Byte
   Dim nRet As Long
   Dim nLen As Long
   Const NameLength = CNLEN + 1
   
    Buffer = String(UNLEN, Chr$(0))
    If m_OSVers = VER_PLATFORM_WIN32_NT Then
            nLen = NameLength * 2
            ReDim yBuffer(0 To nLen - 1) As Byte
            If GetComputerNameW(yBuffer(0), nLen) Then
                Buffer = yBuffer
                CurrentMachineName = Left(Buffer, nLen)
                End If
    Else
        GetComputerNameA Buffer, UNLEN
        CurrentMachineName = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
    End If
    
    
End Function


Public Function Refresh() As Boolean
   Dim lpBuffer As Long
   Dim yUserName() As Byte
   Dim yServer() As Byte
   Dim uUserApi As USER_INFO_3_API
   Dim nRet As Long
   
   ReDim yUserName(0 To 7) As Byte
   ReDim yServer(0 To 12) As Byte
   

   m_UserInfo.Name = m_User.UserName
   m_Server = GetPrimaryDCName()

      If InStr(m_Server, "\\") <> 1 Then
         m_Server = "\\" & m_Server
        End If
             
    If m_OSVers = VER_PLATFORM_WIN32_NT Then
            yUserName = m_UserName & vbNullChar
            yServer = m_Server & vbNullChar
            nRet = NetUserGetInfo(yServer(0), yUserName(0), 3, lpBuffer)
    Else
            yUserName = StrConv(m_UserName & vbNullChar, vbFromUnicode)
            yServer = StrConv(m_Server & vbNullChar, vbFromUnicode)
            nRet = NetUserGetInfoA(yServer(0), yUserName(0), 3, lpBuffer)
   End If
   
   If nRet = NERR_Success Then
      CopyMem uUserApi, ByVal lpBuffer, Len(uUserApi)
      '
      ' Transfer data to VB structure
      '
      m_UserInfo.Name = PointerToString(uUserApi.Name)
      m_UserInfo.Password = PointerToString(uUserApi.Password)
      m_UserInfo.PasswordAge = uUserApi.PasswordAge
      m_UserInfo.Privilege = uUserApi.Privilege
      m_UserInfo.HomeDir = PointerToString(uUserApi.HomeDir)
      m_UserInfo.Comment = PointerToString(uUserApi.Comment)
      m_UserInfo.Flags = uUserApi.Flags
      m_UserInfo.ScriptPath = PointerToString(uUserApi.ScriptPath)
      m_UserInfo.AuthFlags = uUserApi.AuthFlags
      m_UserInfo.FullName = PointerToString(uUserApi.FullName)
      m_UserInfo.UserComment = PointerToString(uUserApi.UserComment)
      m_UserInfo.Parms = PointerToString(uUserApi.Parms)
      m_UserInfo.Workstations = PointerToString(uUserApi.Workstations)
      m_UserInfo.LastLogon = uUserApi.LastLogon
      m_UserInfo.LastLogoff = uUserApi.LastLogoff
      m_UserInfo.AcctExpires = uUserApi.AcctExpires
      m_UserInfo.MaxStorage = uUserApi.MaxStorage
      m_UserInfo.UnitsPerWeek = uUserApi.UnitsPerWeek
      
      CopyMem m_UserInfo.LogonHours(0), ByVal uUserApi.LogonHours, 21
      m_UserInfo.BadPwCount = uUserApi.BadPwCount
      m_UserInfo.NumLogons = uUserApi.NumLogons
      m_UserInfo.LogonServer = PointerToString(uUserApi.LogonServer)
      m_UserInfo.CountryCode = uUserApi.CountryCode
      m_UserInfo.CodePage = uUserApi.CodePage
      m_UserInfo.UserID = uUserApi.UserID
      m_UserInfo.PrimaryGroupID = uUserApi.PrimaryGroupID
      m_UserInfo.Profile = PointerToString(uUserApi.Profile)
      m_UserInfo.HomeDirDrive = PointerToString(uUserApi.HomeDirDrive)
      m_UserInfo.PasswordExpired = CBool(uUserApi.PasswordExpired)
      '
      ' Return success
      '
      Refresh = True
   End If
   '
   ' Clean up
   '
   If lpBuffer Then
        If m_OSVers = VER_PLATFORM_WIN32_NT Then
                Call NetApiBufferFreeW(lpBuffer)
        Else
                Call NetApiBufferFree(lpBuffer)
                End If
      RefreshGroups
   '   RefreshLocalGroups
   End If

End Function

Public Function NetTimeToVbTime(NetDate As Long) As Double
   Const BaseDate# = 25569   'DateSerial(1970, 1, 1)
   Const SecsPerDay# = 86400
   NetTimeToVbTime = BaseDate + (CDbl(NetDate) / SecsPerDay)
End Function

' *********************************************************
'  Private Methods
' *********************************************************
Private Sub RefreshLocalGroups()
   Dim lpBuffer As Long
   Dim yUserName() As Byte
   Dim yServer() As Byte
   Dim lpGroups() As Long
   Dim nRead As Long
   Dim nTotal As Long
   Dim nRet As Long
   Dim i As Long
   'Const Flags& = LG_INCLUDE_INDIRECT 'To get domain groups as well
   Const Flags& = 0
   
   yUserName = m_UserName & vbNullChar
  yUserName = StrConv((m_UserName & Chr(0)), vbFromUnicode)

      If InStr(m_Server, "\\") = 1 Then
         yServer = m_Server & vbNullChar
      Else
         yServer = "\\" & m_Server & vbNullChar
      End If
                        
    yServer = StrConv(m_Server & vbNullChar, vbFromUnicode)

           If m_OSVers = VER_PLATFORM_WIN32_NT Then
                nRet = NetUserGetLocalGroupsW(yServer(0), yUserName(0), 0, Flags, lpBuffer, &H400, nRead, nTotal)
      Else
               yServer = StrConv(m_Server & vbNullChar, vbFromUnicode)
                nRet = NetUserGetLocalGroupsA(yServer(0), yUserName(0), 0, Flags, lpBuffer, &H400, nRead, nTotal)
                End If

   If nRet = NERR_Success Then
      ReDim lpGroups(0 To nRead - 1) As Long
      ReDim m_LocalGroups(0 To nRead - 1) As String
      CopyMem lpGroups(0), ByVal lpBuffer, nRead * 4
      For i = 0 To nRead - 1
         m_LocalGroups(i) = PointerToString(lpGroups(i))
      Next i
   End If
   '
   ' Clean up
   '
   If lpBuffer Then
        If m_OSVers = VER_PLATFORM_WIN32_NT Then
                Call NetApiBufferFreeW(lpBuffer)
        Else
                Call NetApiBufferFree(lpBuffer)
                End If
   End If
End Sub

Private Sub RefreshGroups()
   Dim lpBuffer As Long
   Dim yUserName() As Byte
   Dim yServer() As Byte
   Dim lpGroups() As Long
   Dim nRead As Long
   Dim nTotal As Long
   Dim nRet As Long
   Dim i As Long
   

      If InStr(m_Server, "\\") = 1 Then
         yServer = m_Server & vbNullChar
      Else
         yServer = "\\" & m_Server & vbNullChar
      End If
                              
     If m_OSVers = VER_PLATFORM_WIN32_NT Then
                  yUserName = m_UserName & vbNullChar
                  yServer = m_Server & vbNullChar
                 nRet = NetUserGetGroups(yServer(0), yUserName(0), 0, lpBuffer, &H400, nRead, nTotal)
      Else
               yUserName = StrConv((m_UserName & Chr(0)), vbFromUnicode)
               yServer = StrConv(m_Server & vbNullChar, vbFromUnicode)
                nRet = NetUserGetGroupsA(yServer(0), yUserName(0), 0, lpBuffer, &H400, nRead, nTotal)
                End If
      
   If nRet = NERR_Success Then
      ReDim lpGroups(0 To nRead - 1) As Long
      ReDim m_Groups(0 To nRead - 1) As String
      CopyMem lpGroups(0), ByVal lpBuffer, nRead * 4
      For i = 0 To nRead - 1
         m_Groups(i) = PointerToString(lpGroups(i))
      Next i
   End If
   '
   ' Clean up
   '
   If lpBuffer Then
        If m_OSVers = VER_PLATFORM_WIN32_NT Then
                Call NetApiBufferFreeW(lpBuffer)
        Else
                Call NetApiBufferFree(lpBuffer)
                End If
   End If
End Sub

' *********************************************************
'
' *********************************************************
Private Sub RefreshWkstn()
   Dim lpBuffer As Long
   Dim nRet As Long
   Dim wki As WKSTA_INFO_102
   Dim wkui As WKSTA_USER_INFO_1
   '
   ' These functions only exist in Windows NT!!!
   '
   ' Obtain workstation information
   '
   If m_OSVers = VER_PLATFORM_WIN32_NT Then
            nRet = NetWkstaGetInfoW(ByVal "", 102, lpBuffer)
   Else
            nRet = NetWkstaGetInfoA(ByVal "", 102, lpBuffer)
            End If
            
   If nRet = NERR_Success Then
      '
      ' Transfer data to VB-friendly structure
      '
      CopyMem wki, ByVal lpBuffer, Len(wki)
      m_Wks.PlatformId = wki.wki102_platform_id
      m_Wks.ComputerName = PointerToString(wki.wki102_computername)
      m_Wks.LanGroup = PointerToString(wki.wki102_langroup)
      m_Wks.VerMajor = wki.wki102_ver_major
      m_Wks.VerMinor = wki.wki102_ver_minor
      m_Wks.LanRoot = PointerToString(wki.wki102_lanroot)
      m_Wks.LoggedOnUsers = wki.wki102_logged_on_users
      '
      ' Clean up
      '
      If lpBuffer Then
        If m_OSVers = VER_PLATFORM_WIN32_NT Then
                Call NetApiBufferFreeW(lpBuffer)
        Else
                Call NetApiBufferFree(lpBuffer)
                End If
      End If
   End If
   '
   ' Obtain user information for this workstation
   '
    If m_OSVers = VER_PLATFORM_WIN32_NT Then
        nRet = NetWkstaUserGetInfo(0&, 1&, lpBuffer)
    Else
        nRet = NetWkstaUserGetInfoA(0&, 1&, lpBuffer)
        End If
   
   
   If nRet = NERR_Success Then
      '
      ' Transfer data to VB-friendly structure
      '
      CopyMem wkui, ByVal lpBuffer, Len(wkui)
      m_User.UserName = PointerToString(wkui.wkui1_username)
      m_User.LogonDomain = PointerToString(wkui.wkui1_logon_domain)
      m_User.LogonServer = PointerToString(wkui.wkui1_logon_server)
      m_User.OtherDomains = PointerToString(wkui.wkui1_oth_domains)
      '
      ' Clean up
      '
      If lpBuffer Then
        If m_OSVers = VER_PLATFORM_WIN32_NT Then
                Call NetApiBufferFreeW(lpBuffer)
        Else
                Call NetApiBufferFree(lpBuffer)
                End If
      End If
   End If
   
End Sub

Function GetPrimaryDCName() As String

        Dim Result As Long, DCName As String, lpBuffer As Long
        Dim DCNArray(100) As Byte
        Dim yDomain() As Byte

        ReDim yServer(0 To 15) As Byte

        If m_OSVers = VER_PLATFORM_WIN32_NT Then
                yDomain = m_User.LogonDomain & vbNullChar
                Result = NetGetDCName(0&, yDomain(0), lpBuffer)
        Else
                yDomain = StrConv(m_User.LogonDomain & vbNullChar, vbFromUnicode)
                Result = NetGetDCNameA(0&, yDomain(0), lpBuffer)
        End If

        If Result <> 0 Then
                MsgBox "Unable to determine Domain Controller Name"
                Exit Function
        End If

        If m_OSVers = VER_PLATFORM_WIN32_NT Then
                DCName = PointerToStringW(lpBuffer)
                Result = NetApiBufferFreeW(lpBuffer)
        Else
                DCName = String(64, 0)
                Result = PtrToStr9x(ByVal DCName, lpBuffer)
                Result = NetApiBufferFree(lpBuffer)
        End If

        GetPrimaryDCName = DCName

End Function

Private Function PointerToString(LpString As Long) As String
        
        If m_OSVers = VER_PLATFORM_WIN32_NT Then
              PointerToString = PointerToStringW(LpString)
        Else
              PointerToString = PointerToStringA(LpString)
        End If

End Function

Private Function PointerToStringA(LpString As Long) As String
   
   Dim nLen As Long
   Dim Buffer As String * 80
   Buffer = String(80, 0)
    
   If LpString > 0 Then
      Call lstrcpyA(Buffer, LpString)
      End If
            
PointerToStringA = Left(Buffer, InStr(Buffer, Chr(0)) - 1)

End Function

Private Function PointerToStringW(lpStringW As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMem Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function

