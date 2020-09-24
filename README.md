<div align="center">

## NetUseDrive


</div>

### Description

maps/connects to a network drive in the same fashion as 'NET USE'
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Qaid](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-qaid.md)
**Level**          |Unknown
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\)
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-qaid-netusedrive__1-196/archive/master.zip)

### API Declarations

```
'Define structures
Public Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type
'Declare functions from MPR.DLL
Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
'Define constants
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_UNKNOWN = &HFFFF
Public Const CONNECT_UPDATE_PROFILE = &H1
Public Const NO_ERROR = 0
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_ALREADY_ASSIGNED = 85&
Public Const ERROR_BAD_DEV_TYPE = 66&
Public Const ERROR_BAD_DEVICE = 1200&
Public Const ERROR_BAD_NET_NAME = 67&
Public Const ERROR_BAD_PROFILE = 1206&
Public Const ERROR_BAD_PROVIDER = 1204&
Public Const ERROR_BUSY = 170&
Public Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Public Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Public Const ERROR_DEVICE_IN_USE = 2404&
Public Const ERROR_EXTENDED_ERROR = 1208&
Public Const ERROR_INVALID_PASSWORD = 86&
Public Const ERROR_NO_NET_OR_BAD_PATH = 1203&
Public Const ERROR_NO_NETWORK = 1222&
Public Const ERROR_NOT_CONNECTED = 2250&
Public Const ERROR_OPEN_FILES = 2401&
'Define miscellaneous variables
Private varTemp As Variant
Private sNull As String
```


### Source Code

```
Public Function NetUse(sLocalDevice As String, sShareName As String, Optional sUserID As Variant, Optional sPassword As Variant, Optional varPersistent As Variant) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                                                                 ''
'' The function, NetUseDrive, maps a network drive in the same fashion as 'NET USE'                         ''
''                                                                 ''
'' The function accepts the following parameters:                                          ''
''   sLocalDevice - a (case insensitive) string containing the local device to redirect (ie. "F:" or "LPT1"). If sLocalDevice  ''
''     is empty or is undefined/NULL, a connection to sShareName is made without redirecting a local device (ie. pipe/IPC$).  ''
''   sShareName - the UNC Name for the share to connect to. Must be in the format of "\\server\share"              ''
''   sUserID - optional, the User ID to login with (ie. "TAS01"). If it isn't passed, the User ID                ''
''     and password of the person currently logged in is used. (Actually the program is running in)              ''
''   sPassword - optional, the Password to login with. If it isn't passed, the User ID and password of              ''
''     the person currently logged in is used. (Actually the program is running in)                      ''
''   varPersistent - must be passed True (-1) or False (0) to be considered. Default is True. If False, the connection remains  ''
''     until disconnected, or until the user is logged off.                                  ''
''                                                                 ''
'' The following (long datatype) result codes are returned:                                     ''
''   NO_ERROR            (0)   Drive sLocalDevice was mapped successfully to sShareName.              ''
''   ERROR_ACCESS_DENIED       (5)   Access to the network resource was denied.                     ''
''   ERROR_ALREADY_ASSIGNED     (85)  The local device specified by sShareName is already connected to a network     ''
''                       resource.                                      ''
''   ERROR_BAD_DEV_TYPE       (66)  The type of local device and the type of network resource do not match.       ''
''   ERROR_BAD_DEVICE        (1200) The value specified by sLocalDevice is invalid.                   ''
''   ERROR_BAD_NET_NAME       (67)  The value specified by sShareName is not acceptable to any network resource     ''
''                       provider. The resource name is invalid, or the named resource cannot be located.  ''
''   ERROR_BAD_PROFILE        (1206) The user profile is in an incorrect format.                     ''
''   ERROR_BAD_PROVIDER       (1204) The default network provider is invalid.                      ''
''   ERROR_BUSY           (170)  The router or provider is busy, possibly initializing. The caller should retry.   ''
''   ERROR_CANNOT_OPEN_PROFILE    (1205) The system is unable to open the user profile to process persistent connections.  ''
''   ERROR_DEVICE_ALREADY_REMEMBERED (1202) An entry for the device specified in sShareName is already in the user profile.   ''
''   ERROR_EXTENDED_ERROR      (1208) An unknown network-specific error occured.                     ''
''   ERROR_INVALID_PASSWORD     (86)  The password sPassword is invalid.                         ''
''   ERROR_NO_NET_OR_BAD_PATH    (1203) A network component has not started, or the specified name could not be handled.  ''
''   ERROR_NO_NETWORK        (1222) There is no network present.                            ''
''                                                                 ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim netAddCxn As NETRESOURCE
Dim lCxnType As Long
Dim rc As Long
On Error GoTo ErrorHandler
'Identify the type of connection to make. If unidentified, then return ERROR_BAD_DEVICE and exit the subroutine.
If (sLocalDevice Like "[D-Z]:") Then lCxnType = RESOURCETYPE_DISK                'Network drive
If (sLocalDevice Like "LPT[1-3]") Then lCxnType = RESOURCETYPE_PRINT              'Network printer
If ((sLocalDevice = "") And (sShareName Like "\\*\IPC$")) Then lCxnType = RESOURCETYPE_ANY   'Pipe
If ((Not sLocalDevice Like "[D-Z]:") And (Not sLocalDevice Like "LPT[1-3]") And ((Not sShareName Like "\\*\IPC$") And (Not sLocalDevice = ""))) Or (Not sShareName Like "\\*\*") Then
  NetUse = ERROR_BAD_DEVICE
  GoTo EndOfFunction
End If
'Handle varPersistent
If IsMissing(varPersistent) Then
  varPersistent = CONNECT_UPDATE_PROFILE
Else
  If varPersistent = False Then
    varPersistent = 0&
  Else
    varPersistent = CONNECT_UPDATE_PROFILE
  End If
End If
'Fill in the required members of netAddCxn
With netAddCxn
  .dwType = RESOURCETYPE_DISK
  .lpLocalName = sLocalDevice
  .lpRemoteName = sShareName
  .lpProvider = Chr(0)
End With
'Perform the Net Use statement
If IsMissing(sUserID) Or IsMissing(sPassword) Then
  rc = WNetAddConnection2(netAddCxn, sNull, sNull, varPersistent)
Else
  rc = WNetAddConnection2(netAddCxn, sPassword, sUserID, varPersistent)
End If
'Process and return the result
NetUse = rc
'Handle Errors
GoTo EndOfFunction
ErrorHandler:
varTemp = MsgBox("Error #" & Err.Number & Chr(10) & Err.Description, vbCritical)
EndOfFunction:
End Function
```

