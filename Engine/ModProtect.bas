Attribute VB_Name = "ModProtect"
Option Explicit

Private Const STATUS_ACCESS_DENIED = &HC0000022
Private Const SECTION_MAP_WRITE = &H2
Private Const SECTION_MAP_READ = &H4
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const NO_INHERITANCE = 0
Private Const DACL_SECURITY_INFORMATION = &H4

Private Type UNICODE_STRING
length As Integer
MaximumLength As Integer
Buffer As Long
End Type

Private Type OBJECT_ATTRIBUTES
length As Long
RootDirectory As Long
ObjectName As Long
Attributes As Long
SecurityDeor As Long
SecurityQualityOfService As Long
End Type

Private Enum ACCESS_MODE
NOT_USED_ACCESS
GRANT_ACCESS
SET_ACCESS
DENY_ACCESS
REVOKE_ACCESS
SET_AUDIT_SUCCESS
SET_AUDIT_FAILURE
End Enum

Private Enum MULTIPLE_TRUSTEE_OPERATION
NO_MULTIPLE_TRUSTEE
TRUSTEE_IS_IMPERSONATE
End Enum

Private Enum TRUSTEE_FORM
TRUSTEE_IS_SID
TRUSTEE_IS_NAME
End Enum

Private Enum TRUSTEE_TYPE
TRUSTEE_IS_UNKNOWN
TRUSTEE_IS_USER
TRUSTEE_IS_GROUP
End Enum

Private Type TRUSTEE
pMultipleTrustee As Long
MultipleTrusteeOperation As MULTIPLE_TRUSTEE_OPERATION
TrusteeForm As TRUSTEE_FORM
TrusteeType As TRUSTEE_TYPE
ptstrName As String
End Type

Private Type EXPLICIT_ACCESS
grfAccessPermissions As Long
grfAccessMode As ACCESS_MODE
grfInheritance As Long
TRUSTEE As TRUSTEE
End Type

Private Enum SE_OBJECT_TYPE
SE_UNKNOWN_OBJECT_TYPE = 0
SE_FILE_OBJECT
SE_SERVICE
SE_PRINTER
SE_REGISTRY_KEY
SE_LMSHARE
SE_KERNEL_OBJECT
SE_WINDOW_OBJECT
SE_DS_OBJECT
SE_DS_OBJECT_ALL
SE_PROVIDER_DEFINED_OBJECT
SE_WMIGUID_OBJECT
End Enum

Private Declare Function SetSecurityInfo Lib "advapi32.dll" (ByVal Handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ppsidOwner As Long, ppsidGroup As Long, ppDacl As Any, ppSacl As Any) As Long
Private Declare Function GetSecurityInfo Lib "advapi32.dll" (ByVal Handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ppsidOwner As Long, ppsidGroup As Long, ppDacl As Any, ppSacl As Any, ppSecurityDeor As Long) As Long
Private Declare Function SetEntriesInAcl Lib "advapi32.dll" Alias "SetEntriesInAclA" (ByVal cCountOfExplicitEntries As Long, pListOfExplicitEntries As EXPLICIT_ACCESS, ByVal OldAcl As Long, NewAcl As Long) As Long
Private Declare Sub RtlInitUnicodeString Lib "NTDLL.DLL" (DestinationString As UNICODE_STRING, ByVal SourceString As Long)
Private Declare Function ZwOpenSection Lib "NTDLL.DLL" (SectionHandle As Long, ByVal DesiredAccess As Long, ObjectAttributes As Any) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type

Private verinfo As OSVERSIONINFO
Private g_pMapPhysicalMemory As Long
Private g_hMPM As Long
Private aByte(3) As Byte
Dim cPid As Long
Dim cProt As Boolean

Public Function ProtectPID(value As Boolean)
Dim rPid As Integer
If cProt <> True Then
cPid = GetCurrentProcessId
cProt = True
End If
If value = True And cProt = True Then
Randomize
rPid = Int(Rnd * 212)
ChangeCurrentProcessID (cPid + rPid)
cProt = True
Else
ChangeCurrentProcessID (cPid)
cProt = False
End If
End Function

Private Function ChangeCurrentProcessID(FalsePID As Long)
Dim thread As Long, process As Long, fw As Long, bw As Long
Dim lOffsetFlink As Long, lOffsetBlink As Long, lOffsetPID As Long
verinfo.dwOSVersionInfoSize = Len(verinfo)
If (GetVersionEx(verinfo)) <> 0 Then
If verinfo.dwPlatformId = 2 Then
If verinfo.dwMajorVersion = 5 Then
Select Case verinfo.dwMinorVersion
Case 0
lOffsetPID = &H9C
Case 1
lOffsetPID = &H84
End Select
End If
End If
End If
If OpenPhysicalMemory <> 0 Then
thread = GetData(&HFFDFF124)
process = GetData(thread + &H44)
SetData process + lOffsetPID, FalsePID
CloseHandle g_hMPM
End If
End Function

Private Sub SetPhyscialMemorySectionCanBeWrited(ByVal hSection As Long)
Dim pDacl As Long
Dim pNewDacl As Long
Dim pSD As Long
Dim dwRes As Long
Dim ea As EXPLICIT_ACCESS
GetSecurityInfo hSection, SE_KERNEL_OBJECT, DACL_SECURITY_INFORMATION, 0, 0, pDacl, 0, pSD
ea.grfAccessPermissions = SECTION_MAP_WRITE
ea.grfAccessMode = GRANT_ACCESS
ea.grfInheritance = NO_INHERITANCE
ea.TRUSTEE.TrusteeForm = TRUSTEE_IS_NAME
ea.TRUSTEE.TrusteeType = TRUSTEE_IS_USER
ea.TRUSTEE.ptstrName = "CURRENT_USER" & vbNullChar
SetEntriesInAcl 1, ea, pDacl, pNewDacl
SetSecurityInfo hSection, SE_KERNEL_OBJECT, DACL_SECURITY_INFORMATION, 0, 0, ByVal pNewDacl, 0
CleanUp:
LocalFree pSD
LocalFree pNewDacl
End Sub

Private Function OpenPhysicalMemory() As Long
Dim Status As Long
Dim PhysmemString As UNICODE_STRING
Dim Attributes As OBJECT_ATTRIBUTES
RtlInitUnicodeString PhysmemString, StrPtr("\Device\PhysicalMemory")
Attributes.length = Len(Attributes)
Attributes.RootDirectory = 0
Attributes.ObjectName = VarPtr(PhysmemString)
Attributes.Attributes = 0
Attributes.SecurityDeor = 0
Attributes.SecurityQualityOfService = 0
Status = ZwOpenSection(g_hMPM, SECTION_MAP_READ Or SECTION_MAP_WRITE, Attributes)
If Status = STATUS_ACCESS_DENIED Then
Status = ZwOpenSection(g_hMPM, READ_CONTROL Or WRITE_DAC, Attributes)
SetPhyscialMemorySectionCanBeWrited g_hMPM
CloseHandle g_hMPM
Status = ZwOpenSection(g_hMPM, SECTION_MAP_READ Or SECTION_MAP_WRITE, Attributes)
End If
Dim lDirectoty As Long
verinfo.dwOSVersionInfoSize = Len(verinfo)
If (GetVersionEx(verinfo)) <> 0 Then
If verinfo.dwPlatformId = 2 Then
If verinfo.dwMajorVersion = 5 Then
Select Case verinfo.dwMinorVersion
Case 0
lDirectoty = &H30000
Case 1
lDirectoty = &H39000
End Select
End If
End If
End If
If Status = 0 Then
g_pMapPhysicalMemory = MapViewOfFile(g_hMPM, 4, 0, lDirectoty, &H1000)
If g_pMapPhysicalMemory <> 0 Then OpenPhysicalMemory = g_hMPM
End If
End Function

Private Function LinearToPhys(BaseAddress As Long, addr As Long) As Long
Dim VAddr As Long, PGDE As Long, PTE As Long, PAddr As Long
Dim lTemp As Long
VAddr = addr
CopyMemory aByte(0), VAddr, 4
lTemp = Fix(ByteArrToLong(aByte) / (2 ^ 22))
PGDE = BaseAddress + lTemp * 4
CopyMemory PGDE, ByVal PGDE, 4
If (PGDE And 1) <> 0 Then
lTemp = PGDE And &H80
If lTemp <> 0 Then
PAddr = (PGDE And &HFFC00000) + (VAddr And &H3FFFFF)
Else
PGDE = MapViewOfFile(g_hMPM, 4, 0, PGDE And &HFFFFF000, &H1000)
lTemp = (VAddr And &H3FF000) / (2 ^ 12)
PTE = PGDE + lTemp * 4
CopyMemory PTE, ByVal PTE, 4
If (PTE And 1) <> 0 Then
PAddr = (PTE And &HFFFFF000) + (VAddr And &HFFF)
UnmapViewOfFile PGDE
End If
End If
End If
LinearToPhys = PAddr
End Function

Private Function GetData(addr As Long) As Long
Dim phys As Long, tmp As Long, Ret As Long
phys = LinearToPhys(g_pMapPhysicalMemory, addr)
tmp = MapViewOfFile(g_hMPM, 4, 0, phys And &HFFFFF000, &H1000)
If tmp <> 0 Then
Ret = tmp + ((phys And &HFFF) / (2 ^ 2)) * 4
CopyMemory Ret, ByVal Ret, 4
UnmapViewOfFile tmp
GetData = Ret
End If
End Function

Private Function SetData(ByVal addr As Long, ByVal data As Long) As Boolean
Dim phys As Long, tmp As Long, x As Long
phys = LinearToPhys(g_pMapPhysicalMemory, addr)
tmp = MapViewOfFile(g_hMPM, SECTION_MAP_WRITE, 0, phys And &HFFFFF000, &H1000)
If tmp <> 0 Then
x = tmp + ((phys And &HFFF) / (2 ^ 2)) * 4
CopyMemory ByVal x, data, 4
UnmapViewOfFile tmp
SetData = True
End If
End Function

Private Function ByteArrToLong(inByte() As Byte) As Double
Dim i As Integer
For i = 0 To 3
ByteArrToLong = ByteArrToLong + inByte(i) * (&H100 ^ i)
Next i
End Function



