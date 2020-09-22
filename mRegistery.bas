Attribute VB_Name = "mRegistery"
'Author      : Abdalla Mahmoud
'Age         : 14  06\07\1988
'Country     : Egypt
'Citry       : Mansoura
'E-Mail      : la3toot@hotmail.com  la3toot@yahoo.com
'**********************************
'This Function Include an amzing Function That Delete Reg
'And All Its SubKeys In WinXP
'Note : In WinXP You Can Not Delete Key With all
'It's subKeys With The Ordinary Function
'So This unction Will Help You
'**********************************
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Long, ByVal lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Enum hKeyType
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_user = &H80000001
    HKEY_DYN_DATA = &H80000006
End Enum

Public Enum hValueType
     REG_BINARY = 3
     REG_DWORD = 4
     REG_DWORD_BIG_ENDIAN = 5
     REG_DWORD_LITTLE_ENDIAN = 4
     REG_EXPAND_SZ = 2
     REG_SZ = 1
End Enum

Function ReadReg(Key As hKeyType, SubKey As String, ValueKey As String, Optional ValueType As hValueType = REG_SZ) As Variant
Dim hKey As Long
Dim Ret  As Long
Dim Res As String

Ret = RegOpenKey(Key, SubKey, hKey)
Res = String(255, Chr(0))
Call RegQueryValueEx(hKey, ValueKey, 0, ValueType, ByVal Res, 255)
RegCloseKey hKey
ReadReg = Mid(Res, 1, InStr(Res, Chr(0)) - 1)
End Function

Function WriteReg(Key As hKeyType, SubKey As String, ValueKey As String, Optional Value As Variant, Optional ValueType As hValueType = REG_SZ) As Boolean
Dim hKey As Long
Dim Ret  As Long
Dim Res As String

Ret = RegOpenKey(Key, SubKey, hKey)
Res = Value & String(255 - Len(Value), Chr(0))
Call RegSetValueEx(hKey, ValueKey, 0, ValueType, ByVal Res, 255)
RegCloseKey hKey
WriteReg = Not (Ret)
End Function

Function CreateKey(Key As hKeyType, SubKey As String) As Boolean
Dim Ret  As Long

Call RegCreateKey(Key, SubKey, Ret)
CreateKey = Ret
End Function

Function EnumKeys(Key As hKeyType, Optional SubKey As String) As Collection
Dim Res As Long
Dim Ret As String
Dim I As Long
Dim Coll As New Collection
Dim hKey As Long

Res = RegOpenKey(Key, SubKey, hKey)
Ret = SubKey & String(255, Chr(0))
Res = RegEnumKey(hKey, I, Ret, Len(Ret) + 255)
Do While Res = 0
    I = I + 1
    Coll.Add Mid(Ret, 1, InStr(Ret, Chr(0)) - 1)
    Res = RegEnumKey(hKey, I, Ret, Len(Ret) + 255)
Loop
RegCloseKey hKey
Set EnumKeys = Coll
Set Coll = Nothing
End Function

Function EnumValues(Key As hKeyType, SubKey As String) As Collection
Dim Res As Long
Dim Ret As String
Dim I As Long
Dim Coll As New Collection
Dim hKey As Long

Res = RegOpenKey(Key, SubKey, hKey)
Ret = String(255, Chr(0))
Res = RegEnumValue(hKey, I, Ret, 255, 0&, 0&, 0&, 0&)
Do While Res = 0
    I = I + 1
    Coll.Add Mid(Ret, 1, InStr(Ret, Chr(0)) - 1)
    Res = RegEnumValue(hKey, I, Ret, Len(Ret), 0, 0, 0, 0)
Loop
RegCloseKey hKey
Set EnumValues = Coll
Set Coll = Nothing
End Function

Function IfKeyExists(Key As hKeyType, SubKey As String) As Boolean
Dim Res As Long
Dim hKey As Long

Res = RegOpenKey(Key, SubKey, hKey)
If Res Then Exit Function
RegCloseKey hKey
IfKeyExists = True
End Function

Function DeleteKey(Key As hKeyType, SubKey As String) As Boolean
DeleteKey = Not (RegDeleteKey(Key, SubKey))
End Function

Function DeleteValue(Key As hKeyType, SubKey As String, Optional ValueName As String) As Boolean
Dim hKey As Long

If RegOpenKey(Key, SubKey, hKey) Then Exit Function
DeleteValue = Not (RegDeleteValue(hKey, ValueName))
RegCloseKey hKey
End Function

Function DeleteKeyWinXP(Key As hKeyType, SubKey As String) As Boolean
Dim Coll As New Collection
Dim I As Long

Set Coll = GetAllKeys(Key, SubKey)
For I = Coll.Count To 1 Step -1
    DeleteKey Key, Coll(I)
Next
End Function

Function GetAllKeys(Key As hKeyType, SubKey As String) As Collection
Dim Coll As New Collection
Dim TmpColl As Collection
Dim I As Long
Dim Pos As Long

Pos = 1
Coll.Add SubKey
Do
    If Pos > Coll.Count Then GoTo Finish
    Set TmpColl = EnumKeys(Key, Coll(Pos))
    For I = 1 To TmpColl.Count
        Coll.Add Coll(Pos) & "\" & TmpColl(I)
    Next
    Set TmpColl = Nothing
    Pos = Pos + 1
Loop
Finish:
Set GetAllKeys = Coll
Set Coll = Nothing
End Function
