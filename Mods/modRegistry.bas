Attribute VB_Name = "mdlRegistry"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private Const REG_SZ                         As Long = 1
Private Const REG_DWORD                      As Long = 4
Public Const HKEY_CURRENT_USER               As Long = &H80000001
Private Const ERROR_NONE                     As Integer = 0
Private Const KEY_ALL_ACCESS                 As Long = &H3F
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                            ByVal lpValueName As String, _
                                                                                            ByVal lpReserved As Long, _
                                                                                            lpType As Long, _
                                                                                            ByVal lpData As String, _
                                                                                            lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                          ByVal lpValueName As String, _
                                                                                          ByVal lpReserved As Long, _
                                                                                          lpType As Long, _
                                                                                          lpData As Long, _
                                                                                          lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                          ByVal lpValueName As String, _
                                                                                          ByVal lpReserved As Long, _
                                                                                          lpType As Long, _
                                                                                          ByVal lpData As Long, _
                                                                                          lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                        ByVal lpValueName As String, _
                                                                                        ByVal Reserved As Long, _
                                                                                        ByVal dwType As Long, _
                                                                                        ByVal lpValue As String, _
                                                                                        ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal Reserved As Long, _
                                                                                      ByVal dwType As Long, _
                                                                                      lpValue As Long, _
                                                                                      ByVal cbData As Long) As Long
''Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'<:-):WARNING: Unused Declare RegDeleteKey
'<:-)May be a prototype Declare you have not yet implimented or left over from a deleted Control.
'<:-):UPDATED: Obsolete Type Suffix replaced.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                                                                    ByVal lpValueName As String) As Long
'<:-):UPDATED: Obsolete Type Suffix replaced.

Public Sub DeleteValue(lPredefinedKey As Long, _
                       sKeyName As String, _
                       sValueName As String)

  '<:-):WARNING: Function changed to Sub as nothing is returned via the Function Name.
  ' Description:
  '   This Function will delete a value
  '
  ' Syntax:
  '   DeleteValue Location, KeyName, ValueName
  '
  '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
  '   , HKEY_USERS
  '
  '   KeyName is the name of the key that the value you wish to delete is in
  '   , it may include subkeys (example "Key1\SubKey1")
  '
  '   ValueName is the name of value you wish to delete
  '<:-) Missing Dims Auto-inserted lRetVal As Long
  
  Dim lRetVal As Long

    '<:-):WARNING:  As Type may not be correct
    '<:-)May be Control name using default property or a control (probably Form) Property being used with default assignment (not a good idea; be explicit)
  Dim hKey    As Long            'handle of open key
    'open the specified key
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = RegDeleteValue(hKey, sValueName)
    RegCloseKey (hKey)

End Sub

Public Sub SetKeyValue(lPredefinedKey As Long, _
                       sKeyName As String, _
                       sValueName As String, _
                       vValueSetting As Variant, _
                       lValueType As Long)

  '<:-):WARNING: Function changed to Sub as nothing is returned via the Function Name.
  ' Description:
  '   This Function will set the data field of a value
  '
  ' Syntax:
  '   QueryValue Location, KeyName, ValueName, ValueSetting, ValueType
  '
  '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
  '   , HKEY_USERS
  '
  '   KeyName is the key that the value is under (example: "Key1\SubKey1")
  '
  '   ValueName is the name of the value you want create, or set the value of (example: "ValueTest")
  '
  '   ValueSetting is what you want the value to equal
  '
  '   ValueType must equal either REG_SZ (a string) Or REG_DWORD (an integer)
  '<:-) Missing Dims Auto-inserted lRetVal As Long
  
  Dim lRetVal As Long

    '<:-):WARNING:  As Type may not be correct
    '<:-)May be Control name using default property or a control (probably Form) Property being used with default assignment (not a good idea; be explicit)
  Dim hKey    As Long            'handle of open key
    'open the specified key
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)

End Sub

Private Function SetValueEx(ByVal hKey As Long, _
                            sValueName As String, _
                            lType As Long, _
                            vValue As Variant) As Long

  '<:-):WARNING: Scope Too Large. Reduced to Private '<:-)May be a prototype you have not yet implimented or left over from a deleted Control.
  
  Dim lValue As Long
  Dim sValue As String

    Select Case lType
     Case REG_SZ
        sValue = vValue
        SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
     Case REG_DWORD
        lValue = vValue
        SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select

End Function

''
''Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
'''<:-):WARNING: Unused Function QueryValueEx
'''<:-)May be a prototype Function you have not yet implimented or left over from a deleted Control.
'''<:-):WARNING: Scope Changed to Private
''Dim cch    As Long
''Dim lrc    As Long
''Dim lType  As Long
''Dim lValue As Long
''Dim sValue As String
''On Error GoTo QueryValueExError
''' Determine the size and type of data to be read
''lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
''If lrc <> ERROR_NONE Then
''Error 5
''End If
''Select Case lType
''' For strings
''Case REG_SZ
''sValue = String$(cch, 0)
''lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
''If lrc = ERROR_NONE Then
''vValue = Left$(sValue, cch)
''Else
''vValue = Empty
''End If
''' For DWORDS
''Case REG_DWORD
''lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
''If lrc = ERROR_NONE Then
''vValue = lValue
''End If
''Case Else
'''all other data types not supported
''lrc = -1
''End Select
''QueryValueExExit:
''QueryValueEx = lrc
''Exit Function
''QueryValueExError:
''Resume QueryValueExExit
''End Function
''
''
''Public Sub CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
'''<:-):WARNING: Unused Sub CreateNewKey
'''<:-)May be a prototype Sub you have not yet implimented or left over from a deleted Control.
'''<:-):WARNING: Function changed to Sub as nothing is returned via the Function Name.
''' Description:
'''   This Function will create a new key
'''
''' Syntax:
'''   QueryValue Location, KeyName
'''
'''   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'''   , HKEY_USERS
'''
'''   KeyName is name of the key you wish to create, it may include subkeys (example "Key1\SubKey1")
'''<:-) Missing Dims Auto-inserted lRetVal As Long
''Dim lRetVal As Long
'''<:-):WARNING:  As Type may not be correct
'''<:-)May be Control name using default property or a control (probably Form) Property being used with default assignment (not a good idea; be explicit)
''Dim hNewKey As Long         'handle to the new key
''lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
''RegCloseKey (hNewKey)
''End Function
''
''
''Public Sub DeleteKey(lPredefinedKey As Long, sKeyName As String)
'''<:-):WARNING: Unused Sub DeleteKey
'''<:-)May be a prototype Sub you have not yet implimented or left over from a deleted Control.
'''<:-):WARNING: Function changed to Sub as nothing is returned via the Function Name.
''' Description:
'''   This Function will Delete a key
'''
''' Syntax:
'''   DeleteKey Location, KeyName
'''
'''   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'''   , HKEY_USERS
'''
'''   KeyName is name of the key you wish to delete, it may include subkeys (example "Key1\SubKey1")
'''open the specified key
'''lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
'''<:-) Missing Dims Auto-inserted lRetVal As Long
''Dim lRetVal As Long
'''<:-):WARNING:  As Type may not be correct
'''<:-)May be Control name using default property or a control (probably Form) Property being used with default assignment (not a good idea; be explicit)
''lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
'''RegCloseKey (hKey)
''End Function
''
''
''Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, ByVal sValueName As String)
'''<:-):WARNING: Unused Function QueryValue
'''<:-)May be a prototype Function you have not yet implimented or left over from a deleted Control.
'''<:-):SUGGESTION: Function should be TypeCase but Code Fixer cannot determine the Type to apply.
'''<:-):WARNING: Function will return Variant value.
'''<:-):WARNING: 'ByVal ' inserted for Parameter  'sValueName As String'
''' Description:
'''   This Function will return the data field of a value
'''
''' Syntax:
'''   Variable = QueryValue(Location, KeyName, ValueName)
'''
'''   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'''   , HKEY_USERS
'''
'''   KeyName is the key that the value is under (example: "Software\Microsoft\Windows\CurrentVersion\Explorer")
'''
'''   ValueName is the name of the value you want to access (example: "link")
'''<:-) Missing Dims Auto-inserted lRetVal As Long
''Dim lRetVal As Long
'''<:-):WARNING:  As Type may not be correct
'''<:-)May be Control name using default property or a control (probably Form) Property being used with default assignment (not a good idea; be explicit)
''Dim hKey    As Long           'handle of opened key
''Dim vValue  As Variant       'setting of queried value
''lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
''lRetVal = QueryValueEx(hKey, sValueName, vValue)
'''MsgBox vValue
''QueryValue = vValue
''RegCloseKey (hKey)
''End Function
''
':)Roja's VB Code Fixer V1.1.78 (31.01.2004 18:34:23) 64 + 208 = 272 Lines Thanks Ulli for inspiration and lots of code.

