Attribute VB_Name = "basRegistry"
'Read and Write to any part of the Windows Registry.
'
'Inputs:None
'Returns:None
'Assumes:None
'
Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1 ' Unicode nul terminated string
Public Const REG_DWORD = 4 ' 32-bit number

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function GetFromRegistry(ByVal vlngHKey As Long, ByVal vstrPath As String, ByVal vstrValue As String) As String
    'EXAMPLE:
    '
    'text1.text = GetFromRegistry(HKEY_CURRENT_USER, "Software\VBW\Registry", "String")
    '
    Dim lngKeyHand As Long
    Dim lngValueType As Long
    Dim lngResult As Long
    Dim strBuf As String
    Dim lngDataBufSize As Long
    Dim intZeroPos As Integer
    
    lngResult = RegOpenKey(vlngHKey, vstrPath, lngKeyHand)
    lngResult = RegQueryValueEx(lngKeyHand, vstrValue, 0&, lngValueType, ByVal 0&, lngDataBufSize)

    If lngValueType = REG_SZ Then
        strBuf = String(lngDataBufSize, " ")
        lngResult = RegQueryValueEx(lngKeyHand, vstrValue, 0&, 0&, ByVal strBuf, lngDataBufSize)
        If lngResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetFromRegistry = Left$(strBuf, intZeroPos - 1)
            Else
                GetFromRegistry = strBuf
            End If
        End If
    End If

End Function

Public Function GetAllFromRegistry(ByVal vlngHKey As Long, ByVal vstrPath As String) As String()
'Returns all values for a key in an array.

    Dim lngResult As Long
    Dim lngKeyHand As Long
    Dim lngCount As Long
    Dim strBuffer As String
    Dim strItem As String
    Dim strValue As String
    Dim aryTemp() As String
    
    ReDim aryTemp(0)
    
    RegOpenKey vlngHKey, vstrPath, lngKeyHand
    lngCount = 0
    
    Do
        'Create a buffer.
        strBuffer = String(255, 0)
        'Enumerate the values.
        If RegEnumValue(lngKeyHand, lngCount, strBuffer, 255, 0, ByVal 0&, ByVal 0&, ByVal 0&) <> 0 Then
            Exit Do
        End If
        lngCount = lngCount + 1
        ReDim Preserve aryTemp(lngCount)
        strItem = StripTerminator(strBuffer)
        strValue = GetFromRegistry(vlngHKey, vstrPath, strItem)
        aryTemp(lngCount) = strItem & " = " & strValue
    Loop
    
    'Close the registry.
    RegCloseKey vlngHKey
    
    GetAllFromRegistry = aryTemp

End Function

Public Sub SaveToRegistry(ByVal vlngHKey As Long, ByVal vstrPath As String, ByVal vstrValue As String, ByVal vstrData As String)
    'EXAMPLE:
    '
    'Call SaveToRegistry(HKEY_CURRENT_USER, "Software\VBW\Registry", "String", text1.text)
    '
    Dim lngKeyHand As Long
    Dim lngResult As Long
    
    lngResult = RegCreateKey(vlngHKey, vstrPath, lngKeyHand)
    lngResult = RegSetValueEx(lngKeyHand, vstrValue, 0, REG_SZ, ByVal vstrData, Len(vstrData))
    lngResult = RegCloseKey(lngKeyHand)

End Sub

Public Function GetDWordFromRegistry(ByVal vlngHKey As Long, ByVal vstrPath As String, ByVal vstrValueName As String) As Long
    'EXAMPLE:
    '
    'text1.text = GetDWordFromRegistry(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
    '
    Dim lngResult As Long
    Dim lngValueType As Long
    Dim lngBuf As Long
    Dim lngDataBufSize As Long
    Dim lngKeyHand As Long
    
    lngResult = RegOpenKey(vlngHKey, vstrPath, lngKeyHand)
    ' Get length/data type
    lngDataBufSize = 4
    lngResult = RegQueryValueEx(lngKeyHand, vstrValueName, 0&, lngValueType, lngBuf, lngDataBufSize)

    If lngResult = ERROR_SUCCESS Then
        If lngValueType = REG_DWORD Then
            GetDWordFromRegistry = lngBuf
        End If
    End If
    
    lngResult = RegCloseKey(lngKeyHand)

End Function

Public Sub SaveDWordToRegistry(ByVal vlngHKey As Long, ByVal vstrPath As String, ByVal vstrValueName As String, ByVal vlngData As Long)
    'EXAMPLE"
    '
    'Call SaveDWordToRegistry(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword", text1.text)
    '
    Dim lngResult As Long
    Dim lngKeyHand As Long
    
    lngResult = RegCreateKey(vlngHKey, vstrPath, lngKeyHand)
    lngResult = RegSetValueEx(lngKeyHand, vstrValueName, 0&, REG_DWORD, vlngData, 4)
    lngResult = RegCloseKey(lngKeyHand)

End Sub

Public Sub DeleteRegistryKey(ByVal vlngHKey As Long, ByVal vstrKey As String)
    'EXAMPLE:
    '
    'Call DeleteRegistryKey(HKEY_CURRENT_USER, "Software\VBW")
    '
    Dim lngResult As Long
    
    lngResult = RegDeleteKey(vlngHKey, vstrKey)

End Sub

Public Sub DeleteRegistryValue(ByVal vlngHKey As Long, ByVal vstrPath As String, ByVal vstrValue As String)
    'EXAMPLE:
    '
    'Call DeleteRegistryValue(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
    '
    Dim lngKeyHand As Long
    Dim lngResult As Long
    
    lngResult = RegOpenKey(vlngHKey, vstrPath, lngKeyHand)
    lngResult = RegDeleteValue(lngKeyHand, vstrValue)
    lngResult = RegCloseKey(lngKeyHand)

End Sub

Private Function StripTerminator(ByVal vstrInput As String) As String
    
    Dim intZeroPos As Integer
    
    'Search the first chr$(0).
    intZeroPos = InStr(1, vstrInput, vbNullChar)
    
    If intZeroPos > 0 Then
        StripTerminator = Left$(vstrInput, intZeroPos - 1)
    Else
        StripTerminator = vstrInput
    End If

End Function

