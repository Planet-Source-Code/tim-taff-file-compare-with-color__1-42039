Attribute VB_Name = "basFileCompare"
'Compare two files and display each file side-by-side with differences colored.

Option Explicit

'Global variables to hold application options saved in registry.
Public gstrRegistryPath As String
Public gstrOptionsTextFileTypes As String
Public gstrOptionsDefaultDir1 As String
Public gstrOptionsDefaultDir2 As String
Public gblnOptionsStartupMaximized As Boolean

Public Type FontProperties
    FontName As String
    FontSize As Integer
    FontColor As Long
    FontBold As Boolean
    FontItalic As Boolean
    FontStrikethru As Boolean
    FontUnderline As Boolean
End Type

Public gtypFont As FontProperties

'Recently opened files object.
Public gobjRecent As New clsRecent

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long

Public Sub Main()
'Main entry point.

    On Error GoTo MainError

    Set gobjRecent = Nothing
    Call LoadIni
    frmMain.Show

    Exit Sub

MainError:
    MsgBox "Error " & Format$(Err.Number) & vbCrLf & Err.Description, vbOKOnly + vbCritical, "File Compare"

End Sub

Public Sub LoadIni()
'Load settings from registry.

    Dim strPath As String
    Dim strTemp As String
    Dim I As Integer
    Dim aryTemp As Variant
    
    gstrRegistryPath = "Software\" & App.ProductName

    'Load General options.
    strPath = gstrRegistryPath & "\General"
    gstrOptionsTextFileTypes = GetFromRegistry(HKEY_CURRENT_USER, strPath, "TextFileTypes")
    gstrOptionsDefaultDir1 = GetFromRegistry(HKEY_CURRENT_USER, strPath, "DefaultRoutineDirectory1")
    gstrOptionsDefaultDir2 = GetFromRegistry(HKEY_CURRENT_USER, strPath, "DefaultRoutineDirectory2")
    strTemp = GetFromRegistry(HKEY_CURRENT_USER, strPath, "StartupMaximized")
    
    If Len(strTemp) > 0 Then
        gblnOptionsStartupMaximized = CBool(strTemp)
    Else
        gblnOptionsStartupMaximized = False
    End If

    'Load font and properties.
    strTemp = GetFromRegistry(HKEY_CURRENT_USER, strPath, "Font")

    'Default font.
    If Len(strTemp) = 0 Then
        strTemp = "Courier New;9;0;False;False;False;False"
    End If
    
    With gtypFont
        .FontName = Pc(strTemp, ";", 1)
        .FontSize = CInt(Pc(strTemp, ";", 2))
        .FontColor = CLng(Pc(strTemp, ";", 3))
        .FontBold = CBool(Pc(strTemp, ";", 4))
        .FontItalic = CBool(Pc(strTemp, ";", 5))
        .FontStrikethru = CBool(Pc(strTemp, ";", 6))
        .FontUnderline = CBool(Pc(strTemp, ";", 7))
    End With

End Sub

Public Sub LockWindow(ByVal vobjForm As Form)
'Prevents updates to the window while processes
'are being performed on it.

    LockWindowUpdate vobjForm.hwnd

End Sub

Public Function TruncateString(ByVal vobjControl As Object, ByVal vstrFileName As String) As String
'Truncate a file name from the left, adding '...'.

    Dim strTemp As String
    Dim intPieceNum As Integer
    Dim intNumPieces As Integer
    Dim intMaxLength As Integer
    
    '76 twips per character (8 pt)
    intMaxLength = vobjControl.Width \ 76
    
    strTemp = vstrFileName
    
    If Len(strTemp) > intMaxLength Then
        intNumPieces = NumPc(strTemp, "\")
        
        For intPieceNum = 2 To intNumPieces
            strTemp = SetPc(strTemp, "\", intPieceNum, "...")
            If Len(strTemp) <= intMaxLength Then
                Exit For
            End If
        Next
    End If
    
    TruncateString = strTemp

End Function

Public Sub UnlockWindow()
'Allow updates to the window to occur again.

    LockWindowUpdate 0

End Sub

