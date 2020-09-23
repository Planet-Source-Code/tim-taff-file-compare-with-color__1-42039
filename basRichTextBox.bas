Attribute VB_Name = "basRichTextBox"
'Functions for manipulating a richtextbox.

Option Explicit

Private Declare Function SendMessageAsLong Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As Long) As Long

Private Declare Function SendMessageAsString Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As String) As Long

Private Const EM_CANUNDO = &HC6
Private Const EM_UNDO = &HC7
Private Const EM_EMPTYUNDOBUFFER = &HCD
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_FMTLINES = &HC8
Private Const EM_GETLINE = &HC4
Private Const EM_GETRECT = &HB2
Private Const EM_SCROLL = &HB5
Private Const EM_SCROLLCARET = &HB7
Private Const EM_SETRECT = &HB3
Private Const EM_SETTABSTOPS = &HCB
Private Const EM_LINESCROLL = &HB6
Private Const WM_USER = &H400
Private Const EM_HIDESELECTION = WM_USER + 63

Private Const MAX_CHARS_PER_LINE = 255

Public Function FirstVisibleLine(ByVal vobjTextBox As RichTextBox) As Long
' Return the index of the first visible line
' (0 for the first text line in the control).
' When applied to a single-line control, return the
' index of the first visible character
' (0 for the first character in the control).
    
    FirstVisibleLine = SendMessageAsLong(vobjTextBox.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0)

End Function

Public Function GetColumnNumber(ByVal vobjTextBox As RichTextBox, ByVal vlngLineNum As Long) As Long
'Get the current column number.

    'Subtract the line's start index from the cursor position.
    GetColumnNumber = vobjTextBox.SelStart - LineIndex(vobjTextBox, vlngLineNum)

End Function

Public Function GetLine(ByVal vobjTextBox As RichTextBox, ByVal vlngLineNum As Long) As String
' Return the specified line text.
    
    Dim ByteLo As Byte
    Dim ByteHi As Byte
    Dim strBuffer As String
    Dim lngReturnedSize As Long
    
    ByteLo = MAX_CHARS_PER_LINE And (255)
    ByteHi = Int(MAX_CHARS_PER_LINE / 256)
    strBuffer = Chr$(ByteLo) + Chr$(ByteHi) + Space$(MAX_CHARS_PER_LINE - 2)
    
    lngReturnedSize = SendMessageAsString(vobjTextBox.hwnd, EM_GETLINE, vlngLineNum, strBuffer)
    GetLine = Left$(strBuffer, lngReturnedSize)
    If Right$(GetLine, 2) = vbCrLf Then GetLine = Left$(GetLine, Len(GetLine) - 2)

End Function

Public Function GetLineNumber(ByVal vobjTextBox As RichTextBox) As Long
' Return the current line number.
    
    GetLineNumber = SendMessageAsLong(vobjTextBox.hwnd, EM_LINEFROMCHAR, -1, 0)

End Function

Public Property Get LineCount(ByVal vobjTextBox As RichTextBox) As Long
' Return the number of lines in the control.
    
    LineCount = SendMessageAsLong(vobjTextBox.hwnd, EM_GETLINECOUNT, 0, 0)

End Property

Public Function LineIndex(ByVal vobjTextBox As RichTextBox, ByVal vlngLineNum As Long) As Long
' Return the character offset of the first character of a line.
    
    LineIndex = SendMessageAsLong(vobjTextBox.hwnd, EM_LINEINDEX, vlngLineNum, 0)

End Function

Public Function LineLength(ByVal vobjTextBox As RichTextBox, ByVal vlngLineNum As Long) As Long
' Return the length of the specified line.
    
    Dim lngCharOffset As Long
    
    'Retrieve the character offset of the first character of the line.
    lngCharOffset = LineIndex(vobjTextBox, vlngLineNum)
    
    'Now retrieve the length of the line.
    LineLength = SendMessageAsLong(vobjTextBox.hwnd, EM_LINELENGTH, lngCharOffset, 0)

End Function

Public Sub Scroll(ByVal vobjTextBox As RichTextBox, ByVal vlngHorizScroll As Long, ByVal vlngVertScroll As Long)
' Scroll the contents of the control.
' Positive values scroll left and up, negative values
' scroll right and down.
' NOTE: per Microsoft MSDN article the horizontal scroll does
'      nor work with rich edit controls. TTT 1/26/2001.
    
    SendMessageAsLong vobjTextBox.hwnd, EM_LINESCROLL, vlngHorizScroll, vlngVertScroll

End Sub

Public Sub ScrollToLine(ByVal vobjTextBox As RichTextBox, ByVal vlngTopLine As Long)
'Scroll the contents of the control so that the
'line passed is the first visible line.

    Dim lngLine As Long
    
    lngLine = GetLineNumber(vobjTextBox)
    
    SendMessageAsLong vobjTextBox.hwnd, EM_LINESCROLL, 0, vlngTopLine - lngLine

End Sub

Public Sub ScrollToTop(ByVal vobjTextBox As RichTextBox)
'Scroll the contents of the control so that the cursor
'is at the top line of the window.

    Dim lngTopLine As Long
    Dim lngLine As Long
    
    lngLine = GetLineNumber(vobjTextBox)
    lngTopLine = FirstVisibleLine(vobjTextBox)
    
    SendMessageAsLong vobjTextBox.hwnd, EM_LINESCROLL, 0, lngLine - lngTopLine

End Sub

Public Sub SetCursor(ByVal vobjTextBox As RichTextBox, ByVal rlngLineNum As Long)
'Set the new cursor position.
    
    Dim lngPosition As Long
    
    With vobjTextBox
        lngPosition = LineIndex(vobjTextBox, rlngLineNum - 1)
        .SelStart = lngPosition
        ScrollToTop vobjTextBox
    End With

End Sub

