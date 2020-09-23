VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Basic print class.

Option Explicit

Private mblnPrintDateTime As Boolean
Private mblnPrintPageNumber As Boolean
Private mblnPrintSelection As Boolean
Private msngPageWidth As Single
Private msngPrintWidth As Single
Private msngPageHeight As Single
Private msngPrintHeight As Single
Private msngHorizMargin As Single
Private msngVertMargin As Single
Private msngHorizBorder As Single
Private msngVertBorder As Single
Private maryColumnWidth() As Single
Private maryColumnHeaders() As String
Private maryTitles() As String

Private Sub AlignLeft()
'Sets the print position to left margin of page.

    Printer.CurrentX = msngHorizMargin

End Sub

Private Sub AlignRight(ByVal strText As String)
'Sets the print position to right margin of page.

    Dim sngTextWidth As Single
    
    sngTextWidth = Printer.TextWidth(strText)
    Printer.CurrentX = msngPageWidth - msngHorizBorder - msngHorizMargin - sngTextWidth

End Sub

Private Sub AlignCenter(ByVal strText As String)
'Sets the print position to center of page.

    Dim sngTextWidth As Single
    
    sngTextWidth = Printer.TextWidth(strText)
    Printer.CurrentX = (msngPageWidth - sngTextWidth) \ 2

End Sub

Private Function FitText(ByRef strText As String, ByVal sngWidth As Single) As String
'Return as much text as will fit in this width.
'strText is returned with the remainder of the string that
'will not fit.
    
    Dim strTemp As String
    Dim lngPos As Long

    strTemp = strText
    lngPos = Len(strText)
    
    'Leave a little white space between each column - 0.2 cm.
    Do While Printer.TextWidth(strText) > (sngWidth - 0.2)
        lngPos = Len(strText) - 1
        strText = Left$(strText, lngPos)
    Loop
    
    FitText = strText
    
    If lngPos < Len(strTemp) Then
        strText = Mid$(strTemp, lngPos + 1)
    Else
        strText = ""
    End If

End Function

Private Function PageFeed() As Boolean
'Returns True if current Y position is greater than page length.

    PageFeed = False
    
    If Printer.CurrentY > msngPrintHeight Then
        PageFeed = True
    End If

End Function

Private Sub TopOfForm()
'Print headers at top of form.

    Dim strText As String
    Dim strTemp As String
    Dim sngTextWidth As Single
    Dim I As Integer
    Dim intColumnCount As Integer
    Dim sngX As Single

    Printer.Print "";
    
    'Print today's date.
    If mblnPrintDateTime Then
        AlignLeft
        Printer.CurrentY = msngVertMargin
        Printer.Print Format(Now, "m/d/yyyy");
    End If
    
    'Print title line 1.
    AlignCenter maryTitles(1)
    Printer.Print maryTitles(1);

    'Print page number.
    If mblnPrintPageNumber Then
        strText = "Page " & Printer.Page
        AlignRight strText
        Printer.Print strText
    Else
        Printer.Print ""
    End If
    
    'Print current time.
    If mblnPrintDateTime Then
        AlignLeft
        Printer.Print Format(Now, "Hh:Nn");
    End If
    
    'Print title line 2 through n.
    For I = 2 To UBound(maryTitles)
        AlignCenter maryTitles(I)
        Printer.Print maryTitles(I)
    Next I

    Printer.Print ""
    Printer.Print ""

    'Print column headers.
    intColumnCount = UBound(maryColumnHeaders)
    If intColumnCount > 0 Then
        sngX = msngHorizMargin
        For I = 1 To intColumnCount
            strTemp = maryColumnHeaders(I)
            strText = FitText(strTemp, maryColumnWidth(I))
            Printer.CurrentX = sngX
            Printer.Print strText;
            sngX = sngX + maryColumnWidth(I)
        Next I
        Printer.Print ""
        Printer.Print ""
    End If

End Sub

Public Property Get ColumnWidth(ByVal intColumnNumber As Integer) As Single

    ColumnWidth = maryColumnWidth(intColumnNumber)

End Property

Public Property Let ColumnWidth(ByVal intColumnNumber As Integer, ByVal sngNewValue As Single)

    ReDim Preserve maryColumnWidth(intColumnNumber)
    maryColumnWidth(intColumnNumber) = sngNewValue

End Property

Public Property Get ColumnHeaders(ByVal intColumnNumber As Integer) As String
    
    ColumnHeaders = maryColumnHeaders(intColumnNumber)

End Property

Public Property Let ColumnHeaders(ByVal intColumnNumber As Integer, ByVal strNewValue As String)

    ReDim Preserve maryColumnHeaders(intColumnNumber)
    maryColumnHeaders(intColumnNumber) = strNewValue

End Property

Public Sub EndPrint()

    Printer.EndDoc

End Sub

Public Property Get FontName() As String

    FontName = Printer.FontName

End Property

Public Property Let FontName(ByVal strNewValue As String)

    Printer.FontName = strNewValue

End Property

Public Property Get FontSize() As Single

    FontSize = Printer.FontSize

End Property

Public Property Let FontSize(ByVal sngNewValue As Single)

    Printer.FontSize = sngNewValue

End Property

Public Property Get HorizontalMargin() As Single

    HorizontalMargin = msngHorizMargin

End Property

Public Property Let HorizontalMargin(ByVal sngNewValue As Single)

    msngHorizMargin = sngNewValue - msngHorizBorder
    msngPrintWidth = msngPageWidth - msngHorizBorder - msngHorizMargin - msngHorizMargin

End Property

Public Property Get PageHeight() As Single

    PageHeight = msngPageHeight

End Property

Public Property Let PageHeight(ByVal sngNewValue As Single)

    msngPageHeight = sngNewValue
    msngPrintHeight = msngPageHeight - msngVertBorder - msngVertMargin - msngVertBorder - msngVertMargin

End Property

Public Property Get PageWidth() As Single

    PageWidth = msngPageWidth
    
End Property

Public Property Let PageWidth(ByVal sngNewValue As Single)

    msngPageWidth = sngNewValue
    msngPrintWidth = msngPageWidth - msngHorizBorder - msngHorizMargin - msngHorizMargin

End Property

Public Property Get PrintDateTime() As Boolean

    PrintDateTime = mblnPrintDateTime

End Property

Public Property Let PrintDateTime(ByVal blnNewValue As Boolean)

    mblnPrintDateTime = blnNewValue

End Property

Public Property Get PrintPageNumber() As Boolean

    PrintPageNumber = mblnPrintPageNumber

End Property

Public Property Let PrintPageNumber(ByVal blnNewValue As Boolean)

    mblnPrintPageNumber = blnNewValue
    
End Property

Public Property Get PrintSelection() As Boolean

    PrintSelection = mblnPrintSelection

End Property

Public Property Let PrintSelection(ByVal blnNewValue As Boolean)

    mblnPrintSelection = blnNewValue

End Property

Public Property Get Titles(ByVal intIndex As Integer) As String

    Titles = maryTitles(intIndex)

End Property

Public Property Let Titles(ByVal intIndex As Integer, ByVal strNewValue As String)

    ReDim Preserve maryTitles(intIndex)
    maryTitles(intIndex) = strNewValue

End Property

Public Property Get VerticalMargin() As Single

    VerticalMargin = msngVertMargin

End Property

Public Property Let VerticalMargin(ByVal sngNewValue As Single)

    msngVertMargin = sngNewValue - msngVertBorder
    msngPrintHeight = msngPageHeight - msngVertBorder - msngVertMargin - msngVertBorder - msngVertMargin

End Property

Public Sub PrintLine(ByVal strLine As String, ByVal strAlignment As String, Optional ByVal blnAutoWrap As Boolean = False)
'Print one line.

    Dim strTemp As String
    Static blnTOF As Boolean
    
    If Not blnTOF Then
        TopOfForm
        blnTOF = True
    End If

    Do
        strTemp = FitText(strLine, msngPrintWidth)
        If Not blnAutoWrap Then
            strLine = ""
        End If
        If PageFeed Then
            Printer.NewPage
            TopOfForm
        End If
        Select Case strAlignment
            Case "Right"
                AlignRight strTemp
            Case "Center"
                AlignCenter strTemp
            Case Else
                AlignLeft
        End Select
        Printer.Print strTemp
    Loop Until Len(strLine) = 0

End Sub

Public Sub PrintListView(ByRef objListView As ListView)
'Print contents of a ListView control.

    Dim strItem As String
    Dim strText As String
    Dim strTemp As String
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    Dim intItemCount As Integer
    Dim intColumn As Integer
    Dim intColumnCount As Integer
    Dim sngPageWidth As Single
    Dim sngX As Single
    Dim blnPrint As Boolean

    Printer.Print "";
    
    With objListView
        intColumnCount = .ColumnHeaders.count
        intItemCount = .ListItems.count
        
        If .View = lvwReport Then       'Detail view.
            'Load the column headers.
            For intColumn = 1 To intColumnCount
                strTemp = .ColumnHeaders(intColumn).Text
                If UBound(maryColumnWidth) > 0 Then
                    strItem = FitText(strTemp, maryColumnWidth(intColumn))
                Else
                    strItem = strTemp
                    maryColumnWidth(intColumn) = Printer.TextWidth(strItem)
                End If
                ReDim Preserve maryColumnHeaders(intColumn)
                maryColumnHeaders(intColumn) = strItem
            Next intColumn
            
            TopOfForm
            
            'Print each item with sub-items.
            For intIndex1 = 1 To intItemCount
                blnPrint = True
                If mblnPrintSelection Then
                    If Not .ListItems(intIndex1).Selected Then
                        blnPrint = False
                    End If
                End If
                If blnPrint Then
                    If PageFeed Then
                        Printer.NewPage
                        TopOfForm
                    End If
                    strTemp = .ListItems(intIndex1).Text
                    strItem = FitText(strTemp, maryColumnWidth(1))
                    sngX = msngHorizMargin
                    AlignLeft
                    Printer.Print strItem;
                    sngX = msngHorizMargin + maryColumnWidth(1)
                    
                    'Print sub-items.
                    For intColumn = 2 To intColumnCount
                        intIndex2 = intColumn - 1
                        strTemp = .ListItems(intIndex1).SubItems(intIndex2)
                        strItem = FitText(strTemp, maryColumnWidth(intColumn))
                        Printer.CurrentX = sngX
                        Printer.Print strItem;
                        sngX = sngX + maryColumnWidth(intColumn)
                    Next intColumn
                    
                    Printer.Print ""
                End If
            Next intIndex1
        Else                            'List view.
            TopOfForm
            strText = ""
            For intIndex1 = 1 To intItemCount
                blnPrint = True
                If mblnPrintSelection Then
                    If Not .ListItems(intIndex1).Selected Then
                        blnPrint = False
                    End If
                End If
                If blnPrint Then
                    strItem = .ListItems(intIndex1).Text
                    strItem = Mid$(strItem & "          ", 1, 10)
                    If (Printer.TextWidth(strText) + Printer.TextWidth(strItem)) > msngPrintWidth Then
                        If PageFeed Then
                            Printer.NewPage
                            TopOfForm
                        End If
                        AlignLeft
                        Printer.Print strText
                        strText = ""
                    End If
                    strText = strText & strItem
                End If
            Next intIndex1
            If Len(strText) > 0 Then
                If PageFeed Then
                    Printer.NewPage
                    TopOfForm
                End If
                AlignLeft
                Printer.Print strText
            End If
        End If
    End With
    
    Printer.Print ""
    Printer.Print ""
    AlignLeft
    Printer.Print intItemCount & " items"
    Printer.Print ""
    Printer.EndDoc

End Sub

Public Sub PrintText(ByVal strText As String, ByVal blnAutoWrap As Boolean)
'Print contents of a textbox.
'If blnAutowrap = True, wrap the line to the next line.

    Dim strLine As String
    Dim strTemp As String
    Dim aryLines() As String
    Dim varLine As Variant
    
    aryLines = Split(strText, vbCrLf)

    TopOfForm

    For Each varLine In aryLines
        strLine = CStr(varLine)
        Do
            strTemp = FitText(strLine, msngPrintWidth)
            If Not blnAutoWrap Then
                strLine = ""
            End If
            If PageFeed Then
                Printer.NewPage
                TopOfForm
            End If
            AlignLeft
            Printer.Print strTemp
        Loop Until Len(strLine) = 0
    Next

    Printer.EndDoc

End Sub

Public Sub PrintTreeView(ByRef objTreeview As Object, ByVal blnAutoWrap As Boolean)
'Print contents of a treeview control.

    Dim tvwNode As Node
    Dim blnPrint As Boolean

    TopOfForm
    
    'Print each top level node recursively.
    Set tvwNode = objTreeview.Nodes(1)
    
    Do Until tvwNode Is Nothing
        blnPrint = True
        If mblnPrintSelection Then
            If Not tvwNode.Selected Then
                blnPrint = False
            End If
        End If
        If blnPrint Then
            AlignLeft
            PrintTreeViewNode tvwNode, blnAutoWrap
        End If
        Set tvwNode = tvwNode.Next
    Loop
        
    Printer.EndDoc
    Set tvwNode = Nothing

End Sub

Private Sub PrintTreeViewNode(ByVal tvwNode As Node, ByVal blnAutoWrap As Boolean)
'Recursively print a treeview node and its children.

    Const INDENT = 0.5
    
    Dim strText As String
    Dim strTemp As String
    Dim sngMinX As Single
    Dim sngChildX As Single
    Dim sngBaseY As Single
    Dim sngLineX As Single
    Dim sngLineHeight As Single
    Dim sngY As Single
    Dim blnPrint As Boolean
    
    sngMinX = Printer.CurrentX

    With tvwNode
        strText = .Text
        If .Children > 0 Then
            If .Expanded Then
                strText = "-" & strText
            Else
                strText = "+" & strText
            End If
        End If
    End With
    
    'Print the node.
    Do
        strTemp = FitText(strText, msngPrintWidth)
        If Not blnAutoWrap Then
            strText = ""
        End If
        If PageFeed Then
            Printer.NewPage
            TopOfForm
        End If
        Printer.CurrentX = sngMinX
        Printer.Print strTemp
    Loop Until Len(strText) = 0
    
    'Print the node's children.
    If tvwNode.Expanded Then
        sngChildX = sngMinX + INDENT
        sngLineX = sngMinX + INDENT / 2
        sngLineHeight = Printer.TextHeight("X")
        sngBaseY = Printer.CurrentY
        Set tvwNode = tvwNode.Child
        Do Until tvwNode Is Nothing
            blnPrint = True
            If mblnPrintSelection Then
                If Not tvwNode.Selected Then
                    blnPrint = False
                End If
            End If
            If blnPrint Then
                'Draw a line to the child.
                sngY = Printer.CurrentY
                Printer.Line (sngLineX, sngBaseY)-(sngLineX, Printer.CurrentY + sngLineHeight / 2)
                Printer.Line -Step(INDENT / 2, 0)
        
                'Draw the child recursively.
                Printer.CurrentY = sngY
                Printer.CurrentX = sngChildX
                PrintTreeViewNode tvwNode, blnAutoWrap
            End If
            Set tvwNode = tvwNode.Next
        Loop
    End If

End Sub

Private Sub Class_Initialize()
'Set default values.

    mblnPrintDateTime = True
    mblnPrintPageNumber = True
    mblnPrintSelection = False
    
    Printer.ScaleMode = vbCentimeters
    msngPageWidth = 21.5    '21.5 cm = 8.5 inches
    msngPageHeight = 27.9   '27.9 cm = 11.0 inches
    
    'Physical borders of the page for this printer.
    msngHorizBorder = (msngPageWidth - Printer.ScaleWidth) / 2
    msngVertBorder = (msngPageHeight - Printer.ScaleHeight) / 2
    
    msngHorizMargin = 2.5 - msngHorizBorder
    msngVertMargin = 2.5 - msngVertBorder
    msngPrintWidth = msngPageWidth - msngHorizBorder - msngHorizMargin - msngHorizMargin
    msngPrintHeight = msngPageHeight - msngVertBorder - msngVertMargin - msngVertMargin

    'Initialize arrays.
    ReDim maryColumnWidth(0)
    ReDim maryColumnHeaders(0)

    'Always assume at least one title line.
    ReDim maryTitles(1)

End Sub
