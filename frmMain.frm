VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   Caption         =   "File Compare"
   ClientHeight    =   6405
   ClientLeft      =   945
   ClientTop       =   1515
   ClientWidth     =   10575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6405
   ScaleWidth      =   10575
   Begin VB.VScrollBar vscScrollBar 
      Height          =   5655
      Left            =   10320
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   6090
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14896
            MinWidth        =   5292
            Key             =   "Message"
            Object.ToolTipText     =   "Message"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12/31/2002"
            Key             =   "Date"
            Object.ToolTipText     =   "Current date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "5:22 PM"
            Key             =   "Time"
            Object.ToolTipText     =   "Current time"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open routines for compare"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print compare"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Previous"
            Object.ToolTipText     =   "Previous Difference"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "Next Difference"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   9240
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "Previous"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041C
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":052E
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0640
            Key             =   "Print"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   8640
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbFile1 
      Height          =   5115
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9022
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   99999
      TextRTF         =   $"frmMain.frx":0B82
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbFile2 
      Height          =   4500
      Left            =   5160
      TabIndex        =   3
      Top             =   840
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   7938
      _Version        =   393217
      BackColor       =   14737632
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   99999
      TextRTF         =   $"frmMain.frx":0BF2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblFile2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File 2"
      Height          =   345
      Left            =   5160
      TabIndex        =   6
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label lblFile1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File 1"
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   5055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   ""
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewNext 
         Caption         =   "&Next Instance"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewPrevious 
         Caption         =   "&Previous Instance"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrFile1 As String
Private mstrFileText1 As String
Private mstrFile2 As String
Private mstrFileText2 As String
Private mblnScrollLocked As Boolean
Private maryDifferences1() As Long      'Array to hold the line numbers for the sections that are different in file 1.
Private maryDifferences2() As Long      'Array to hold the line numbers for the sections that are different in file 2.

Private Sub Form_Load()

    Dim strPath As String
    Dim strValue As String

    If gblnOptionsStartupMaximized Then
        Me.WindowState = 2 'Maximize form
    Else
        Me.WindowState = 0 'Normal
    End If
        
    With tbToolbar
        .ImageList = ImageList
        .Buttons("Print").Image = "Print"
        .Buttons("Open").Image = "Open"
        .Buttons("Next").Image = "Next"
        .Buttons("Next").Enabled = False
        .Buttons("Previous").Image = "Previous"
        .Buttons("Previous").Enabled = False
    End With
    
    'Load window settings.
    strPath = gstrRegistryPath & "\General"
    strValue = GetFromRegistry(HKEY_CURRENT_USER, strPath, "MainWindowLeft")
    
    If Len(strValue) > 0 Then
        Me.Left = CLng(strValue)
    End If
    
    strValue = GetFromRegistry(HKEY_CURRENT_USER, strPath, "MainWindowTop")
    
    If Len(strValue) > 0 Then
        Me.Top = CLng(strValue)
    End If
    
    strValue = GetFromRegistry(HKEY_CURRENT_USER, strPath, "MainWindowWidth")
    
    If Len(strValue) > 0 Then
        Me.Width = CLng(strValue)
    End If
    
    strValue = GetFromRegistry(HKEY_CURRENT_USER, strPath, "MainWindowHeight")
    
    If Len(strValue) > 0 Then
        Me.Height = CLng(strValue)
    End If
    
    'Load recently used files.
    strPath = gstrRegistryPath & "\Recent"
    
    With gobjRecent
        .Number = 10
        .Load strPath
        .Update frmMain
    End With
    
    lblFile1.Caption = ""
    lblFile2.Caption = ""
    
    InitRichTextBoxes
        
    vscScrollBar.Tag = 1
    Form_Resize
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2

End Sub

Private Sub Form_Resize()

    Dim lngTop As Long
    Dim lngHeight As Long
    Dim lngWidth As Long
    Dim lngVSCWidth As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub

    lngVSCWidth = 250
    lngTop = tbToolbar.Height
    lngHeight = ScaleHeight - tbToolbar.Height - sbStatusBar.Height
    lngWidth = (ScaleWidth - lngVSCWidth) \ 2
    
    With vscScrollBar
        .Top = lngTop
        .Height = lngHeight
        .Left = ScaleWidth - lngVSCWidth
        .Width = lngVSCWidth
    End With
    
    lblFile1.Move 0, lngTop, lngWidth - 10, 285
    lblFile2.Move lngWidth + 10, lngTop, lngWidth - 10, 285
    
    lngTop = lngTop + 285
    rtbFile1.Move 0, lngTop, lngWidth - 10, lngHeight
    rtbFile2.Move lngWidth + 10, lngTop, lngWidth - 10, lngHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim strPath As String
    Dim intCount As Integer

    'Close all sub forms.
    For intCount = Forms.count - 1 To 1 Step -1
        Unload Forms(intCount)
    Next
    
    'Save window settings.
    If Me.WindowState <> vbMinimized Then
        strPath = gstrRegistryPath & "\General"
        SaveToRegistry HKEY_CURRENT_USER, strPath, "MainWindowLeft", CLng(Me.Left)
        SaveToRegistry HKEY_CURRENT_USER, strPath, "MainWindowTop", CLng(Me.Top)
        SaveToRegistry HKEY_CURRENT_USER, strPath, "MainWindowWidth", CLng(Me.Width)
        SaveToRegistry HKEY_CURRENT_USER, strPath, "MainWindowHeight", CLng(Me.Height)
    End If

   'Save recently opened files.
    strPath = gstrRegistryPath & "\Recent"
    gobjRecent.Save strPath

    Erase maryDifferences1
    Erase maryDifferences2
    Set gobjRecent = Nothing
    
End Sub

Private Sub mnuFile_Click()

    If ((Len(mstrFileText1) > 0) And (Len(mstrFileText2) > 0)) Then
        mnuFilePrint.Enabled = True
    Else
        mnuFilePrint.Enabled = False
    End If

End Sub

Private Sub mnuFileExit_Click()

    Unload Me
    End

End Sub

Private Sub mnuFileOpen_Click()
'Open files 1 and 2 for compare.

    Dim intFileNumber As Integer
    Dim blnSuccess As Boolean

    With tbToolbar
        .Buttons("Next").Enabled = False
        .Buttons("Previous").Enabled = False
    End With
    
    'Get file #1.
    intFileNumber = 1
    blnSuccess = GetFile(intFileNumber, mstrFile1, mstrFileText1)
    
    If Not blnSuccess Then
        sbStatusBar.Panels("Message").Text = ""
        Exit Sub
    End If
    
    sbStatusBar.Panels("Message").Text = "File #1 selected."
    
    'Get file #2.
    intFileNumber = 2
    mstrFile2 = mstrFile1
    
    blnSuccess = GetFile(intFileNumber, mstrFile2, mstrFileText2)
    
    If Not blnSuccess Then
        sbStatusBar.Panels("Message").Text = ""
        Exit Sub
    End If
    
    sbStatusBar.Panels("Message").Text = "File #2 selected. Generating compare, please wait..."
    
    'Initialize controls and perform compare.
    lblFile1.Caption = ""
    lblFile2.Caption = ""
    rtbFile1.Text = ""
    rtbFile2.Text = ""
    vscScrollBar.Tag = 1
    DoEvents
    
    Compare
    
End Sub

Private Sub mnuFilePrint_Click()

    On Error Resume Next
    
    With CommonDialog
        .DialogTitle = "Print File Compare"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums + cdlPDAllPages
        .ShowPrinter
        If Err = cdlCancel Then
            Exit Sub
        End If
    End With

    PrintCompare

End Sub

Private Sub mnuFileRecent_Click(Index As Integer)
'Open recently used file.

    Dim strFileName As String
    Dim strFileText As String
    Dim strTemp As String
    Dim I As Integer
    Dim blnSuccess As Boolean
    
    strFileName = mnuFileRecent(Index).Caption
    I = InStr(1, strFileName, " ")
    strFileName = Mid$(strFileName, I + 1, Len(strFileName))
    
    If Dir(strFileName) <> "" Then
        strFileText = LoadTextFromFile(strFileName)
        If Len(strFileText) = 0 Then
            sbStatusBar.Panels("Message").Text = ""
            Exit Sub
        End If
    Else
        MsgBox "File does not exist.", vbOKOnly + vbExclamation, "File Compare"
        Exit Sub
    End If

    If (Len(mstrFileText1) = 0) Or _
     ((Len(mstrFileText1) > 0) And (Len(mstrFileText2) > 0)) Then
            mstrFile1 = strFileName
            mstrFileText1 = strFileText
            mstrFile2 = ""
            mstrFileText2 = ""
            lblFile1.Caption = ""
            lblFile2.Caption = ""
            rtbFile1.Text = ""
            rtbFile2.Text = ""
            vscScrollBar.Tag = 1
            
            With tbToolbar
                .Buttons("Next").Enabled = False
                .Buttons("Previous").Enabled = False
            End With
            
            sbStatusBar.Panels("Message").Text = "File #1 selected, select File #2."
    Else
            mstrFile2 = strFileName
            mstrFileText2 = strFileText
            sbStatusBar.Panels("Message").Text = "File #2 selected. Generating compare, please wait..."
            DoEvents
            
            Compare
    End If

End Sub

Private Sub mnuView_Click()

    mnuViewNext.Enabled = tbToolbar.Buttons("Next").Enabled
    mnuViewPrevious.Enabled = tbToolbar.Buttons("Previous").Enabled

End Sub

Private Sub mnuViewNext_Click()
'Go to the next different section.

    Dim lngLineNum1 As Long
    Dim lngLineNum2 As Long
    Dim strTag1 As String
    Dim strTag2 As String
    Dim blnFound As Boolean
    
    'Go to the line number of the next different section
    'in routine 1.
    blnFound = NextDiff(rtbFile1, maryDifferences1, lngLineNum1)
    SetCursor rtbFile1, lngLineNum1

    'Go to the line number of the next different section
    'in routine 2.
    blnFound = NextDiff(rtbFile2, maryDifferences2, lngLineNum2)
    SetCursor rtbFile2, lngLineNum2
    
    'See if there are more differences...
    blnFound = NextDiff(rtbFile1, maryDifferences1, lngLineNum1)
    
    If Not blnFound Then
        'No more, disable the tollbar button.
        tbToolbar.Buttons("Next").Enabled = False
    End If

    'See if there are any previous differences...
    blnFound = PrevDiff(rtbFile1, maryDifferences1, lngLineNum1)
    
    If Not blnFound Then
        'No more, disable the tollbar button.
        tbToolbar.Buttons("Previous").Enabled = False
    Else
        tbToolbar.Buttons("Previous").Enabled = True
    End If

End Sub

Private Sub mnuViewPrevious_Click()
'Goto the previous different section.

    Dim lngLineNum1 As Long
    Dim lngLineNum2 As Long
    Dim blnFound As Boolean
    
    'Find the line number of the next different section
    'in routine 1.
    blnFound = PrevDiff(rtbFile1, maryDifferences1, lngLineNum1)
    SetCursor rtbFile1, lngLineNum1

    'Find the line number of the next different section
    'in routine 2.
    blnFound = PrevDiff(rtbFile2, maryDifferences2, lngLineNum2)
    SetCursor rtbFile2, lngLineNum2
    
    'See if there are more differences...
    blnFound = NextDiff(rtbFile1, maryDifferences1, lngLineNum1)
    
    If Not blnFound Then
        'No more, disable the tollbar button.
        tbToolbar.Buttons("Next").Enabled = False
    Else
        tbToolbar.Buttons("Next").Enabled = True
    End If

    'See if there are any previous differences...
    blnFound = PrevDiff(rtbFile1, maryDifferences1, lngLineNum1)
    
    If Not blnFound Then
        'No more, disable the tollbar button.
        tbToolbar.Buttons("Previous").Enabled = False
    End If

End Sub

Private Sub mnuToolsOptions_Click()
'Show the application options form.

    With frmOptions
        .Show vbModal
        If .OkPressed Then
            InitRichTextBoxes
            If ((Len(mstrFileText1) > 0) And (Len(mstrFileText2) > 0)) Then
                Compare
            End If
        End If
    End With

End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As Button)

    Select Case Button.Key
        Case Is = "Print"
            PrintCompare
        Case Is = "Open"
            mnuFileOpen_Click
        Case Is = "Next"
            mnuViewNext_Click
        Case Is = "Previous"
            mnuViewPrevious_Click
        Case Else
            MsgBox "Invalid button pressed.", vbOKOnly + vbCritical, "File Compare"
    End Select

End Sub

Private Sub vscScrollBar_Change()

    If Not mblnScrollLocked Then
        vscScrollBar_Scroll
    End If
    
End Sub

Private Sub vscScrollBar_Scroll()

    Dim lngValue As Long
    Dim lngOldValue As Long
    Dim lngDiff As Long

    mblnScrollLocked = True
    
    With vscScrollBar
        lngOldValue = CLng(.Tag)
        lngValue = .Value
    End With
    
    lngDiff = lngValue - lngOldValue
    
    Scroll rtbFile1, 0, lngDiff
    Scroll rtbFile2, 0, lngDiff
    
    vscScrollBar.Tag = lngValue
    mblnScrollLocked = False
    
End Sub

Private Sub AddToTextBox(ByVal vobjTextBox As RichTextBox, ByVal vstrText As String, ByVal vblnDifferent As Boolean)
'Add a file line to the appropriate text box.

    With vobjTextBox
        .SelStart = Len(.Text) + 1
        If vblnDifferent Then
            .SelColor = vbRed
            .SelBold = True
        Else
            .SelColor = gtypFont.FontColor
            .SelBold = False
        End If
        .SelText = vstrText
    End With

End Sub

Private Sub Compare()
'Do the compare.

    Dim strBuffer1 As String
    Dim strBuffer2 As String
    Dim intLine1 As Integer
    Dim intLine2 As Integer
    Dim intBeginLine1 As Integer
    Dim intBeginLine2 As Integer
    Dim intEndLine1 As Integer
    Dim intEndLine2 As Integer
    Dim intMaxLine1 As Integer
    Dim intMaxLine2 As Integer
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    Dim intIndex1A As Integer
    Dim intIndex2A As Integer
    Dim intIndex1B As Integer
    Dim intIndex2B As Integer
    Dim intCount1 As Integer
    Dim intCount2 As Integer
    Dim intStart1 As Integer
    Dim intStart2 As Integer
    Dim intEnd1 As Integer
    Dim intEnd2 As Integer
    Dim intDiff1Count As Integer
    Dim intDiff2Count As Integer
    Dim intNumberOfChanges As Integer
    Dim intLen As Integer
    Dim blnQuit As Boolean
    Dim blnFound As Boolean
    Dim blnSuccess As Boolean
    Dim aryFile1() As String
    Dim aryFile2() As String
    
    Screen.MousePointer = vbHourglass
    
    aryFile1 = Split(vbCrLf & mstrFileText1, vbCrLf)
    aryFile2 = Split(vbCrLf & mstrFileText2, vbCrLf)
    intMaxLine1 = UBound(aryFile1)
    intMaxLine2 = UBound(aryFile2)
    
    With vscScrollBar
        .Value = 0
        If intMaxLine1 > intMaxLine2 Then
            .Max = intMaxLine1
        Else
            .Max = intMaxLine2
        End If
    End With
    
    DoEvents
    
    LockWindow Me
    
    strBuffer1 = ""
    strBuffer2 = ""
    intDiff1Count = 0
    intDiff2Count = 0
    intNumberOfChanges = 0
    blnQuit = False
    intLine1 = 1
    intLine2 = 1
    
    Do
        If aryFile1(intLine1) = aryFile2(intLine2) Then
            'No difference.
            strBuffer1 = strBuffer1 & aryFile1(intLine1) & vbCrLf
            strBuffer2 = strBuffer2 & aryFile2(intLine2) & vbCrLf
        Else
            'Lines are different.
            
            'Dump the common lines collected so far.
            If Len(strBuffer1) > 0 Then
                AddToTextBox rtbFile1, strBuffer1, False
            End If
            
            If Len(strBuffer2) > 0 Then
                AddToTextBox rtbFile2, strBuffer2, False
            End If
            
            strBuffer1 = ""
            strBuffer2 = ""
            
            'Save these lines in the array indicating different sections.
            intDiff1Count = intDiff1Count + 1
            
            ReDim Preserve maryDifferences1(1, intDiff1Count)
            maryDifferences1(0, intDiff1Count) = intLine1
            
            intDiff2Count = intDiff2Count + 1
            
            ReDim Preserve maryDifferences2(1, intDiff2Count)
            maryDifferences2(0, intDiff2Count) = intLine2
            
            intNumberOfChanges = intNumberOfChanges + 1
            intBeginLine1 = intLine1
            intBeginLine2 = intLine2
            intEndLine1 = intMaxLine1
            intEndLine2 = intMaxLine2
            
            'Search for the next line that is the same in both routines.
            If intLine1 = intMaxLine1 Then
                intEndLine1 = intLine1
            ElseIf intLine2 = intMaxLine2 Then
                intEndLine2 = intLine2
            ElseIf aryFile1(intLine1 + 1) = aryFile2(intLine2 + 1) Then
                intEndLine1 = intLine1 + 1
                intEndLine2 = intLine2 + 1
            Else
                intStart1 = intBeginLine1
                intStart2 = intBeginLine2
                intEnd1 = intEndLine1
                intEnd2 = intEndLine2
                
                Do
                    intCount1 = 0
                    
                    For intIndex1A = intStart1 To intMaxLine1
                        blnFound = False
                        
                        For intIndex2A = intBeginLine2 To intEnd2
                            If aryFile1(intIndex1A) = aryFile2(intIndex2A) Then
                                If intIndex1A <> intMaxLine1 And (intIndex2A <> intMaxLine2) _
                                        And (aryFile1(intIndex1A + 1) = aryFile2(intIndex2A + 1)) Then
                                    If intIndex2A < intEnd2 Then
                                        intEnd2 = intIndex2A
                                        intEndLine1 = intIndex1A
                                        intEndLine2 = intIndex2A
                                        blnFound = True
                                    End If
                                    Exit For
                                End If
                            End If
                        Next intIndex2A
                        
                        If blnFound Then
                            intCount1 = intCount1 + 1
                        Else
                            Exit For
                        End If
                    Next intIndex1A
                    
                    intStart1 = intIndex1A
            
                    intCount2 = 0
                    
                    For intIndex2B = intStart2 To intMaxLine2
                        blnFound = False
                        
                        For intIndex1B = intBeginLine1 To intEnd1
                            If aryFile2(intIndex2B) = aryFile1(intIndex1B) Then
                                If intIndex2B <> intMaxLine2 And (intIndex1B <> intMaxLine1) _
                                        And (aryFile2(intIndex2B + 1) = aryFile1(intIndex1B + 1)) Then
                                    If intIndex1B < intEnd1 Then
                                        intEnd1 = intIndex1B
                                        intEndLine1 = intIndex1B
                                        intEndLine2 = intIndex2B
                                        blnFound = True
                                    End If
                                    Exit For
                                End If
                            End If
                        Next intIndex1B
                        
                        If blnFound Then
                            intCount2 = intCount2 + 1
                        Else
                            Exit For
                        End If
                    Next intIndex2B
                    
                    intStart2 = intIndex2B
                    
                    If intCount1 <> 0 Then Exit Do
                    If intCount2 <> 0 Then Exit Do
                    intStart1 = intStart1 + 1
                    intStart2 = intStart2 + 1
                    
                    If (intStart1 > intMaxLine1) And (intStart2 > intMaxLine2) Then
                        Exit Do
                    End If
                Loop
                
                If intCount1 = 0 Then intCount1 = intMaxLine1
                If intCount2 = 0 Then intCount2 = intMaxLine2
            End If
            
            maryDifferences1(1, intDiff1Count) = intEndLine1
            maryDifferences2(1, intDiff2Count) = intEndLine2

            'Display the lines in both files up to that point.
            For intIndex1 = intBeginLine1 To (intEndLine1 - 1)
                strBuffer1 = strBuffer1 & aryFile1(intIndex1) & vbCrLf
            Next
            
            AddToTextBox rtbFile1, strBuffer1, True                        'Changed lines.
            AddToTextBox rtbFile1, aryFile1(intEndLine1) & vbCrLf, False   'Common line.
            
            For intIndex2 = intBeginLine2 To (intEndLine2 - 1)
                strBuffer2 = strBuffer2 & aryFile2(intIndex2) & vbCrLf
            Next
            
            AddToTextBox rtbFile2, strBuffer2, True                        'Changed lines.
            AddToTextBox rtbFile2, aryFile2(intEndLine2) & vbCrLf, False   'Common line.
            strBuffer1 = ""
            strBuffer2 = ""
            intLine1 = intEndLine1
            intLine2 = intEndLine2
        End If
        
        intLine1 = intLine1 + 1
        intLine2 = intLine2 + 1
        If intLine1 > intMaxLine1 Then blnQuit = True
        If intLine2 > intMaxLine2 Then blnQuit = True
    Loop Until blnQuit
    
    'Dump anything left in the buffers.
    If Len(strBuffer1) > 0 Then
       AddToTextBox rtbFile1, strBuffer1, False
    End If
    If Len(strBuffer2) > 0 Then
        AddToTextBox rtbFile2, strBuffer2, False
    End If
    
    lblFile1.Caption = TruncateString(lblFile1, mstrFile1)
    lblFile2.Caption = TruncateString(lblFile2, mstrFile2)
        
    rtbFile2.SelStart = 0
    
    With rtbFile1
        .SelStart = 0
        .SetFocus
    End With
    
    Select Case intNumberOfChanges
        Case 0
            sbStatusBar.Panels("Message").Text = "No differences found."
            With tbToolbar
                .Buttons("Next").Enabled = False
                .Buttons("Previous").Enabled = False
            End With
        Case 1
            sbStatusBar.Panels("Message").Text = "Total of " & CStr(intNumberOfChanges) & " difference found."
            With tbToolbar
                .Buttons("Next").Enabled = True
                .Buttons("Previous").Enabled = False
            End With
        Case Else
            sbStatusBar.Panels("Message").Text = "Total of " & CStr(intNumberOfChanges) & " differences found."
            With tbToolbar
                .Buttons("Next").Enabled = True
                .Buttons("Previous").Enabled = False
            End With
    End Select
    
    UnlockWindow
    Screen.MousePointer = vbDefault
        
End Sub

Private Function GetFile(ByVal vintFileNumber, ByRef rstrFileName As String, ByRef rstrFileText As String) As Boolean

    Dim strFileNumber As String
    Dim strFileName As String
    Dim strDirectory As String
    Dim intPieces As Integer
    
    On Error Resume Next
    
    If Len(rstrFileName) = 0 Then
        strFileName = rstrFileName
    Else
        intPieces = NumPc(rstrFileName, "\")
        strFileName = Pc(rstrFileName, "\", intPieces)
    End If
    
    strFileNumber = "File #" & CStr(vintFileNumber)
    
    If vintFileNumber = 1 Then
        strDirectory = gstrOptionsDefaultDir1
    Else
        strDirectory = gstrOptionsDefaultDir2
    End If
    
    With CommonDialog
        .DialogTitle = "Select " & strFileNumber & " for Compare"
        .CancelError = True
        .DefaultExt = ".txt"
        .FileName = strFileName
        .Filter = "Text Files(" & gstrOptionsTextFileTypes & ")|" & gstrOptionsTextFileTypes & "|All Files(*.*)|*.*"
        .Flags = cdlOFNFileMustExist Or cdlOFNExplorer
        .InitDir = strDirectory
        .ShowOpen
        
        If Err = cdlCancel Then
            GetFile = False
            Exit Function
        End If
        
        rstrFileName = .FileName
    End With
    
    rstrFileText = LoadTextFromFile(rstrFileName)

    With gobjRecent
        .Add rstrFileName
        .Update frmMain
    End With
    
    GetFile = True

End Function

Private Sub InitRichTextBoxes()
'Set the font properties of the richtext boxes and initialize.

    With rtbFile1
        .Text = ""
        With .Font
            .Name = gtypFont.FontName
            .Size = gtypFont.FontSize
            .Bold = gtypFont.FontBold
            .Italic = gtypFont.FontItalic
            .Strikethrough = gtypFont.FontStrikethru
            .Underline = gtypFont.FontUnderline
        End With
    End With
    
    With rtbFile2
        .Text = ""
        With .Font
            .Name = gtypFont.FontName
            .Size = gtypFont.FontSize
            .Bold = gtypFont.FontBold
            .Italic = gtypFont.FontItalic
            .Strikethrough = gtypFont.FontStrikethru
            .Underline = gtypFont.FontUnderline
        End With
    End With

End Sub

Private Function LoadTextFromFile(ByVal vstrFileName As String) As String

    Dim strFileText As String
    Dim intFileNum As Integer

    On Error GoTo LoadTextFromFile_Error
    
    intFileNum = FreeFile
    Open vstrFileName For Input As intFileNum
    strFileText = Input$(LOF(intFileNum), intFileNum)
    Close intFileNum
    LoadTextFromFile = strFileText
    Exit Function

LoadTextFromFile_Error:
    
    MsgBox "Error " & Format$(Err.Number) & " opening file." & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "File Compare"
    LoadTextFromFile = ""
    Exit Function
    
End Function

Private Function NextDiff(ByVal vobjTextBox As RichTextBox, ByRef raryDiff() As Long, ByRef rlngLineNum As Long) As Boolean
'Find the line number of the next different section in file 1.
'Returns True if a next section is found, else False if no more.
'Returns the starting line number for the next different section.
    
    Dim intMax As Integer
    Dim intIndex As Integer
    
    NextDiff = False
    
    intMax = UBound(raryDiff, 2)
    rlngLineNum = GetLineNumber(vobjTextBox) + 1   'Current line number.
    
    For intIndex = 1 To intMax
        If raryDiff(0, intIndex) > rlngLineNum Then
            rlngLineNum = raryDiff(0, intIndex)
            NextDiff = True
            Exit For
        End If
    Next intIndex

End Function

Private Function PrevDiff(ByVal vobjTextBox As RichTextBox, ByRef raryDiff() As Long, ByRef rlngLineNum As Long) As Boolean
'Find the line number of the previous different section in file 1.
'Returns True if a next section is found, else False if no more.
'Returns the starting line number for the previous different section.
    
    Dim intMax As Integer
    Dim intIndex As Integer
    
    PrevDiff = False
    
    intMax = UBound(raryDiff, 2)
    rlngLineNum = GetLineNumber(vobjTextBox) + 1       'Current line number.
    
    For intIndex = intMax To 1 Step -1
        If raryDiff(0, intIndex) < rlngLineNum Then
            rlngLineNum = raryDiff(0, intIndex)
            PrevDiff = True
            Exit For
        End If
    Next intIndex

End Function

Private Sub PrintCompare()

    Dim strLine As String
    Dim lngBegLine As Long
    Dim lngEndLine As Long
    Dim lngLine As Long
    Dim intIndex As Integer
    Dim intMax As Integer
    Dim intLen As Integer
    Dim strAst As String
    Dim strDash As String
    Dim aryFile1() As String
    Dim aryFile2() As String
    Dim objPrint As New clsPrintControl

    With objPrint
        .PrintDateTime = True
        .PrintPageNumber = True
        .FontName = "Courier New"
        .FontSize = 9
        .Titles(1) = "File Compare"
        .Titles(2) = "File 1: " & lblFile1.Caption
        .Titles(3) = "File 2: " & lblFile2.Caption
    End With

    strAst = String(80, "*")
    strDash = String(80, "-")
    intMax = UBound(maryDifferences1, 2)
    aryFile1 = Split(vbCrLf & mstrFileText1, vbCrLf)
    aryFile2 = Split(vbCrLf & mstrFileText2, vbCrLf)

    For intIndex = 1 To intMax
        intLen = (80 - Len(mstrFile1) - 2) \ 2
        objPrint.PrintLine Mid$(strAst, 1, intLen) & " " & mstrFile1 & " " & Mid$(strAst, 1, intLen), "Left"
        
        'Print file 1 differences.
        lngBegLine = maryDifferences1(0, intIndex)
        lngEndLine = maryDifferences1(1, intIndex)
        For lngLine = lngBegLine To lngEndLine
            strLine = aryFile1(lngLine)
            objPrint.PrintLine strLine, "Left", True
        Next lngLine
        objPrint.PrintLine "", "Left"
    
        'Print file 2 differences.
        intLen = (80 - Len(mstrFile2) - 2) \ 2
        objPrint.PrintLine Mid$(strDash, 1, intLen) & " " & mstrFile2 & " " & Mid$(strDash, 1, intLen), "Left"
        lngBegLine = maryDifferences2(0, intIndex)
        lngEndLine = maryDifferences2(1, intIndex)
        For lngLine = lngBegLine To lngEndLine
            strLine = aryFile2(lngLine)
            objPrint.PrintLine strLine, "Left", True
        Next lngLine
    
        objPrint.PrintLine strAst & vbCrLf, "Left"
    Next intIndex

    objPrint.PrintLine vbCrLf, "Left"
    objPrint.PrintLine sbStatusBar.Panels("Message").Text, "Left"
    objPrint.EndPrint
    
    Set objPrint = Nothing

End Sub

