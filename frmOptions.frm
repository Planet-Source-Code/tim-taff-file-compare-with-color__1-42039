VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3225
   ClientLeft      =   1860
   ClientTop       =   1890
   ClientWidth     =   7770
   ForeColor       =   &H00000021&
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3225
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFont 
      Caption         =   "&Font..."
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3840
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox txtFileTypes 
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Text            =   "*.bat;*.ini;*.txt"
      Top             =   165
      Width           =   5295
   End
   Begin VB.CommandButton cmdBrowse2 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtDefaultDirectory2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   6495
   End
   Begin VB.CommandButton cmdBrowse1 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtDefaultDirectory1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6495
   End
   Begin VB.CheckBox chkStartMaximized 
      Caption         =   "&Maximize window on startup"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "A&pply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Text File Types:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Default Routine &2 Directory:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Default Routine &1 Directory:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOkPressed As Boolean

Private Sub chkStartMaximized_Click()

    cmdApply.Enabled = True

End Sub

Private Sub cmdApply_Click()

    Dim strDir As String
    Dim strTemp As String
    
    If ValidDirectories = True Then
        Call SaveOptions
        cmdApply.Enabled = False
    End If

End Sub

Private Sub cmdBrowse1_Click()

    Load frmOptionsBrowseDir
    
    With frmOptionsBrowseDir
        .dirLook.Path = txtDefaultDirectory1.Text
        .Show vbModal
        If .OkPressed Then
            txtDefaultDirectory1.Text = .txtFolderName
        End If
    End With

    Unload frmOptionsBrowseDir
    Set frmOptionsBrowseDir = Nothing
    
End Sub

Private Sub cmdBrowse2_Click()

    Load frmOptionsBrowseDir
    
    With frmOptionsBrowseDir
        .dirLook.Path = txtDefaultDirectory2.Tag
        .Show vbModal
        If .OkPressed Then
            txtDefaultDirectory2.Text = .txtFolderName
        End If
    End With

    Unload frmOptionsBrowseDir
    Set frmOptionsBrowseDir = Nothing
    
End Sub

Private Sub cmdCancel_Click()

    mblnOkPressed = False
    Hide
    
End Sub

Private Sub cmdFont_Click()

    On Error Resume Next
        
    With CommonDialog
        .FontName = gtypFont.FontName
        .FontSize = gtypFont.FontSize
        .Color = gtypFont.FontColor
        .FontBold = gtypFont.FontBold
        .FontItalic = gtypFont.FontItalic
        .FontStrikethru = gtypFont.FontStrikethru
        .FontUnderline = gtypFont.FontUnderline
        .Flags = cdlCFScreenFonts Or cdlCFANSIOnly Or cdlCFEffects Or cdlCFForceFontExist
        .ShowFont
        If Err = cdlCancel Then
            Exit Sub
        End If
        gtypFont.FontName = .FontName
        gtypFont.FontSize = .FontSize
        gtypFont.FontColor = .Color
        gtypFont.FontBold = .FontBold
        gtypFont.FontItalic = .FontItalic
        gtypFont.FontStrikethru = .FontStrikethru
        gtypFont.FontUnderline = .FontUnderline
    End With
    
    cmdApply.Enabled = True

End Sub

Private Sub cmdOk_Click()

    Dim strTemp As String
    
    If ValidDirectories = True Then
        Call SaveOptions
    End If

    mblnOkPressed = True
    Hide
    
End Sub

Private Sub Form_Load()

    If Len(gstrOptionsTextFileTypes) > 0 Then
        txtFileTypes.Text = gstrOptionsTextFileTypes
    End If
    
    txtDefaultDirectory1.Text = gstrOptionsDefaultDir1
    txtDefaultDirectory2.Text = gstrOptionsDefaultDir2
    
    If gblnOptionsStartupMaximized Then
        chkStartMaximized.Value = 1
    Else
        chkStartMaximized.Value = 0
    End If

    cmdApply.Enabled = False
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
End Sub

Private Sub txtDefaultDirectory1_Change()

    cmdApply.Enabled = True

End Sub

Private Sub txtDefaultDirectory2_Change()

    cmdApply.Enabled = True

End Sub

Private Sub SaveOptions()

    Dim strPath As String
    Dim strTemp As String
    
    On Error Resume Next

    'Save General options.
    gstrOptionsTextFileTypes = txtFileTypes.Text
    gstrOptionsDefaultDir1 = txtDefaultDirectory1.Text
    gstrOptionsDefaultDir2 = txtDefaultDirectory2.Text
    gblnOptionsStartupMaximized = CBool(chkStartMaximized.Value)
    
    strPath = gstrRegistryPath & "\General"
    SaveToRegistry HKEY_CURRENT_USER, strPath, "TextFileTypes", gstrOptionsTextFileTypes
    SaveToRegistry HKEY_CURRENT_USER, strPath, "DefaultRoutineDirectory1", gstrOptionsDefaultDir1
    SaveToRegistry HKEY_CURRENT_USER, strPath, "DefaultRoutineDirectory2", gstrOptionsDefaultDir2
    SaveToRegistry HKEY_CURRENT_USER, strPath, "StartupMaximized", gblnOptionsStartupMaximized

    'Save font properties.
    With gtypFont
        strTemp = CStr(.FontName) & ";" _
                & CStr(.FontSize) & ";" _
                & CStr(.FontColor) & ";" _
                & CStr(.FontBold) & ";" _
                & CStr(.FontItalic) & ";" _
                & CStr(.FontStrikethru) & ";" _
                & CStr(.FontUnderline)
    End With
    
    SaveToRegistry HKEY_CURRENT_USER, strPath, "Font", strTemp

End Sub

Private Function ValidDirectories() As Boolean

    Dim strDir As String
    Dim strTemp As String
    
    strDir = txtDefaultDirectory1.Text
    
    If Len(strDir) > 0 Then
        strTemp = Dir$(strDir, vbDirectory)
        If Len(strTemp) = 0 Then
            MsgBox "Default directory 1 does not exist.", vbExclamation + vbOKOnly, "Error in Directory Name"
            txtDefaultDirectory1.SetFocus
            ValidDirectories = False
            Exit Function
        End If
    End If
    
    strDir = txtDefaultDirectory2.Text
    
    If Len(strDir) > 0 Then
        strTemp = Dir$(strDir, vbDirectory)
        If Len(strTemp) = 0 Then
            MsgBox "Default directory 2 does not exist.", vbExclamation + vbOKOnly, "Error in Directory Name"
            txtDefaultDirectory2.SetFocus
            ValidDirectories = False
            Exit Function
        End If
    End If
    
    ValidDirectories = True
    
End Function

Public Property Get OkPressed() As Boolean

    OkPressed = mblnOkPressed
    
End Property


