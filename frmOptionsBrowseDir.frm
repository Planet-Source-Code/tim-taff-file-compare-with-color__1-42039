VERSION 5.00
Begin VB.Form frmOptionsBrowseDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Folder..."
   ClientHeight    =   4065
   ClientLeft      =   4020
   ClientTop       =   1965
   ClientWidth     =   6360
   Icon            =   "frmOptionsBrowseDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4065
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFolderName 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   3600
      Width           =   3855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.DirListBox dirLook 
      Height          =   2955
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4815
   End
   Begin VB.DriveListBox drvChange 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Folder &name:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Look &in:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmOptionsBrowseDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOkPressed As Boolean

Private Sub cmdCancel_Click()

    mblnOkPressed = False
    Hide
    
End Sub

Private Sub cmdOk_Click()

    mblnOkPressed = True
    Hide
    
End Sub

Private Sub dirLook_Change()
    
    On Error GoTo DirError
    
    txtFolderName.Text = dirLook.Path
    Exit Sub

DirError:
    MsgBox Err.Description, vbOKOnly, "Select Drive"
    Resume Next

End Sub

Private Sub drvChange_Change()
  
    On Error GoTo DriveError
    
    dirLook.Path = drvChange.Drive
    Exit Sub

DriveError:
    MsgBox Err.Description, vbOKOnly, "Select Drive"
    Resume Next

End Sub

Private Sub Form_Load()

    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        cmdCancel.Value = True
    End If

End Sub

Public Property Get OkPressed() As Boolean

    OkPressed = mblnOkPressed
    
End Property

