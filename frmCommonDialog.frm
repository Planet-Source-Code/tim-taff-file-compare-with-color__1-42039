VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCommonDialog 
   BorderStyle     =   0  'None
   ClientHeight    =   765
   ClientLeft      =   3330
   ClientTop       =   1845
   ClientWidth     =   1725
   ControlBox      =   0   'False
   Icon            =   "frmCommonDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   765
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

    Me.Hide

End Sub

Private Sub Form_Load()

    Move Screen.Width / 6, Screen.Height / 6
    
End Sub
