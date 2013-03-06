VERSION 5.00
Begin VB.Form frmUpdater 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Changelog Updater"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtVersion 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtChange 
      Height          =   1575
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opened As Boolean
Dim filenum As Integer
Dim verNum As String
Dim oldtxt As String

Private Function Rfh() As Boolean
    filenum = FreeFile
    Open App.Path & "\Changelog.txt" For Input As #filenum
    Line Input #filenum, verNum
    oldtxt = Input(LOF(1), #filenum)
    txtVersion.Text = verNum
    Close #filenum
End Function

Private Sub cmdRefresh_Click()
Call Rfh

End Sub

Private Sub cmdUpdate_Click()
    Open App.Path & "\Changelog.txt" For Append As #filenum
    If txtChange.Text <> "" Then
        Print #1, txtChange.Text
        Print
End Sub
