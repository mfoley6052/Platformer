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
      MultiLine       =   -1  'True
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
Dim filenum(1) As Integer
Dim verNum As String
Dim oldtext As String
Dim strPath As String
Dim revision() As String

Private Function Rfh() As Boolean
Dim temp As String
'On Error GoTo terminate
If Dir(strPath) <> "" Then
    filenum(0) = FreeFile
    Open strPath For Input As #filenum(0)
    Line Input #filenum(0), verNum
    txtVersion.Text = verNum
    Do Until EOF(filenum(0))
        Line Input #filenum(0), temp
        oldtext = oldtext & vbclrf & temp
    Loop
    Close #filenum(0)
If Dir(App.Path & "\Changelog.old") <> "" Then
    Kill App.Path & "\Changelog.old"
    Name App.Path & "\Changelog.txt" As App.Path & "\Changelog.old"
End If
    SetAttr App.Path & "\Changelog.old", vbHidden
    
Else
    Open strPath For Output As #1
    Close #1
End If
'terminate:
'For x = 0 To 1
'    Close #x
'Next x
End Function

Private Sub cmdRefresh_Click()
Call Rfh

End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo terminate
    Call rvsn
    filenum(1) = FreeFile
    Open strPath For Append As #filenum(1)
    If txtChange.Text <> "" Then
        Print #filenum(1), verNum
        Print #filenum(1), "####################################################################"
        Print #filenum(1), txtChange.Text
        Print #filenum(1), oldtext
    End If
    Close #filenum(1)
    txtChange.Text = ""
terminate:
For x = 0 To 1
    Close #x
Next x
txtVersion = verNum
End Sub

Private Sub Form_Load()
strPath = App.Path & "\Changelog.txt"
Call Rfh
'Call rvsn
End Sub

Private Sub Form_Terminate()
For x = 0 To 1
    Close #x
Next x
End Sub
Private Function rvsn() As Boolean

    revision = Split(verNum, ".")
   ' ReDim revision(3) As String
    If Val(revision(3)) < 99 Then
        revision(3) = Val(revision(3)) + 1
    Else
        revision(3) = 0
        If Val(revision(2)) < 99 Then
            revision(2) = Val(revision(2)) + 1
        Else
            revision(2) = 0
            If Val(revision(1)) < 99 Then
                revision(1) = Val(revision(1)) + 1
            Else
                revision(1) = 0
                revision(0) = Val(revision(0)) + 1
            End If
        End If
    End If
    verNum = revision(0) & "." & revision(1) & "." & revision(2) & "." & revision(3)
   ' MsgBox (verNum)
End Function

