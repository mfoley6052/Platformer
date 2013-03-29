VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   737
   StartUpPosition =   3  'Windows Default
   Begin VB.Line ln 
      Index           =   0
      X1              =   64
      X2              =   280
      Y1              =   16
      Y2              =   16
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim numLines As Integer
Dim x As Integer
Attribute x.VB_VarHelpID = -1

Private Sub Form_Load()
numLines = InputBox("How many lines?")
For x = 1 To numLines
    Load ln(x)
    With ln(x)
        .X1 = ln(x - 1).X2
        .Y1 = ln(x - 1).Y2
        .Visible = True
        .X2 = ln(x - 1).X1 + 100
        .Y2 = ln(x - 1).Y1 + 30
    End With
Next x
End Sub
