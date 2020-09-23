VERSION 5.00
Begin VB.UserControl Hline 
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1635
   ScaleHeight     =   540
   ScaleWidth      =   1635
   ToolboxBitmap   =   "ctrlLiner.ctx":0000
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   240
      X2              =   1560
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   1560
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "Hline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''*****************************************************************
'' File Name: ctrlLiner.ctl
'' Purpose: Control used to draw a border line

Option Explicit


Private Sub UserControl_Paint()
Line1.x1 = 0
Line1.y1 = 0
Line1.x2 = UserControl.Width
Line1.y2 = 0

Line2.x1 = 0
Line2.y1 = 20
Line2.x2 = UserControl.Width
Line2.y2 = 20
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 30
    UserControl_Paint
End Sub
