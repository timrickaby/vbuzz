VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0
   Caption         =   "vbGo - VBuzz"
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0
   MinButton       =   0
   ScaleHeight     =   3765
   ScaleWidth      =   6750
   ShowInTaskbar   =   0
   StartUpPosition =   2
   Begin VB.Timer tmrSplashScreen 
      Interval        =   3500
      Left            =   6120
      Top             =   600
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   6740
      X2              =   6740
      Y1              =   3720
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   15
      X2              =   6720
      Y1              =   3750
      Y2              =   3750
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   6360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Top             =   0
      Width           =   6750
   End
End

Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' vbGo
' VBuzz - The VB Information Platform
' Copyright © 1999 - 2001 vbGo. All Rights Reserved
' Version 1.0 Build 2

Option Explicit


Private Sub tmrSplashScreen_Timer()
   ' Show splash screen for 3.5 seconds (see timer property)
    Unload frmSplash
    frmMain.Show
End Sub
