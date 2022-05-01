VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3
   Caption         =   "vbGo VBuzz"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   ControlBox      =   0
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0
   MinButton       =   0
   Moveable        =   0
   ScaleHeight     =   2625
   ScaleWidth      =   5895
   ShowInTaskbar   =   0
   StartUpPosition =   2
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Visible         =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   1920
      Visible         =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   1920
      Visible         =   0
      Width           =   1455
   End
   Begin VB.Label lblMessageBody 
      BackStyle       =   0
      Caption         =   "Message Body"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label lblMessageTitle 
      BackStyle       =   0
      Caption         =   "Message Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   4575
   End
   Begin VB.Image imgInternet 
      Height          =   480
      Left            =   240
      Picture         =   "frmMessage.frx":0000
      Top             =   240
      Visible         =   0
      Width           =   480
   End
   Begin VB.Line lneWhite 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   240
      X2              =   5640
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Line lneGray 
      BorderColor     =   &H80000010&
      BorderStyle     =   6
      Index           =   0
      X1              =   240
      X2              =   5640
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Image imgError 
      Height          =   480
      Left            =   240
      Picture         =   "frmMessage.frx":0442
      Top             =   240
      Visible         =   0
      Width           =   480
   End
   Begin VB.Image imgInformation 
      Height          =   480
      Left            =   240
      Picture         =   "frmMessage.frx":0884
      Top             =   240
      Visible         =   0
      Width           =   480
   End
   Begin VB.Image imgQuestion 
      Height          =   480
      Left            =   240
      Picture         =   "frmMessage.frx":0CC6
      Top             =   240
      Visible         =   0
      Width           =   480
   End
   Begin VB.Image imgAlert 
      Height          =   480
      Left            =   240
      Picture         =   "frmMessage.frx":1108
      Top             =   240
      Visible         =   0
      Width           =   480
   End
End

Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' vbGo
' VBuzz - The Personal VB Information Platform
' Copyright © 1999 - 2001 vbGo. All Rights Reserved

Option Explicit

Public strReturnType As String

Private Sub cmdOK_Click()
    strReturnType = "BUTTONOK"
    Unload frmMessage
End Sub
Private Sub cmdCancel_Click()
    strReturnType = "BUTTONCANCEL"
    Unload frmMessage
End Sub
Private Sub cmdClose_Click()
    strReturnType = "BUTTONCLOSE"
    Unload frmMessage
End Sub
