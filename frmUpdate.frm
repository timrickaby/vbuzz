VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   3
   Caption         =   "vbGo VBuzz"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   ControlBox      =   0
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0
   MinButton       =   0
   ScaleHeight     =   2970
   ScaleWidth      =   5880
   ShowInTaskbar   =   0
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStatusBack 
      Height          =   255
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   1320
      Width           =   4695
      Begin VB.PictureBox picStatusFore 
         BackColor       =   &H8000000D&
         BorderStyle     =   0 
         Height          =   200
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   735
         TabIndex        =   2
         Top             =   0
         Width           =   735
      End
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
      TabIndex        =   0
      Top             =   2280
      Visible         =   0
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmUpdate.frx":0000
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label lblMessageBody 
      BackStyle       =   0
      Caption         =   "Please wait while VBuzz connects to vbGo and checks for a newer version of your content viewer."
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
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label lblMessageTitle 
      BackStyle       =   0
      Caption         =   "Updating Current Viewer..."
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
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   4695
   End
   Begin VB.Line lneWhite 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   240
      X2              =   5640
      Y1              =   2055
      Y2              =   2055
   End
   Begin VB.Line lneGray 
      BorderColor     =   &H80000010&
      BorderStyle     =   6
      Index           =   0
      X1              =   240
      X2              =   5640
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image imgDownload 
      Height          =   480
      Left            =   240
      Picture         =   "frmUpdate.frx":044A
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblStatusLabel 
      AutoSize        =   -1
      BackStyle       =   0
      Caption         =   "Establishing Connection to vbGo..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   2760
   End
End

Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' vbGo
' VBuzz - The VB Information Platform
' Copyright © 1999 - 2001 vbGo. All Rights Reserved
' Version 1.0 Build 2

Option Explicit