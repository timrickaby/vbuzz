VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1
   BorderStyle     =   1
   Caption         =   "vbGo VBuzz"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   ControlBox      =   0
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0
      Italic          =   0
      Strikethrough   =   0
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0
   MinButton       =   0
   Moveable        =   0
   ScaleHeight     =   295
   ScaleMode       =   3
   ScaleWidth      =   458
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
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
      Left            =   1440
      TabIndex        =   11
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdWebSite 
      Caption         =   "vbGo &Online"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
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
      Left            =   5280
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   -120
      ScaleHeight     =   1155
      ScaleWidth      =   7275
      TabIndex        =   4
      TabStop         =   0   
      Top             =   -150
      Width           =   7335
      Begin VB.PictureBox picLogo2 
         Appearance      =   0 
         AutoSize        =   -1
         BackColor       =   &H80000005&
         BorderStyle     =   0
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   720
         Picture         =   "frmAbout.frx":000C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         TabStop         =   0 
         Top             =   480
         Width           =   480
      End
      Begin VB.PictureBox picLogo1 
         Appearance      =   0  
         AutoSize        =   -1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   360
         Picture         =   "frmAbout.frx":08D6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         TabStop         =   0  
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1 
         BackStyle       =   0  
         Caption         =   "The VB Information Platform"
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
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   3330
      End
      Begin VB.Line lneDivider 
         BorderColor     =   &H8000000F&
         X1              =   1395
         X2              =   1395
         Y1              =   210
         Y2              =   1050
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  
         BackStyle       =   0  
         Caption         =   "vbGo VBuzz"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1725
      End
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":11A0
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   
         Italic          =   0   
         Strikethrough   =   0   
      EndProperty
      Height          =   810
      Left            =   1440
      TabIndex        =   10
      Top             =   2040
      Width           =   5295
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1
      BackStyle       =   0 
      Caption         =   "Software Version 1"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label lblAbout1 
      AutoSize        =   -1  
      BackStyle       =   0  
      Caption         =   "VBuzz Viewer && Content by vbGo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   
         Italic          =   0   
         Strikethrough   =   0   
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   3990
   End
   Begin VB.Label lblAbout2 
      AutoSize        =   -1
      BackStyle       =   0 
      Caption         =   "Copyright © 1999 - 2001 vbGo. All Rights Reserved."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0 
         Italic          =   0 
         Strikethrough   =   0 
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   5025
   End
   Begin VB.Line lneLight 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   96
      X2              =   440
      Y1              =   233
      Y2              =   233
   End
   Begin VB.Line lneDark 
      BorderColor     =   &H80000010&
      BorderStyle     =   6
      Index           =   1
      X1              =   96
      X2              =   440
      Y1              =   232
      Y2              =   232
   End
End

Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' vbGo
' VBuzz - The VB Information Platform
' Copyright © 1999 - 2001 vbGo. All Rights Reserved
' Viewer Information Screen

Option Explicit

Private Sub cmdClose_Click()
    Unload frmAbout
End Sub