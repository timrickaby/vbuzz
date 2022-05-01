VERSION 5.00
Begin VB.Form frmDownload 
   BorderStyle     =   3
   Caption         =   "vbGo VBuzz"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   ControlBox      =   0  
   ForeColor       =   &H8000000F&
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0  
   MinButton       =   0  
   ScaleHeight     =   2970
   ScaleWidth      =   6585
   ShowInTaskbar   =   0  
   StartUpPosition =   2  
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
      Left            =   4800
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdMoreLess 
      Caption         =   "&More >>"
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
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
   Begin VB.PictureBox picStatusBack 
      Height          =   255
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   5235
      TabIndex        =   4
      Top             =   1320
      Width           =   5295
      Begin VB.PictureBox picStatusFore 
         BackColor       =   &H8000000D&
         BorderStyle     =   0
         Height          =   200
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   735
         TabIndex        =   5
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0
      Caption         =   "Current Size"
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
      Left            =   360
      TabIndex        =   20
      Top             =   4560
      Width           =   2520
   End
   Begin VB.Label lblContentVersion2 
      BackStyle       =   0
      Caption         =   "??????"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3120
      TabIndex        =   19
      Top             =   3360
      Width           =   2460
   End
   Begin VB.Label Label14 
      BackStyle       =   0
      Caption         =   "??????"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3120
      TabIndex        =   18
      Top             =   5400
      Width           =   2460
   End
   Begin VB.Label Label13 
      BackStyle       =   0
      Caption         =   "??????"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3120
      TabIndex        =   17
      Top             =   5160
      Width           =   2460
   End
   Begin VB.Label lblCurrentFile2 
      BackStyle       =   0
      Caption         =   "####"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3120
      TabIndex        =   16
      Top             =   4320
      Width           =   2460
   End
   Begin VB.Label lblCurrentSize 
      BackStyle       =   0
      Caption         =   "####"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3120
      TabIndex        =   15
      Top             =   4560
      Width           =   2460
   End
   Begin VB.Label lblTotalSize2 
      BackStyle       =   0
      Caption         =   "####"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3120
      TabIndex        =   14
      Top             =   4080
      Width           =   2460
   End
   Begin VB.Label lblNumberOfFiles2 
      BackStyle       =   0
      Caption         =   "####"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3120
      TabIndex        =   13
      Top             =   3840
      Width           =   2460
   End
   Begin VB.Line lneTableHorMiddle 
      BorderColor     =   &H80000010&
      X1              =   6240
      X2              =   240
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label8 
      BackStyle       =   0
      Caption         =   "Updating File"
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
      Left            =   360
      TabIndex        =   12
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Line lneTableVertMiddle 
      BorderColor     =   &H80000010&
      X1              =   3000
      X2              =   3000
      Y1              =   3240
      Y2              =   5760
   End
   Begin VB.Label Label7 
      BackStyle       =   0
      Caption         =   "Updating Section"
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
      Left            =   360
      TabIndex        =   11
      Top             =   5160
      Width           =   2520
   End
   Begin VB.Label lblTotalSize 
      BackStyle       =   0
      Caption         =   "Total Size"
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
      Left            =   360
      TabIndex        =   10
      Top             =   4080
      Width           =   2520
   End
   Begin VB.Label lblCurrentFile 
      BackStyle       =   0
      Caption         =   "Current File"
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
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   2565
   End
   Begin VB.Label lblNumberOfFiles 
      BackStyle       =   0
      Caption         =   "Number of Files"
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
      Left            =   360
      TabIndex        =   8
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label lblContentVersion 
      BackStyle       =   0
      Caption         =   "Content Version"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   3360
      Width           =   2580
   End
   Begin VB.Line lneGray 
      BorderColor     =   &H80000010&
      BorderStyle     =   6
      Index           =   1
      X1              =   240
      X2              =   6240
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line lneWhite 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   240
      X2              =   6240
      Y1              =   3015
      Y2              =   3015
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
      TabIndex        =   6
      Top             =   1680
      Width           =   2760
   End
   Begin VB.Image imgDownload 
      Height          =   480
      Left            =   240
      Picture         =   "frmDownload.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Line lneGray 
      BorderColor     =   &H80000010&
      BorderStyle     =   6
      Index           =   0
      X1              =   240
      X2              =   6240
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line lneWhite 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   240
      X2              =   6240
      Y1              =   2055
      Y2              =   2055
   End
   Begin VB.Label lblMessageTitle 
      BackStyle       =   0
      Caption         =   "Downloading VBuzz Content..."
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
      TabIndex        =   3
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblMessageBody 
      BackStyle       =   0
      Caption         =   "Please wait while VBuzz connects and downloads the latest content for your viewer..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   5295
   End
   Begin VB.Shape shpTableTop 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   495
      Left            =   240
      Top             =   3240
      Width           =   6015
   End
   Begin VB.Shape shpTableBottom 
      BorderColor     =   &H80000010&
      Height          =   2055
      Left            =   240
      Top             =   3720
      Width           =   6015
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' vbGo
' VBuzz - The VB Information Platform
' Copyright © 1999 - 2001 vbGo. All Rights Reserved
' Content Download Screen

Option Explicit

Private Sub cmdCancel_Click()
    Unload frmDownload
End Sub

Private Sub cmdMoreLess_Click()
    If cmdMoreLess.Caption = "&More >>" Then
        cmdMoreLess.Caption = "<< &Less"
        frmDownload.Height = 6345
        frmDownload.Top = (Screen.Height / 2) - (frmDownload.Height / 2)
    
    ElseIf cmdMoreLess.Caption = "<< &Less" Then
        
        cmdMoreLess.Caption = "&More >>"
        frmDownload.Height = 3345
        frmDownload.Top = (Screen.Height / 2) - (frmDownload.Height / 2)
    End If
End Sub


