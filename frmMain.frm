VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1
   Caption         =   "#PRODUCT# - #TITLE#"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   ControlBox      =   0
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0
   MinButton       =   0
   ScaleHeight     =   445
   ScaleMode       =   3
   ScaleWidth      =   704
   StartUpPosition =   2
   Begin VB.CommandButton cmdMoreAdds 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      Height          =   900
      Left            =   9960
      TabIndex        =   30
      ToolTipText     =   "More Advertisements"
      Top             =   270
      Width           =   375
   End
   Begin VB.PictureBox picBannerAdd 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0
         Italic          =   0
         Strikethrough   =   0
      EndProperty
      Height          =   900
      Left            =   2280
      MouseIcon       =   "frmMain.frx":08CA
      MousePointer    =   99
      Picture         =   "frmMain.frx":0A1C
      ScaleHeight     =   56
      ScaleMode       =   3
      ScaleWidth      =   496
      TabIndex        =   5
      ToolTipText     =   "Visit [...]"
      Top             =   270
      Width           =   7500
   End
   Begin VB.PictureBox picSideBar 
      BackColor       =   &H80000010&
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6795
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   -120
      Width           =   1935
      Begin VB.OptionButton optAbout 
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         Height          =   315
         Left            =   1200
         Style           =   1
         TabIndex        =   60
         ToolTipText     =   "About"
         Top             =   6240
         Width           =   375
      End
      Begin VB.OptionButton optClose 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         Height          =   315
         Left            =   240
         Style           =   1
         TabIndex        =   59
         ToolTipText     =   "Close"
         Top             =   6240
         Width           =   375
      End
      Begin VB.OptionButton optMinimise 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         Height          =   315
         Left            =   720
         Style           =   1
         TabIndex        =   58
         ToolTipText     =   "Minimise"
         Top             =   6240
         Width           =   375
      End
      Begin VB.OptionButton optvbGoHome 
         Caption         =   "vbGo &Home"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         Height          =   915
         Left            =   240
         Picture         =   "frmMain.frx":CBD2
         Style           =   1
         TabIndex        =   39
         ToolTipText     =   "vbGo Home"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optLinks 
         Caption         =   "&Web Links"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         Height          =   915
         Left            =   240
         Picture         =   "frmMain.frx":D49C
         Style           =   1
         TabIndex        =   31
         ToolTipText     =   "Browse The Links List"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.OptionButton optProducts 
         Caption         =   "&Products"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         Height          =   915
         Left            =   240
         Picture         =   "frmMain.frx":DD66
         Style           =   1
         TabIndex        =   3
         ToolTipText     =   "Browse The Products List"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.OptionButton optNews 
         Caption         =   "&News"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         Height          =   915
         Left            =   240
         Picture         =   "frmMain.frx":E630
         Style           =   1
         TabIndex        =   2
         ToolTipText     =   "Current News"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.OptionButton optHome 
         Caption         =   "&VBuzz Home"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         Height          =   915
         Left            =   240
         Picture         =   "frmMain.frx":EEFA
         Style           =   1
         TabIndex        =   1
         ToolTipText     =   "vbGo VBuzz Home"
         Top             =   360
         Value           =   -1
         Width           =   1335
      End
      Begin VB.Line lneWhite 
         BorderColor     =   &H80000016&
         Index           =   0
         X1              =   225
         X2              =   1545
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Line lneGray 
         BorderColor     =   &H80000015&
         BorderStyle     =   6
         Index           =   1
         X1              =   240
         X2              =   1560
         Y1              =   2535
         Y2              =   2535
      End
      Begin VB.Line lneGray 
         BorderColor     =   &H80000015&
         BorderStyle     =   6
         Index           =   3
         X1              =   240
         X2              =   1560
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line lneWhite 
         BorderColor     =   &H80000016&
         Index           =   3
         X1              =   240
         X2              =   1560
         Y1              =   6015
         Y2              =   6015
      End
      Begin VB.Label Label1 
         Alignment       =   2
         BackStyle       =   0
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3840
         Width           =   1455
      End
   End
   Begin VB.PictureBox picVBuzzHomeWrapper 
      BackColor       =   &H80000005&
      Height          =   4575
      Left            =   2280
      ScaleHeight     =   4515
      ScaleWidth      =   7995
      TabIndex        =   14
      Top             =   1920
      Width           =   8055
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options..."
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
         Left            =   720
         TabIndex        =   61
         ToolTipText     =   "Viewer Options"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect && Retrieve"
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
         Left            =   5880
         TabIndex        =   38
         ToolTipText     =   "Connect & Retrieve The Latest Content"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update VBuzz"
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
         Left            =   3720
         TabIndex        =   57
         ToolTipText     =   "Update Your Viewer"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   720
         X2              =   7800
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label lblInfoBody 
         BackStyle       =   0
         Caption         =   ""
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
         Height          =   435
         Index           =   1
         Left            =   1320
         TabIndex        =   21
         Top             =   2280
         Width           =   6450
      End
      Begin VB.Label lblInfoTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "Find What You Need - Fast:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   20
         Top             =   2040
         Width           =   2640
      End
      Begin VB.Label lblInfoBody 
         BackStyle       =   0
         Caption         =   "Stay up to date with the latest developments at vbGo and its partner sites."
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
         Height          =   435
         Index           =   0
         Left            =   1320
         TabIndex        =   19
         Top             =   1320
         Width           =   6450
      End
      Begin VB.Label lblInfoTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "Keep Up To Date:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   18
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Image imgVBuzzHome 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":F7C4
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblHomeTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "vbGo VBuzz Home"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   120
         Width           =   1740
      End
      Begin VB.Label lblHomeSubTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "Your VB Information Platform"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   15
         Top             =   360
         Width           =   3330
      End
      Begin VB.Label lblHomeTitleBack 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000017&
         Height          =   735
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.PictureBox picvbGoHomeWrapper 
      BackColor       =   &H80000005&
      Height          =   4575
      Left            =   2280
      ScaleHeight     =   4515
      ScaleWidth      =   7995
      TabIndex        =   40
      Top             =   1920
      Width           =   8055
      Begin VB.VScrollBar scrvbGoHomeScroller 
         Height          =   3615
         Left            =   7680
         TabIndex        =   45
         Top             =   840
         Width           =   210
      End
      Begin VB.PictureBox picvbGoHomeScroller 
         BackColor       =   &H80000005&
         BorderStyle     =   0
         Height          =   3615
         Left            =   120
         ScaleHeight     =   3615
         ScaleWidth      =   7455
         TabIndex        =   46
         Top             =   840
         Width           =   7455
         Begin VB.Label lblvbGoHomeItemBody 
            BackStyle       =   0
            Caption         =   "Currently featured on the vbGo home page is the MacroWeaver product."
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
            Height          =   915
            Index           =   0
            Left            =   1080
            TabIndex        =   48
            Top             =   1560
            Width           =   6210
         End
         Begin VB.Label lblvbGoHomeItemTitle 
            AutoSize        =   -1
            BackStyle       =   0
            Caption         =   "Currently Featured:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0
               Italic          =   0
               Strikethrough   =   0
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   47
            Top             =   1320
            Width           =   1920
         End
         Begin VB.Image imgvbGoHomeItem 
            Height          =   1050
            Index           =   0
            Left            =   720
            Picture         =   "frmMain.frx":1008E
            Top             =   240
            Width           =   2850
         End
      End
      Begin VB.Label lblvbGoHomeSubTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "The VB community delivered directly to your desktop."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   42
         Top             =   360
         Width           =   4605
      End
      Begin VB.Label lblvbGoHomeTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "vbGo Home"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   41
         Top             =   120
         Width           =   1110
      End
      Begin VB.Image imgvbGoHome 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":11C6E
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblvbGoHome 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000017&
         Height          =   735
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.PictureBox picWebLinksWrapper 
      BackColor       =   &H80000005&
      Height          =   4575
      Left            =   2280
      ScaleHeight     =   4515
      ScaleWidth      =   7995
      TabIndex        =   32
      Top             =   1920
      Visible         =   0
      Width           =   8055
      Begin VB.Label lblInfoBody 
         BackStyle       =   0
         Caption         =   "Browse through the list of vbGo created products ready to add to your VB toolbox."
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
         Height          =   675
         Index           =   4
         Left            =   1560
         TabIndex        =   36
         Top             =   1320
         Width           =   6210
      End
      Begin VB.Label lblInfoTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "Abstract Visual Basic"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   1560
         TabIndex        =   35
         Top             =   1080
         Width           =   2040
      End
      Begin VB.Image imgWebLinks 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":12538
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblWebLinksTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "VB Web Links"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   34
         Top             =   120
         Width           =   1290
      End
      Begin VB.Label lblWebLinksSubTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "Links to some great VB code and information web sites"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   33
         Top             =   360
         Width           =   4725
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   840
         Picture         =   "frmMain.frx":12E02
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblWebLinksBack 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000017&
         Height          =   735
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.PictureBox picvbGoProductsWrapper 
      BackColor       =   &H80000005&
      Height          =   4575
      Left            =   2280
      ScaleHeight     =   4515
      ScaleWidth      =   7995
      TabIndex        =   49
      Top             =   1920
      Width           =   8055
      Begin VB.VScrollBar scrvbGoProducts 
         Height          =   3615
         Left            =   7680
         TabIndex        =   56
         Top             =   840
         Width           =   210
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         BorderStyle     =   0
         Height          =   3615
         Left            =   120
         ScaleHeight     =   3615
         ScaleWidth      =   7455
         TabIndex        =   53
         Top             =   840
         Width           =   7455
         Begin VB.Image imgvbGoProducts 
            Height          =   1050
            Index           =   0
            Left            =   720
            Picture         =   "frmMain.frx":139C4
            Top             =   240
            Width           =   2850
         End
         Begin VB.Label lblvbGoProductsItemTitle 
            AutoSize        =   -1
            BackStyle       =   0
            Caption         =   "Currently Featured:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0
               Italic          =   0
               Strikethrough   =   0
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   55
            Top             =   1320
            Width           =   1920
         End
         Begin VB.Label lblvbGoProductsItemBody 
            BackStyle       =   0
            Caption         =   "Currently featured on the vbGo home page is the MacroWeaver product."
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
            Height          =   915
            Index           =   0
            Left            =   1080
            TabIndex        =   54
            Top             =   1560
            Width           =   6210
         End
      End
      Begin VB.Image imgvbGoProducts 
         Height          =   480
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":155A4
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblvbGoProductsTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "vbGo Products List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   51
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblvbGoProductsSubTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "Advanced software && component products from vbGo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   50
         Top             =   360
         Width           =   4590
      End
      Begin VB.Label lblvbGoProductsBack 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000017&
         Height          =   735
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.PictureBox picProductsWrapper 
      BackColor       =   &H80000005&
      Height          =   4575
      Left            =   2280
      ScaleHeight     =   4515
      ScaleWidth      =   7995
      TabIndex        =   22
      Top             =   1920
      Visible         =   0
      Width           =   8055
      Begin VB.Image imgProductsItem 
         Height          =   480
         Index           =   1
         Left            =   840
         MouseIcon       =   "frmMain.frx":15E6E
         MousePointer    =   99
         Picture         =   "frmMain.frx":15FC0
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image imgProductsItem 
         Height          =   480
         Index           =   0
         Left            =   840
         MouseIcon       =   "frmMain.frx":16B82
         MousePointer    =   99
         Picture         =   "frmMain.frx":16CD4
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblProductsSubTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "Your VB Information Platform"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   28
         Top             =   360
         Width           =   3330
      End
      Begin VB.Label lblProductsTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "vbGo && Third Party Products"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   27
         Top             =   120
         Width           =   2775
      End
      Begin VB.Image imgProducts 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":17896
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblProductsItemTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "vbGo Products List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmMain.frx":18160
         MousePointer    =   99
         TabIndex        =   26
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblProductsItemBody 
         BackStyle       =   0
         Caption         =   "Browse through the list of vbGo created products ready to add to your VB toolbox."
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
         Height          =   435
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmMain.frx":182B2
         MousePointer    =   99
         TabIndex        =   25
         Top             =   1320
         Width           =   6210
      End
      Begin VB.Label lblProductsItemTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "Other Third Party Products List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1560
         MouseIcon       =   "frmMain.frx":18404
         MousePointer    =   99
         TabIndex        =   24
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label lblProductsItemBody 
         BackStyle       =   0
         Caption         =   "View the best commercial products on the market."
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
         Height          =   435
         Index           =   1
         Left            =   1560
         MouseIcon       =   "frmMain.frx":18556
         MousePointer    =   99
         TabIndex        =   23
         Top             =   2280
         Width           =   6210
      End
      Begin VB.Label lblProductsTitleBack 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000017&
         Height          =   735
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.PictureBox picNewsWrapper 
      BackColor       =   &H80000005&
      Height          =   4575
      Left            =   2280
      ScaleHeight     =   4515
      ScaleWidth      =   7995
      TabIndex        =   6
      Top             =   1920
      Visible         =   0
      Width           =   8055
      Begin VB.VScrollBar scrNewsScroller 
         Height          =   3615
         Left            =   7680
         TabIndex        =   11
         Top             =   840
         Width           =   210
      End
      Begin VB.PictureBox picNewsScroller 
         BackColor       =   &H80000005&
         BorderStyle     =   0
         Height          =   3615
         Left            =   120
         ScaleHeight     =   3615
         ScaleWidth      =   7455
         TabIndex        =   10
         Top             =   840
         Width           =   7455
         Begin VB.Label lblNewsItemBody 
            BackStyle       =   0
            Caption         =   "News Item Body"
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
            Height          =   915
            Index           =   1
            Left            =   1200
            TabIndex        =   13
            Top             =   480
            Width           =   6075
         End
         Begin VB.Label lblNewsItemTitle 
            AutoSize        =   -1
            BackStyle       =   0
            Caption         =   "News Item Title"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0
               Italic          =   0
               Strikethrough   =   0
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   12
            Top             =   240
            Width           =   1530
         End
      End
      Begin VB.Label lblNewsSubTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "Discover whats happening at vbGo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   2985
      End
      Begin VB.Label lblNewsTitle 
         AutoSize        =   -1
         BackStyle       =   0
         Caption         =   "vbGo News"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0
            Italic          =   0
            Strikethrough   =   0
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   195
         Left            =   840
         TabIndex        =   7
         Top             =   120
         Width           =   1065
      End
      Begin VB.Image imgNewsLogo 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":186A8
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblNewsTitleBack 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000017&
         Height          =   735
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.Label lblCurrentlyViewing 
      AutoSize        =   -1
      BackStyle       =   0
      Caption         =   "Currently Viewing: 'VBuzz Home'"
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
      Left            =   2280
      TabIndex        =   44
      Top             =   1560
      Width           =   2835
   End
   Begin VB.Line lneGray 
      BorderColor     =   &H80000010&
      BorderStyle     =   6
      Index           =   0
      X1              =   152
      X2              =   688
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line lneWhite 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   152
      X2              =   688
      Y1              =   97
      Y2              =   97
   End
End

Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' vbGo
' VBuzz - The VB Information Platform
' Copyright © 1999 - 2001 vbGo. All Rights Reserved

Option Explicit


Public Sub cmdConnect_Click()
    Call mdlUpdateContent.CheckForUpdate
End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show vbModal
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Dim msgAnswer As VbMsgBoxResult
    
    msgAnswer = MsgBox("Are you sure you want to close VBuzz?", vbQuestion + vbYesNo, "vbGo VBuzz")
    
    If msgAnswer = vbYes Then
        End: Exit Sub
    
    ElseIf msgAnswer = vbNo Then
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub optClose_Click()
    optClose.Value = False
    frmMain.picSideBar.SetFocus
    Unload frmMain
End Sub

Private Sub optMinimise_Click()
    optMinimise.Value = False
    frmMain.picSideBar.SetFocus
    frmMain.WindowState = 1
End Sub

Private Sub optAbout_Click()
    optAbout.Value = False
    frmMain.picSideBar.SetFocus
    frmAbout.Show vbModal
End Sub

Private Sub optHome_Click()
    lblCurrentlyViewing.Caption = "Currently Viewing: 'VBuzz Home'"
    frmMain.picSideBar.SetFocus
    frmMain.picVBuzzHomeWrapper.Visible = True
    frmMain.picVBuzzHomeWrapper.ZOrder 0
End Sub

Private Sub optvbGoHome_Click()
    lblCurrentlyViewing.Caption = "Currently Viewing: 'vbGo Home'"
    frmMain.picSideBar.SetFocus
    frmMain.picvbGoHomeWrapper.Visible = True
    frmMain.picvbGoHomeWrapper.ZOrder 0
End Sub

Private Sub optNews_Click()
    lblCurrentlyViewing.Caption = "Currently Viewing: 'News'"
    frmMain.picSideBar.SetFocus
    frmMain.picNewsWrapper.Visible = True
    frmMain.picNewsWrapper.ZOrder 0
End Sub

Private Sub optLinks_Click()
    lblCurrentlyViewing.Caption = "Currently Viewing: 'Web Links'"
    frmMain.picSideBar.SetFocus
    frmMain.picWebLinksWrapper.Visible = True
    frmMain.picWebLinksWrapper.ZOrder 0
End Sub

Private Sub optProducts_Click()
    lblCurrentlyViewing.Caption = "Currently Viewing: 'Products'"
    frmMain.picSideBar.SetFocus
    frmMain.picProductsWrapper.Visible = True
    frmMain.picProductsWrapper.ZOrder 0
End Sub