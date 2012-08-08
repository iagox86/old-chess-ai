VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "74.319 Project Client (Chess)"
   ClientHeight    =   5355
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lstMoves 
      BackColor       =   &H8000000F&
      Height          =   3495
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   5280
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame fraImages 
      Caption         =   "Invisible"
      DragMode        =   1  'Automatic
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   32
         Left            =   5040
         Picture         =   "frmMain.frx":0000
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   31
         Left            =   4200
         Picture         =   "frmMain.frx":0BD2
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   30
         Left            =   3480
         Picture         =   "frmMain.frx":1854
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   29
         Left            =   2640
         Picture         =   "frmMain.frx":2426
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   28
         Left            =   1800
         Picture         =   "frmMain.frx":30A8
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   27
         Left            =   960
         Picture         =   "frmMain.frx":3D2A
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   26
         Left            =   4920
         Picture         =   "frmMain.frx":49AC
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   25
         Left            =   4200
         Picture         =   "frmMain.frx":562E
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   24
         Left            =   3360
         Picture         =   "frmMain.frx":6370
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   23
         Left            =   2640
         Picture         =   "frmMain.frx":6FF2
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   22
         Left            =   1800
         Picture         =   "frmMain.frx":7D34
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   21
         Left            =   960
         Picture         =   "frmMain.frx":8A76
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   20
         Left            =   120
         Picture         =   "frmMain.frx":97B8
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   705
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":99B2
         Stretch         =   -1  'True
         Top             =   240
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   660
         Index           =   12
         Left            =   4920
         Picture         =   "frmMain.frx":9C4C
         Top             =   1080
         Width           =   645
      End
      Begin VB.Image ChessPiece 
         Height          =   660
         Index           =   11
         Left            =   4080
         Picture         =   "frmMain.frx":A81E
         Top             =   1080
         Width           =   720
      End
      Begin VB.Image ChessPiece 
         Height          =   660
         Index           =   10
         Left            =   3360
         Picture         =   "frmMain.frx":B4A0
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image ChessPiece 
         Height          =   660
         Index           =   9
         Left            =   2520
         Picture         =   "frmMain.frx":C072
         Top             =   1080
         Width           =   720
      End
      Begin VB.Image ChessPiece 
         Height          =   660
         Index           =   8
         Left            =   1800
         Picture         =   "frmMain.frx":CCF4
         Top             =   1080
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   660
         Index           =   7
         Left            =   960
         Picture         =   "frmMain.frx":D976
         Top             =   1080
         Width           =   705
      End
      Begin VB.Image ChessPiece 
         Height          =   720
         Index           =   6
         Left            =   4920
         Picture         =   "frmMain.frx":E5F8
         Top             =   240
         Width           =   645
      End
      Begin VB.Image ChessPiece 
         Height          =   720
         Index           =   5
         Left            =   4080
         Picture         =   "frmMain.frx":F27A
         Top             =   240
         Width           =   720
      End
      Begin VB.Image ChessPiece 
         Height          =   720
         Index           =   4
         Left            =   3360
         Picture         =   "frmMain.frx":FFBC
         Top             =   240
         Width           =   615
      End
      Begin VB.Image ChessPiece 
         Height          =   720
         Index           =   3
         Left            =   2520
         Picture         =   "frmMain.frx":10186
         Top             =   240
         Width           =   720
      End
      Begin VB.Image ChessPiece 
         Height          =   720
         Index           =   2
         Left            =   1800
         Picture         =   "frmMain.frx":10350
         Top             =   240
         Width           =   675
      End
      Begin VB.Image ChessPiece 
         Height          =   720
         Index           =   1
         Left            =   960
         Picture         =   "frmMain.frx":1051A
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Label Label18 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label17 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   20
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label16 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   19
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label15 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label14 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   17
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label13 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   16
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label12 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label11 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   14
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label10 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "Enter move manually:"
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Move codes:"
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   63
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   62
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   61
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   60
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   59
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   58
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   57
      Left            =   960
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   56
      Left            =   360
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   55
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   54
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   53
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   52
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   51
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   50
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   49
      Left            =   960
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   48
      Left            =   360
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   47
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   46
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   45
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   44
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   43
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   42
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   41
      Left            =   960
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   40
      Left            =   360
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   39
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   38
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   37
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   36
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   35
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   34
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   33
      Left            =   960
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   32
      Left            =   360
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   31
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   30
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   29
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   28
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   27
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   26
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   25
      Left            =   960
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   24
      Left            =   360
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   23
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   22
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   21
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   20
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   19
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   18
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   17
      Left            =   960
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   16
      Left            =   360
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   15
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   14
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   13
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   12
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   11
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   10
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   9
      Left            =   960
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   8
      Left            =   360
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   7
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   6
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   5
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   4
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   3
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   2
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   1
      Left            =   960
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Place 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   0
      Left            =   360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewgame 
         Caption         =   "&New game"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' White always goes first, therefore white is player 1
' Shah-Mat (check mate) is Persian for "The king is dead"

' Keeping track:
' Each move is 2 or 4 characters
' The code for the piece
' blank = pawn
' N = kNight
' K = King
' Q = Queen
' R = Rook
' B = Bishop
' The destination file/column, from a to h (left to right)
' The destination rank/row, from 7 - 0 (bottem left is a1)
' An "x" at the end if a piece is captured

Option Explicit

Private Function GetMoveString(irow As Integer, icol As Integer, frow As Integer, fcol As Integer) As String
    GetMoveString = irow & " " & icol & " " & frow & " " & fcol
    
        
    
End Function

Private Sub cmdHelp_Click()
    MsgBox "This program is designed to go with Ron Bowes's chess program for 74.319 final project" & vbCrLf & _
           "It takes the inputs that the program gives in the format ""row1 col1 row2 col2"" and outputs the same way", vbInformation
End Sub

Private Sub cmdOK_Click()
    Dim Data() As String
    Dim iIndex As Integer
    
    Data = Split(txtCode.Text, " ", 4)
    Dim SRC As Integer
    SRC = GetLocation(Val(Data(0)), Val(Data(1)))
    Dim Dest As Integer
    Dest = GetLocation(Val(Data(2)), Val(Data(3)))
    
    lstMoves.Text = txtCode.Text & vbCrLf & lstMoves.Text
    
    Place(Dest).Tag = Place(SRC).Tag
    Place(SRC).Tag = Blank
        
    movearray Lastboard, Lastlastboard
    movearray Gameboard, Lastboard
        
    Gameboard(GetRow(Dest), GetColumn(Dest)) = Gameboard(GetRow(SRC), GetColumn(SRC))
    Gameboard(GetRow(SRC), GetColumn(SRC)) = Blank
    DrawBoard
    
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)
    
End Sub

Private Sub Form_Load()
    SetDefaultBoard
    DrawBoard
End Sub

Private Sub mnuEditUndo_Click()
    Dim iIndex As Integer
    
    For iIndex = 0 To Place.UBound
        Place(iIndex).Tag = Gameboard(GetRow(iIndex), GetColumn(iIndex))
    Next
    
    movearray Lastboard, Gameboard
    movearray Lastlastboard, Lastboard
    
    DrawBoard
End Sub

Private Sub mnuFileNewgame_Click()
    If (MsgBox("Are you sure you want to reset the current game?", vbQuestion + vbYesNo, "New game") = vbYes) Then
        SetDefaultBoard
        DrawBoard
    End If
End Sub

Private Sub DrawBoard()
    ' This function positions the pieces according to CB
    ' it should be called after every move
    
    Dim row As Integer
    Dim column As Integer
    Dim NewImage As Image
    
  
    For row = 0 To 7
        For column = 0 To 7
            If ((row + column) Mod 2) Then
                If Place(column + (8 * row)).Picture <> ChessPiece(Gameboard(row, column) + 20).Picture Then
                    Place(column + (8 * row)).Picture = ChessPiece(Gameboard(row, column) + 20).Picture
                    If Gameboard(row, column) <> Blank Then
                        Place(column + (8 * row)).DragMode = 1
                    Else
                        Place(column + (8 * row)).DragMode = 0
                    End If
                End If
            Else
                If Place(column + (8 * row)).Picture <> ChessPiece(Gameboard(row, column)).Picture Then
                    Place(column + (8 * row)).Picture = ChessPiece(Gameboard(row, column)).Picture
                    If Gameboard(row, column) <> Blank Then
                        Place(column + (8 * row)).DragMode = 1
                    Else
                        Place(column + (8 * row)).DragMode = 0
                    End If
                End If
            End If
        Next
    Next
End Sub


Private Sub SetDefaultBoard()
    Dim row As Integer
    Dim column As Integer
    For row = 0 To 7
        For column = 0 To 7
            Gameboard(row, column) = EmptyBoard(row, column)
            Place(column + (8 * row)).Tag = Gameboard(row, column)
            Lastboard(row, column) = EmptyBoard(row, column)
            Lastlastboard(row, column) = EmptyBoard(row, column)
        Next
    Next
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "This is a very simple chess program written by Ron Bowes" & vbCrLf & vbCrLf & "It doesn't check player, or moves, or inputs, or anything else at all." & "In fact, the only difference between this and a real chess board are two things:" & vbCrLf & "This one has 2 levels of undo, and this one keeps a list of all the moves both players have made."
End Sub

Private Sub Place_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Dim iIndex As Integer
    If (IsLegalmove(Source, Place(Index))) Then
        Place(Index).Tag = Source.Tag
        Source.Tag = Blank
        
        lstMoves.Text = GetMoveString(GetRow(Source.Index) + 1, GetColumn(Source.Index) + 1, GetRow(Index) + 1, GetColumn(Index) + 1) & vbCrLf & lstMoves.Text
        
        movearray Lastboard, Lastlastboard
        movearray Gameboard, Lastboard
        
        
        Gameboard(GetRow(Index), GetColumn(Index)) = Gameboard(GetRow(Source.Index), GetColumn(Source.Index))
        Gameboard(GetRow(Source.Index), GetColumn(Source.Index)) = Blank
        DrawBoard
    End If
End Sub

Private Function IsLegalmove(Source As Control, Dest As Control) As Boolean
    If Source = Dest Then
        IsLegalmove = False
    Else
        If (Source.Picture <> ChessPiece(Blank).Picture And Source.Picture <> ChessPiece(Blank + 20).Picture) Then
            IsLegalmove = True
        End If
    End If
End Function

Sub movearray(array1() As Integer, array2() As Integer)
    Dim Rows As Integer
    Dim Columns As Integer
    
    For Rows = 0 To 7
        For Columns = 0 To 7
            array2(Rows, Columns) = array1(Rows, Columns)
        Next
    Next
End Sub
