VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BLACK Plenty V 4.0"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox p2 
      BackColor       =   &H0000C0C0&
      Height          =   255
      Left            =   9240
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   54
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox p1 
      BackColor       =   &H00808000&
      Height          =   195
      Left            =   9180
      ScaleHeight     =   135
      ScaleWidth      =   375
      TabIndex        =   53
      Top             =   1320
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8700
      Width           =   1815
   End
   Begin VB.PictureBox picOpen 
      BackColor       =   &H00000000&
      Height          =   3975
      Left            =   12060
      ScaleHeight     =   3915
      ScaleWidth      =   9375
      TabIndex        =   40
      Top             =   660
      Width           =   9435
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   120
         TabIndex        =   49
         Top             =   660
         Width           =   6915
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   180
         Width           =   7455
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Törlés"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   8400
         TabIndex        =   50
         Top             =   3480
         Width           =   690
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Kategória"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   300
         Left            =   360
         TabIndex        =   41
         Top             =   180
         Width           =   1170
      End
   End
   Begin VB.PictureBox picUj 
      BackColor       =   &H00000000&
      Height          =   3915
      Left            =   12240
      ScaleHeight     =   3855
      ScaleWidth      =   9375
      TabIndex        =   28
      Top             =   2460
      Width           =   9435
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2640
         TabIndex        =   52
         Top             =   2340
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1500
         TabIndex        =   37
         Top             =   1680
         Width           =   7695
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   6180
         TabIndex        =   36
         Text            =   "0"
         Top             =   1140
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1860
         TabIndex        =   35
         Text            =   "0"
         Top             =   1140
         Width           =   3075
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   540
         Width           =   7635
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Hónap, ha nem a jelen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   600
         TabIndex        =   51
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Mégsem"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   285
         Left            =   6120
         TabIndex        =   39
         Top             =   2460
         Width           =   960
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   285
         Left            =   7740
         TabIndex        =   38
         Top             =   2460
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Leirás"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   540
         TabIndex        =   34
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "vagy EURO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   5100
         TabIndex        =   33
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Összeg - LEJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   540
         TabIndex        =   32
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Új hozzáadása"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   12
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   60
         Width           =   8955
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Kategória"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   540
         TabIndex        =   29
         Top             =   600
         Width           =   825
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H00000000&
      Height          =   6675
      Left            =   0
      ScaleHeight     =   6615
      ScaleWidth      =   11775
      TabIndex        =   16
      Top             =   1320
      Width           =   11835
      Begin MSChart20Lib.MSChart MSC 
         Height          =   4515
         Left            =   120
         OleObjectBlob   =   "frmMain.frx":030A
         TabIndex        =   27
         ToolTipText     =   "Click - nagyitás - kicsinyités"
         Top             =   720
         Width           =   8115
      End
      Begin VB.Label Label25 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   180
         TabIndex        =   48
         Top             =   6240
         Width           =   3315
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10800
         TabIndex        =   47
         Top             =   5760
         Width           =   120
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10800
         TabIndex        =   46
         Top             =   5400
         Width           =   120
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10800
         TabIndex        =   45
         Top             =   5040
         Width           =   120
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10800
         TabIndex        =   44
         Top             =   4680
         Width           =   120
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10800
         TabIndex        =   43
         Top             =   4320
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "ÉVI-HAVI"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   9780
         TabIndex        =   26
         Top             =   6180
         Width           =   1260
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "KÖVETKEZÖ"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   7860
         TabIndex        =   25
         Top             =   6180
         Width           =   1485
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "ELÖZÖ"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   6480
         TabIndex        =   24
         Top             =   6180
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "MA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   5640
         TabIndex        =   23
         Top             =   6180
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nekem tartoznak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   9000
         TabIndex        =   22
         Top             =   5760
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Adósság"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   9780
         TabIndex        =   21
         Top             =   5400
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Differencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   9540
         TabIndex        =   20
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Össz kimenet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   9360
         TabIndex        =   19
         Top             =   4680
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Össz bejövet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   9360
         TabIndex        =   18
         Top             =   4320
         Width           =   1200
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "ma, 13/07/2003 - pénzügyi infó"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   14.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Width           =   9075
      End
   End
   Begin VB.PictureBox Picture15 
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   15780
      Picture         =   "frmMain.frx":282F
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   14
      Top             =   7440
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox Picture14 
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   13920
      Picture         =   "frmMain.frx":349B
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   13
      Top             =   7440
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox Picture13 
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   12060
      Picture         =   "frmMain.frx":4118
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   12
      Top             =   7380
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox Picture12 
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   10140
      Picture         =   "frmMain.frx":4D99
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   11
      Top             =   7320
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox Picture11 
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   7980
      Picture         =   "frmMain.frx":561D
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   10
      Top             =   7320
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox Picture10 
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   15780
      Picture         =   "frmMain.frx":5FDB
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   9
      Top             =   8700
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox Picture9 
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   13860
      Picture         =   "frmMain.frx":6DE3
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   8
      Top             =   8640
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox Picture8 
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   11940
      Picture         =   "frmMain.frx":79CC
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   7
      Top             =   8640
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   10020
      Picture         =   "frmMain.frx":85CB
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   6
      Top             =   8640
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox Picture6 
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   8100
      Picture         =   "frmMain.frx":9063
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   5
      Top             =   8640
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   10080
      Picture         =   "frmMain.frx":9C08
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   4
      Top             =   0
      Width           =   1830
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   7740
      Picture         =   "frmMain.frx":A874
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   1830
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   4920
      Picture         =   "frmMain.frx":B4F1
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   1830
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2460
      Picture         =   "frmMain.frx":C172
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1830
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   60
      Picture         =   "frmMain.frx":C9F6
      ScaleHeight     =   900
      ScaleWidth      =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   1830
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   11475
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Click()
If Combo2.Text = Combo2.List(0) Then
Data1.RecordSource = "T1"
Data1.Refresh
ElseIf Combo2.Text = Combo2.List(1) Then
Data1.RecordSource = "T2"
Data1.Refresh
ElseIf Combo2.Text = Combo2.List(2) Then
Data1.RecordSource = "T3"
Data1.Refresh
ElseIf Combo2.Text = Combo2.List(3) Then
Data1.RecordSource = "T4"
Data1.Refresh
End If

On Error Resume Next
List1.Clear

Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
List1.AddItem Data1.Recordset(0).Value & " - " & Trim(Data1.Recordset(2).Value) & "  (" & Trim(Data1.Recordset(1).Value) & ")"
Data1.Recordset.MoveNext
Loop


End Sub

Private Sub Form_Load()
picOpen.Left = picInfo.Left
picUj.Left = picInfo.Left
picOpen.Top = picInfo.Top
picUj.Top = picInfo.Top
picInfo.Visible = True
picUj.Visible = False
picOpen.Visible = False
ii = 0
EVI = False
Text4.Text = Format(Date, "mm/yyyy")
picUj.Height = picInfo.Height
picUj.Width = picInfo.Width
picOpen.Height = picInfo.Height
picOpen.Width = picInfo.Width

MSC.Left = 120
MSC.Top = 660
MSC.Width = 8715
MSC.Height = 5535



Dim eu As String
eu = GetSetting(App.Title, "set", "eu")
If eu = "" Then eu = "37800"
EUF = Val(eu)
Label25.Caption = "Beállitott euró árfolyam: " & EUF
CRDB
Data1.DatabaseName = App.Path & "\data.h"
Data1.RecordSource = "T1"
Data1.Refresh

MakeInfo Format(Date, "mm/yyyy")

Label2.Caption = "Ma, " & Format(Date, "dd/mm/yyyy") & " - pénzügyi infó - EURO"

Combo1.AddItem "Bejövet"
Combo1.AddItem "Kiadás"
Combo1.AddItem "Adósság"
Combo1.AddItem "Nekem tartoznak"
Combo1.Text = Combo1.List(0)
Combo2.AddItem "Bejövet"
Combo2.AddItem "Kiadás"
Combo2.AddItem "Adósság"
Combo2.AddItem "Nekem tartoznak"
Combo2.Text = Combo2.List(0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Picture = Picture11.Picture
Picture2.Picture = Picture12.Picture
Picture3.Picture = Picture13.Picture
Picture4.Picture = Picture14.Picture
Picture5.Picture = Picture15.Picture
Label1.Caption = ""
End Sub

Private Sub Label10_Click()
ii = ii + 1
Dim dData As Date
If EVI = False Then
dData = DateAdd("m", ii, Date)
Label2.Caption = Format(dData, "dd/mm/yyyy") & " - pénzügyi infó"
MakeInfo Format(dData, "mm/yyyy")
Else
dData = DateAdd("yyyy", ii, Date)
Label2.Caption = Format(dData, "yyyy") & " - pénzügyi infó"
MakeEvi Format(dData, "yyyy")

End If
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = p1.BackColor

Label9.ForeColor = p1.BackColor
Label10.ForeColor = p2.BackColor
Label11.ForeColor = p1.BackColor
Label1.Caption = "Következõ hónap vagy év"
End Sub

Private Sub Label11_Click()
If EVI = False Then
MakeEvi Format(Date, "yyyy")
Label2.Caption = "Évi - " & Format(Date, "yyyy") & " - pénzügyi infó - EURO"

EVI = True
Else
MakeInfo Format(Date, "mm/yyyy")
Label2.Caption = "Ma, " & Format(Date, "dd/mm/yyyy") & " - pénzügyi infó - EURO"

EVI = False
End If
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = p1.BackColor

Label9.ForeColor = p1.BackColor
Label10.ForeColor = p1.BackColor
Label11.ForeColor = p2.BackColor
Label1.Caption = "Váltás évi-havi nézetek között"
End Sub

Private Sub Label17_Click()
If Combo1.Text = Combo1.List(0) Then
Data1.RecordSource = "T1"
Data1.Refresh
ElseIf Combo1.Text = Combo1.List(1) Then
Data1.RecordSource = "T2"
Data1.Refresh
ElseIf Combo1.Text = Combo1.List(2) Then
Data1.RecordSource = "T3"
Data1.Refresh
ElseIf Combo1.Text = Combo1.List(3) Then
Data1.RecordSource = "T4"
Data1.Refresh
End If
Dim eu As Single
eu = Val(Text2.Text)
Data1.Recordset.AddNew
Data1.Recordset(0).Value = Text4.Text
Data1.Recordset(1).Value = Text3.Text
Data1.Recordset(2).Value = eu
Data1.Recordset.Update

Combo1.Text = Combo1.List(0)
Text1.Text = "0"
Text2.Text = "0"
Text3.Text = ""

Picture1_Click
End Sub

Private Sub Label18_Click()
Combo1.Text = Combo1.List(0)
Text1.Text = "0"
Text2.Text = "0"
Text3.Text = ""

Picture1_Click
End Sub

Private Sub Label26_Click()
If List1.Text <> "" Then
Dim aaa
aaa = MsgBox("Tényleg törli?", vbQuestion + vbYesNo, "Figyelem!")
If aaa = vbYes Then
On Error Resume Next
Data1.Recordset.MoveFirst
Data1.Recordset.Move List1.ListIndex
Data1.Recordset.Delete
List1.RemoveItem List1.ListIndex
End If
End If
End Sub

Private Sub Label8_Click()
If EVI = False Then
Label2.Caption = "Ma, " & Format(Date, "dd/mm/yyyy") & " - pénzügyi infó"
MakeInfo Format(Date, "mm/yyyy")
ii = 0
Else
MakeEvi Format(Date, "yyyy")
Label2.Caption = "Évi - " & Format(Date, "yyyy") & " - pénzügyi infó - EURO"
End If
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = p2.BackColor

Label9.ForeColor = p1.BackColor
Label10.ForeColor = p1.BackColor
Label11.ForeColor = p1.BackColor
Label1.Caption = "Mai nap (e hónap vagy év)"
End Sub

Private Sub Label9_Click()
ii = ii - 1
Dim dData As Date
If EVI = False Then
dData = DateAdd("m", ii, Date)
Label2.Caption = Format(dData, "dd/mm/yyyy") & " - pénzügyi infó"
MakeInfo Format(dData, "mm/yyyy")
Else
dData = DateAdd("yyyy", ii, Date)
Label2.Caption = Format(dData, "yyyy") & " - pénzügyi infó"
MakeEvi Format(dData, "yyyy")

End If
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = p1.BackColor

Label9.ForeColor = p2.BackColor
Label10.ForeColor = p1.BackColor
Label11.ForeColor = p1.BackColor
Label1.Caption = "Elõzõ hónap vagy év"
End Sub

Private Sub MSC_Click()
If MSC.Left = 60 Then
MSC.Left = 120
MSC.Top = 660
MSC.Width = 8715
MSC.Height = 5535

Else
MSC.Left = 60
MSC.Top = 60
MSC.Width = 11415
MSC.Height = 6495

End If
End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = p1.BackColor

Label9.ForeColor = p1.BackColor
Label10.ForeColor = p1.BackColor
Label11.ForeColor = p1.BackColor

End Sub

Private Sub Picture1_Click()
picInfo.Visible = True
picUj.Visible = False
picOpen.Visible = False

MakeInfo Format(Date, "mm/yyyy")

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Picture = Picture6.Picture
Picture2.Picture = Picture12.Picture
Picture3.Picture = Picture13.Picture
Picture4.Picture = Picture14.Picture
Picture5.Picture = Picture15.Picture
Label1.Caption = "Információk havi ill évi módban (váltás alul)"
End Sub

Private Sub Picture2_Click()
picInfo.Visible = False
picUj.Visible = True
picOpen.Visible = False
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Picture = Picture7.Picture
Picture1.Picture = Picture11.Picture
Picture3.Picture = Picture13.Picture
Picture4.Picture = Picture14.Picture
Picture5.Picture = Picture15.Picture
Label1.Caption = "Új dolgok hozzáadása"
End Sub

Private Sub Picture3_Click()
Data1.RecordSource = "T1"
Data1.Refresh
On Error Resume Next
List1.Clear

Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
List1.AddItem Data1.Recordset(0).Value & " - " & Trim(Data1.Recordset(2).Value) & "  (" & Trim(Data1.Recordset(1).Value) & ")"
Data1.Recordset.MoveNext
Loop

picInfo.Visible = False
picUj.Visible = False
picOpen.Visible = True
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Picture = Picture8.Picture
Picture1.Picture = Picture11.Picture
Picture2.Picture = Picture12.Picture
Picture4.Picture = Picture14.Picture
Picture5.Picture = Picture15.Picture
Label1.Caption = "Adatok megnyitása és törlése"
End Sub

Private Sub Picture4_Click()
Dim evel As String
evel = InputBox("Euro árfolyam:", "Beállitás", "38000")
If evel <> "" Then
SaveSetting App.Title, "set", "eu", evel
EUF = Val(evel)
End If
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture4.Picture = Picture9.Picture
Picture1.Picture = Picture11.Picture
Picture2.Picture = Picture12.Picture
Picture3.Picture = Picture13.Picture
Picture5.Picture = Picture15.Picture
Label1.Caption = "Beállitás - Euró árfolyam"
End Sub

Private Sub Picture5_Click()
Unload Me
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture5.Picture = Picture10.Picture
Picture1.Picture = Picture11.Picture
Picture2.Picture = Picture12.Picture
Picture3.Picture = Picture13.Picture
Picture4.Picture = Picture14.Picture
Label1.Caption = "Kilépés a programból"
End Sub

Private Sub Text3_GotFocus()
If Val(Text1.Text) <> 0 Then
Text2.Text = Format(Val(Text1.Text) / EUF, "###")
Else
Text1.Text = Format(Val(Text2.Text) * EUF, "###,###")
End If
End Sub
