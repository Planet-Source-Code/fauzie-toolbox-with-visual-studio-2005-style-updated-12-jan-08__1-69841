VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ToolBox Control"
   ClientHeight    =   2085
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4815
   ClipControls    =   0   'False
   Icon            =   "fAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   139
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      Picture         =   "fAbout.frx":000C
      ScaleHeight     =   421.4
      ScaleMode       =   0  'User
      ScaleWidth      =   421.4
      TabIndex        =   1
      Top             =   360
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   0
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fauzie's"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   6
      X2              =   312
      Y1              =   91
      Y2              =   91
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2007 by Fauzie's Software"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1170
      TabIndex        =   2
      Top             =   975
      Width           =   2985
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ToolBox Control"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1170
      TabIndex        =   3
      Top             =   390
      Width           =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      Index           =   0
      X1              =   7
      X2              =   312
      Y1              =   92
      Y2              =   92
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "v2.00"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3675
      TabIndex        =   4
      Top             =   585
      Width           =   390
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub
