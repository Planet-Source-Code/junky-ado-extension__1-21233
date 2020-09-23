VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ADO Extension"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   795
      TabIndex        =   0
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Feel free to use any part of the code for educational purpose."
      Height          =   495
      Left            =   330
      TabIndex        =   9
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "www.junky-tech.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1110
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "URL:"
      Height          =   255
      Left            =   510
      TabIndex        =   7
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "E-mail:"
      Height          =   255
      Left            =   510
      TabIndex        =   6
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "This is an example on how to use ADO Extension to create Access 2000 tables."
      Height          =   435
      Left            =   330
      TabIndex        =   5
      Top             =   1320
      Width           =   2880
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "cool_junkman@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1110
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Created by JunKy @ JunKy Technology 2001"
      Height          =   375
      Left            =   323
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "ADO Extension"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Label1"
      Height          =   510
      Left            =   795
      TabIndex        =   1
      Top             =   3360
      Width           =   1950
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Label4_Click()
    Shell "c:\program files\internet explorer\iexplore.exe mailto:cool_junkman@yahoo.com", vbMaximizedFocus
End Sub

Private Sub Label8_Click()
    Shell "c:\program files\internet explorer\iexplore.exe www.junkytech.com", vbMaximizedFocus
End Sub
