VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Main screen"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   1575
      Left            =   1800
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   $"Live Update.frx":0000
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Make this the main form of your program you can call it frmMenu or just renam the code in the update"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmUpdate.Visible = True ' makes the first step appear
End Sub
