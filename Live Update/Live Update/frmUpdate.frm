VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Live Update - Step 1"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Next 
      Caption         =   "&Next"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   0
      Picture         =   "frmUpdate.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Live Update"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   1680
      Picture         =   "frmUpdate.frx":1530A
      Top             =   120
      Width           =   3660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1560
      X2              =   5550
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmUpdate.frx":1CE40
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Welcometext 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1560
      X2              =   5550
      Y1              =   3855
      Y2              =   3855
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
frmUpdate.Visible = False ' if u press the cancel button on the first form closes it and returns to to the main form
frmMenu.Visible = True ' Add your main form here to set it to visible
End Sub

Private Sub Next_Click()
frmUpdate2.Visible = True 'makes the next step appear
frmUpdate.Visible = False 'hides this form
frmUpdate2.NextButton.Enabled = False ' makes the next button on the next form enabled
frmUpdate2.Back.Enabled = False ' makes the Back button on the next form enabled
frmUpdate2.UpdateTimer.Enabled = True 'makes the updatetimer enabled
frmUpdate2.CheckConnect.Enabled = True 'makes the timer to check if u are connected to the net
frmUpdate2.Connectionstatus.Caption = "Connecting..." 'sets the caption of a label to Connecting...
End Sub
