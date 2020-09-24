VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdate2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Live Update - Step 2"
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
   Begin VB.Timer CheckConnect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   3240
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   3915
      TabIndex        =   8
      Top             =   2520
      Width           =   3975
      Begin VB.Label Connectionstatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Connected"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.Timer UpdateTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   3240
   End
   Begin VB.CommandButton Back 
      Caption         =   "&Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar UpdateStatus 
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2880
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton NextButton 
      Caption         =   "&Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   0
      Picture         =   "frmUpdate2.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton CloseUpdate 
      Caption         =   "&End"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   1680
      Picture         =   "frmUpdate2.frx":1530A
      Top             =   120
      Width           =   3660
   End
   Begin VB.Label Welcometext 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
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
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while we connect to this company who made thos program to check for an update..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   3615
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
      TabIndex        =   3
      Top             =   720
      Width           =   3375
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1560
      X2              =   5550
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Label UpdateYes 
      BackStyle       =   0  'Transparent
      Caption         =   "There is a update avalible, please click next..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "frmUpdate2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RASCONN
    dwSize As Long
    hRasConn As Long
    szEntryName(256) As Byte          ' CODE TO CHECK IF UN ARE CONNECTED TO THE INTERNET
    szDeviceType(16) As Byte
    szDeviceName(128) As Byte
    End Type

Private Declare Function RasEnumConnectionsA& Lib "RasApi32.DLL" (lprasconn As Any, lpcb&, lpcConnections&)


Private Sub Back_Click()
frmUpdate.Visible = True        'back button if u want to go back a step it will close this from
frmUpdate2.Visible = False       'hide this form and make the first appear again
UpdateStatus.Value = "0"      'sets the progressbar value 2 "0" so u dont get errors
Connectionstatus.Caption = "Disconnected" 'sets the caption text to Disconnected
End Sub

Private Sub Cancel_Click()
frmUpdate.Visible = False 'if u click cancel resets everything and go's back to main form
frmUpdate2.Visible = False
frmMenu.Visible = True
UpdateStatus.Value = "0" 'sets progressbar value 2 "0" so u dont get errors
UpdateTimer.Enabled = False ' disabled the timer
End Sub

Private Sub Command1_Click()
frmUpdate.Visible = True       'proceeds 2 next step
frmUpdate2.Visible = False
End Sub


Private Sub CheckConnect_Timer()
'THIS CODE CHECKS IF U ARE CONNECTED 2 THE NET IF U ARE NOT THEN U CANNOT CARRY ON WITH THE UPDATE!!!
    Dim Checkin As RASCONN
    Dim StatusSize, CheckOk As Long
    Checkin.dwSize = 412
    StatusSize = Checkin.dwSize


    If RasEnumConnectionsA(Checkin, StatusSize, CheckOk) = 0 Then


        If CheckOk = 0 Then
             UpdateTimer.Enabled = False
            Connectionstatus.Caption = "You are not connected to the Net..."
            UpdateStatus.Value = "0"
            UpdateTimer.Enabled = False
            Cancel.Enabled = False
            Back.Enabled = False
            NextButton.Enabled = False
            NextButton.Visible = False
            CloseUpdate.Visible = True
            UpdateYes.Caption = "You are not connected to the Internet please connect to the Internet before attempting to update..."
        Else
            CheckConnect.Enabled = False
        End If
    End If

End Sub

Private Sub CloseUpdate_Click()
frmUpdate.Visible = False
frmUpdate2.Visible = False
frmUpdate3.Visible = False
UpdateTimer.Enabled = False     'Cancels everything and returns to main form
UpdateStatus.Value = "0"
Connectionstatus.Caption = "Disconnected"
frmMenu.Visible = True
End Sub


Private Sub NextButton_Click()
UpdateTimer.Enabled = False
UpdateStatus.Value = "0"
frmUpdate3.Visible = True      'Proceeds to next step
frmUpdate2.Visible = False
frmUpdate3.cmdupdate.Enabled = True
End Sub

Private Sub UpdateTimer_Timer()
'U will need to get code to check if a file on a website
'exists so for instance if  MyAppUpdate.exe
'exists on the site then u can
'carry on otherwise halt this operation and close the update process
'If someone knows this code please comment it and e-mail
'it to me!!!  Sammyrhys@hotmail.com or ICQ : 113682686 Thnx!

If UpdateStatus.Value > 97 Then
UpdateTimer.Enabled = False
NextButton.Enabled = True
Back.Enabled = True       'Adds 2 to the value of the progressbar
UpdateYes.Visible = True       'then when it gets to higher than 97 it stps itself so u dont get an error
Connectionstatus.Caption = "Connection Establised"
End If

y = 2          'Adds itself to the value of the progressbar to make it go up
UpdateStatus.Value = UpdateStatus.Value + y
End Sub
