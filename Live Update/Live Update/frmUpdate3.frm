VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Live Update - Step 3"
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
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   0
   End
   Begin VB.CommandButton cmdFin 
      Caption         =   "&Finish"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   3120
      Width           =   3975
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Connected"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3975
      End
   End
   Begin MSComctlLib.ProgressBar UpdateStatus 
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   3480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   0
      Picture         =   "frmUpdate3.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   5160
      Picture         =   "frmUpdate3.frx":1530A
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1560
      Picture         =   "frmUpdate3.frx":1574C
      Top             =   2520
      Width           =   480
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
      Index           =   0
      Left            =   1680
      TabIndex        =   17
      Top             =   720
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   2280
      Picture         =   "frmUpdate3.frx":15B8E
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   8
      Left            =   1680
      TabIndex        =   16
      Top             =   2805
      Width           =   435
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   1
      Left            =   2160
      TabIndex        =   15
      Top             =   2805
      Width           =   555
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   2
      Left            =   2760
      TabIndex        =   14
      Top             =   2805
      Width           =   435
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   3
      Left            =   3240
      TabIndex        =   13
      Top             =   2805
      Width           =   435
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   4
      Left            =   3720
      TabIndex        =   12
      Top             =   2805
      Width           =   435
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   5
      Left            =   4200
      TabIndex        =   11
      Top             =   2805
      Width           =   435
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   6
      Left            =   4680
      TabIndex        =   10
      Top             =   2805
      Width           =   315
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   7
      Left            =   5040
      TabIndex        =   9
      Top             =   2805
      Width           =   435
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3120
      Picture         =   "frmUpdate3.frx":15FD0
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
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
   Begin VB.Image Image1 
      Height          =   645
      Left            =   1680
      Picture         =   "frmUpdate3.frx":16412
      Top             =   120
      Width           =   3660
   End
   Begin VB.Label Welcometext 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Continue..."
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
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please click the ""Update"" button to update If at any time you want to cancel this update please click the cancel button."
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
      TabIndex        =   8
      Top             =   1680
      Width           =   3615
   End
End
Attribute VB_Name = "frmUpdate3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataByte() As Byte
Dim I As Integer                'Code for downloading the file
Public Terr As Boolean
Public Dw_Url As String

Function AddBackSlash(TPathName As String) As String
'Code for the downloading of the file DO NOT EDIT THE ABOVE OR THIS CODE!!!!
    If Right(TPathName, 1) = "\" Then AddBackSlash = TPathName Else AddBackSlash = TPathName & "\"
End Function
Private Sub Cancel_Click()
Dim GetFile
    If Inet1.StillExecuting Then
        GetFile = MsgBox("Are you sure you want to cancel the update...", _
        vbYesNo)
        If GetFile = vbNo Then
            Exit Sub
        Else                     'Cancels the download and cancels the update
        frmUpdate3.Visible = False 'then returns to main form
        frmMenu.Visible = True
        Inet1.Cancel
        Terr = True
    End If
    End If
End Sub



Private Sub cmdupdate_Click()
Timer2.Enabled = True
Label2.Caption = "Please wait... while we download and update for your conviniance... This Process should not take more than a few minutes..."
On Error Resume Next
Dim TFile As Long
    TFile = FreeFile
    Label3.Caption = "Please wait downloading update..." & New_Ver
    Timer1.Enabled = True
                  '''Edit the link here of the file'''
    Dw_Url = "http://www.SVKputfilenamehere.com/update.zip"
    DataByte() = Inet1.OpenURL(Dw_Url, 1)
                            'Change file name of download (see update folder for info)
            Open AddBackSlash(App.Path) & "update.zip" For Binary As #TFile
                Put #TFile, , DataByte()
            Close #TFile
            'DO NOT EDIT ANY OF THIS CODE!!!!!!!!!
            'EXCEPT WHAT I HAVE SAID 2 EDIT!!!!
            If Inet1.StillExecuting = False Then
                Label3.Caption = "Update Downloaded"
                'update is complete
                Timer1.Enabled = False
                Label1(I + 1).BackColor = vbBlack
                Image2.Picture = Image4.Picture
                cmdFin.Enabled = True
                cmdCancel.Enabled = False
                cmdupdate.Enabled = False
                Inet1.Cancel
            End If
            Welcometext.Caption = "Complete..."
            UpdateStatus.Value = "100"
            Timer2.Enabled = False
            Label3.Caption = "Download Complete"
       ' MsgBox "You have successfully download the Virus Klean update file " & (App.Path) & "\update.zip please wait while we automatically install this ", vbInformation
cmdFin.Enabled = True
Cancel.Enabled = False
cmdupdate.Enabled = False
End Sub


Private Sub Timer1_Timer()

'Makes the red lijnes go across for animation
    I = I + 1
    If I = 7 Then: Image2.Picture = Image3.Picture: I = 0: Label1(7).BackColor = vbBlack
    Label1(I).BackColor = vbBlack
    Label1(I + 1).BackColor = vbRed
    If I = 1 Then Image2.Picture = Image4.Picture
End Sub

Private Sub Timer2_Timer()
'add's and clears the progressbar for better
'looking update
If UpdateStatus.Value > 94 Then UpdateStatus.Value = "0"
'If the progressbar is higher than 94 it will clear it then restart
y = 2
UpdateStatus.Value = UpdateStatus.Value + y
End Sub
Private Sub cmdFin_Click()
    Inet1.Cancel     'Closes the Inet control
'gives a msg the program is going to end then reboot itself
MsgBox "This program will now restart...", vbInformation, "Samosoft Virus Klean"
Shell "MyAppUpdate.exe"  ' Opens the updater See Update folder for the code and how it works
End 'Ends program
End Sub
