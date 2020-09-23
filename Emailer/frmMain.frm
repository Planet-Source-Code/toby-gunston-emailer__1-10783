VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emailer"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClearStatus 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   5040
   End
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   480
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrLoadDefault 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5640
      Top             =   5160
   End
   Begin VB.Timer tmrLogo1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5160
      Top             =   5160
   End
   Begin VB.Timer tmrLogo 
      Interval        =   500
      Left            =   4680
      Top             =   5160
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   25000
      Left            =   5640
      Top             =   4680
   End
   Begin VB.TextBox txtFrom 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      ToolTipText     =   "E-mail address"
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtSubject 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Subject of E-mail"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txtTo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "E-mail address"
      Top             =   120
      Width           =   3735
   End
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   4680
   End
   Begin VB.Timer tmrCheckFields 
      Interval        =   1
      Left            =   0
      Top             =   4560
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "This button will be enabled when all fields are filled in"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtBody 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Body of the message"
      Top             =   2040
      Width           =   5955
   End
   Begin VB.TextBox txtServer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Your server name"
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Caption         =   "Status:"
      ForeColor       =   &H80000007&
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   3840
      Width           =   6135
      Begin VB.Label lblSendingStatus 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5775
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5160
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblSaveEmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image imgLogo 
      Height          =   735
      Left            =   5160
      Picture         =   "frmMain.frx":0000
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String
Dim Start As Single, Tmr As Single

Private Sub cmdClear_Click()
'Clear the form and disable timers that might be running
txtTo = ""
txtFrom = ""
txtSubject = ""
txtServer = ""
txtBody = ""
cmdSend.Enabled = True
lblSendingStatus.Caption = ""
tmrAnimate.Enabled = False
tmrTimeOut.Enabled = False
txtTo.SetFocus
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Put the labels back to how they were
'while we are not pointing at them.
lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H0& 'Black
lblAbout.FontSize = 8
lblAbout.ForeColor = &H0& 'Black
End Sub

Private Sub cmdSend_Click()
'Do stuff here like disable this button and activate the
'timers to animate the sending caption and activate the
'timeout timer that will cancel the email if it is still
'trying to send after 25 seconds.
tmrCheckFields.Enabled = False
cmdSend.Enabled = False
tmrTimeOut.Enabled = True
lblSendingStatus.Caption = "Sending"
tmrAnimate.Enabled = True
'Call the SendEmail procedure
SendEmail txtTo, txtFrom, txtSubject, txtBody, txtServer
tmrAnimate.Enabled = False
lblSendingStatus.Caption = " Mail sent successfully "
Beep
tmrCheckFields.Enabled = True
'cmdSend.Enabled = True
tmrClearStatus.Enabled = True
tmrTimeOut.Enabled = False
Close
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Put the labels back to how they were
'while we are not pointing at them.
lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H0& 'Black
lblAbout.FontSize = 8
lblAbout.ForeColor = &H0& 'Black
End Sub

Private Sub imgLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Put the labels back to how they were
'while we are not pointing at them.
lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H0& 'Black
lblAbout.FontSize = 8
lblAbout.ForeColor = &H0& 'Black
End Sub

Private Sub lblAbout_Click()
frmAbout.Show
End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Highlight me
lblAbout.FontSize = 10
lblAbout.ForeColor = &HFF&
End Sub

Private Sub lblSaveEmail_Click()
'Show text files in the "Save as" box
cdSave.Filter = "Text files (*.txt)|*.txt|"
'Default filename is "Email"
cdSave.FileName = "Email"
'Show the "Save as" dialog as apposed to "Open"
cdSave.ShowSave
'Open the file they saved and write the data to it
Open cdSave.FileName For Append As #1
Print #1, "-------------------------"
Print #1, txtTo
Print #1, txtFrom
Print #1, txtSubject
Print #1, txtServer
Print #1,
Print #1, txtBody
Print #1, "-------------------------"
Close #1
End Sub

Private Sub lblSaveEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Highlight me
lblSaveEmail.FontSize = 10
lblSaveEmail.ForeColor = &HFF&
End Sub

Private Sub tmrAnimate_Timer()
'Make dots in the caption
lblSendingStatus.Caption = lblSendingStatus.Caption + " ."
End Sub

Private Sub tmrCheckFields_Timer()
'Check all fields are filled in before enabling the Send
'button else it will be disabled
If txtFrom <> "" And txtTo <> "" And txtSubject <> "" And txtServer <> "" And txtBody <> "" Then
 cmdSend.Enabled = True
Else
 cmdSend.Enabled = False
End If
End Sub

Private Sub tmrClearStatus_Timer()
lblSendingStatus.Caption = ""
tmrClearStatus.Enabled = False
End Sub

'The following timers control loading the pictures into
'the imgLogo to make the animation
Private Sub tmrLoadDefault_Timer()
imgLogo.Picture = LoadPicture(App.Path & "\vb.bmp")
tmrLoadDefault.Enabled = False
tmrLogo.Enabled = True
End Sub

Private Sub tmrLogo_Timer()
imgLogo.Picture = LoadPicture(App.Path & "\vb1.bmp")
tmrLogo.Enabled = False
tmrLogo1.Enabled = True
End Sub

Private Sub tmrLogo1_Timer()
imgLogo.Picture = LoadPicture(App.Path & "\vb2.bmp")
tmrLogo1.Enabled = False
tmrLoadDefault.Enabled = True
End Sub

Private Sub tmrTimeOut_Timer()
lblSendingStatus.Caption = "Error"
tmrAnimate.Enabled = False
MsgBox "Connection timed out, make sure all fields except the body have no spaces", vbCritical, "Connection timed out"
End Sub

Private Sub txtFrom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Put the labels back to how they were
'while we are not pointing at them.
lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H0& 'Black
lblAbout.FontSize = 8
lblAbout.ForeColor = &H0& 'Black
End Sub

Private Sub txtServer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Put the labels back to how they were
'while we are not pointing at them.
lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H0& 'Black
lblAbout.FontSize = 8
lblAbout.ForeColor = &H0& 'Black
End Sub

Private Sub txtSubject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Put the labels back to how they were
'while we are not pointing at them.
lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H0& 'Black
lblAbout.FontSize = 8
lblAbout.ForeColor = &H0& 'Black
End Sub

Private Sub txtTo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Put the labels back to how they were
'while we are not pointing at them.
lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H0& 'Black
lblAbout.FontSize = 8
lblAbout.ForeColor = &H0& 'Black
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Response ' Check for incoming response
End Sub

'This is the procedure to send the E-mail
Sub SendEmail(EmailTo As String, From As String, Subject As String, Body As String, Server As String)
          
Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail pre program start
    
If Winsock1.State = sckClosed Then ' Check to see if socet is closed
    DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
    first = "mail from:" + Chr(32) + txtFrom + vbCrLf ' Get who's sending E-Mail address
    Second = "rcpt to:" + Chr(32) + txtTo + vbCrLf ' Get who mail is going to
    Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
    Fourth = "From:" + Chr(32) + vbCrLf ' Who's Sending
    Fifth = "To:" + Chr(32) + vbCrLf ' Who's it going to
    Sixth = "Subject:" + Chr(32) + txtSubject + vbCrLf ' Subject of E-Mail
    Seventh = txtBody + vbCrLf ' E-mail body
    Ninth = "Emailer 0.1" + vbCrLf ' What program sent the e-mail, customize this
    Eighth = Fourth + Third + Ninth + Fifth + Sixth  ' Combine for proper SMTP sending

    Winsock1.Protocol = sckTCPProtocol ' Set protocol for sending
    Winsock1.RemoteHost = txtServer ' Set the server address
    Winsock1.RemotePort = 25 ' Set the SMTP Port
    Winsock1.Connect ' Start connection
    
    WaitFor ("220")
    'Tell the user the program is connecting
    lblSendingStatus.Caption = "Connecting"
    
    Winsock1.SendData ("HELO hotmail.com" + vbCrLf)

    WaitFor ("250")

    lblSendingStatus.Caption = "Connected"
    
    Winsock1.SendData (first)

    WaitFor ("250")

    Winsock1.SendData (Second)

    WaitFor ("250")

    Winsock1.SendData ("data" + vbCrLf)
    
    WaitFor ("354")


    Winsock1.SendData (Eighth + vbCrLf)
    Winsock1.SendData (Seventh + vbCrLf)
    Winsock1.SendData ("." + vbCrLf)

    WaitFor ("250")

    Winsock1.SendData ("quit" + vbCrLf)
    
    WaitFor ("221")

    Winsock1.Close
Else
    MsgBox (Str(Winsock1.State))
End If
   
End Sub
Sub WaitFor(ResponseCode As String)
    Start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = Start - Timer
        DoEvents ' Let System keep checking for incoming response
        If Tmr > 50 Then ' Time to wait in seconds
            MsgBox "SMTP service error, timed out while waiting for response", 64, MsgTitle
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64, MsgTitle
            Exit Sub
        End If
    Wend
Response = "" ' Send response code to blank
End Sub


