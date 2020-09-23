VERSION 5.00
Begin VB.Form TweetForm 
   BackColor       =   &H00B6987B&
   Caption         =   "Twitter like a bird - using Pure Visual Basic 6.0 Code - INCLUDES XMLHTTP POST WITH AUTHENTICATION"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10890
   Icon            =   "TweetForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Response 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3960
      Width           =   10455
   End
   Begin VB.CommandButton TweetButton 
      BackColor       =   &H00C0C0C0&
      Caption         =   "POST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox YourTweet 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      MaxLength       =   140
      TabIndex        =   7
      Top             =   2880
      Width           =   9375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00B6987B&
      Height          =   2100
      Left            =   4440
      ScaleHeight     =   2040
      ScaleWidth      =   6075
      TabIndex        =   1
      Top             =   240
      Width           =   6135
      Begin VB.TextBox TheirUsername 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox YourPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   450
         Width           =   2775
      End
      Begin VB.TextBox YourUserName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   3
         Top             =   450
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipient's User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your own Twitter User Name and Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4500
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2100
      Left            =   120
      Picture         =   "TweetForm.frx":030A
      ScaleHeight     =   2040
      ScaleWidth      =   4170
      TabIndex        =   0
      Top             =   240
      Width           =   4230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XML RESPONSE FROM TWITTER.COM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   4320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your ""TWEET"" ( Simple TEXT up to 140 Characters in Length )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   6510
   End
End
Attribute VB_Name = "TweetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TweetButton_Click()
'================================================================
'CHECK THE TWITTER API FOR "GET" OPTIONS AND OTHER "POST" OPTIONS
'================================================================
 UN$ = Trim$(YourUserName)
 If UN$ = "" Then
  Beep
  MsgBox "Sorry, you need to enter your Twitter Username.", vbExclamation, "Whoops!"
  Exit Sub
 End If
 PW$ = Trim$(YourPassword)
 If PW$ = "" Then
  Beep
  MsgBox "Sorry, you need to enter your Twitter Password.", vbExclamation, "Whoops!"
  Exit Sub
 End If
 Recipient$ = Trim$(TheirUsername)
 If Recipient$ = "" Then
  Beep
  MsgBox "Sorry, you need to enter your Friend's Username.", vbExclamation, "Whoops!"
  Exit Sub
 End If
 Tweet$ = Trim$(YourTweet)
 If Tweet$ = "" Then
  Beep
  MsgBox "Sorry, a Tweet may be between 1 and 140 characters in length.", vbExclamation, "Whoops!"
  Exit Sub
 End If
'==========================
'DEFINE THE TWITTER API URL
'==========================
 cURL$ = "http://twitter.com/direct_messages/new.xml?user=" & Recipient$
 cURL$ = cURL$ & "&text=" & Tweet$
'=============================================
'POST THE STRING WITH USER/PASS AUTHENTICATION
'=============================================
 Screen.MousePointer = 11
 Response = CredentialPostURLSource(cURL$, UN$, PW$)
 Screen.MousePointer = Default
'===============================================================
'REMEMBER, YOU ARE LIMITED TO 100 REQUESTS PER HOUR WITH TWITTER
'===============================================================
End Sub

Sub BuildPostData(BYTEARRAY() As Byte, ByVal strPostData As String)
 Dim lngNewBytes As Long
 Dim strCH As String
 Dim i As Long
 lngNewBytes = Len(strPostData) - 1
 If lngNewBytes < 0 Then
  Exit Sub
 End If
 ReDim BYTEARRAY(lngNewBytes)
 For i = 0 To lngNewBytes
  strCH = Mid$(strPostData, i + 1, 1)
  BYTEARRAY(i) = Asc(strCH)
 Next
End Sub

Function UrlEncode(sText As String) As String
 sText = Replace(sText, " ", "+")
 For i = 1 To Len(sText)
  sChar = Mid$(sText, i, 1)
  If InStr("+abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", sChar) Then
   sResult = sResult & sChar
  Else
   sResult = sResult & "%" & Right$("0" & Hex(Asc(sChar)), 2)
  End If
 Next
 UrlEncode = sFinal & sResult
End Function

Function CredentialPostURLSource(TheURL As String, UN As String, PS As String) As String
'======================================================
'"?" Prefix 1st Field, "&" prefix for subsequent fields
'======================================================
 S = InStr(TheURL, "?")
 If S = 0 Then
  Exit Function
 End If
 SiteASP$ = Left$(TheURL, S - 1)
 StringtoPost = Right$(TheURL, Len(TheURL) - S)
 Dim bytpostdata() As Byte
 Dim strPostData As String
 Dim strHeader As String
 Dim varPostData As Variant
'====================================
'Pack the post data into a byte array
'====================================
 strPostData = StringtoPost
 BuildPostData bytpostdata(), strPostData
'=============================
'Write the byte into a variant
'=============================
 varPostData = bytpostdata
'=================
'Create the Header
'=================
 strHeader = "application/x-www-form-urlencoded" + Chr(10) + Chr(13)
'=============
'Post the data
'=============
 Dim xmlhttp As Object
 Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
 xmlhttp.Open "POST", SiteASP$, False, UN, PS
 xmlhttp.setRequestHeader "Content-Type", strHeader
 xmlhttp.Send varPostData
 HTTPText$ = xmlhttp.responseText
 Set xmlhttp = Nothing
 CredentialPostURLSource = HTTPText$
End Function

