VERSION 5.00
Begin VB.Form frm_SystemInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Information"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frm_SystemInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUser 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox txtCoy 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtNav 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox txtIE 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox txtRev 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtVer 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton cmdOkay 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   3600
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4560
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "User:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Company:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Netscape:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "IE Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Revision:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   705
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      Caption         =   "System Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frm_SystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOkay_Click()
    End
End Sub

Private Sub Form_Load()
    Dim oVer As New cVersion
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    With oVer
        txtVer.Text = .WindowsVersion
        txtRev.Text = .Revision
        txtCoy.Text = .RegisteredOrg
        txtUser.Text = .RegisteredUser
        txtIE.Text = .BrowserVersion
        txtNav.Text = .NetscapeVersion
    End With
    Set oVer = Nothing
End Sub

Private Sub txtCoy_GotFocus()
    txtCoy.SelStart = 0
    txtCoy.SelLength = Len(txtCoy.Text)
End Sub

Private Sub txtIE_GotFocus()
    txtIE.SelStart = 0
    txtIE.SelLength = Len(txtIE.Text)
End Sub

Private Sub txtNav_GotFocus()
    txtNav.SelStart = 0
    txtNav.SelLength = Len(txtNav.Text)
End Sub

Private Sub txtRev_GotFocus()
    txtRev.SelStart = 0
    txtRev.SelLength = Len(txtRev.Text)
End Sub

Private Sub txtUser_GotFocus()
    txtUser.SelStart = 0
    txtUser.SelLength = Len(txtUser.Text)
End Sub

Private Sub txtVer_GotFocus()
    txtVer.SelStart = 0
    txtVer.SelLength = Len(txtVer.Text)
End Sub
