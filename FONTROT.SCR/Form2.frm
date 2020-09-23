VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rotate Scr"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MaxLength       =   40
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "(Egypt +20) (12) 7242974"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "for quick support call me on"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "marco_s2@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "for more information send to:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "Rotate Text Screen Saver Creating and Full Programming by Marco Samy, all rights are reserved."
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Choose the Scentence to display :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim MyStr As String
MyStr = Text1.Text
If Trim(MyStr) = "" Then MyStr = "Marco Samy Nasif"
SaveSetting "TechnosoftScreen", "MarcoRotate", "SaverText", MyStr
End
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Form_Load()
Dim MyStr As String
MyStr = GetSetting("TechnosoftScreen", "MarcoRotate", "SaverText")
If Trim(MyStr) = "" Then MyStr = "Marco Samy Nasif"
Text1.Text = MyStr
End Sub
