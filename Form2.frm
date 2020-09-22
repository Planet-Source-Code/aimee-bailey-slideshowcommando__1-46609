VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SlideShow Commando v1.0"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   5415
      Begin VB.CheckBox Check4 
         Caption         =   "Show Info (I)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox Check3 
         Caption         =   "AutoSize Images  (R)"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Caption         =   " "
         Height          =   975
         Left            =   2760
         TabIndex        =   11
         Top             =   240
         Width           =   2535
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   960
            TabIndex        =   14
            Text            =   "1000"
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Timer  (T)"
            Height          =   255
            Left            =   180
            TabIndex        =   12
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "ms"
            Height          =   255
            Left            =   2160
            TabIndex        =   15
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "Delay"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Randomize (Y)"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   3000
      Pattern         =   "*.jpg;*.gif;*.bmp"
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   4440
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "M = Show Mouse"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "N = Hide Mouse"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Remember Num5 is quick exit!!"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0 Files Found"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If File1.ListCount < 1 Then GoTo err
Form1.Show
If Check3.Value = 1 Then Form1.ResiZ = True
If Check1.Value = 1 Then Form1.RandomZ = True
Form1.File1.Path = File1.Path
Form1.File1.Refresh
Form1.File1.ListIndex = 0
Form1.DoImage
If Check2.Value = 1 Then
    Form1.Timer1.Enabled = True
    Form1.Timer1.Interval = Text1.Text
End If


Unload Me
Exit Sub
err:
MsgBox "You cant run a slideshow without any pics!!!!"
End Sub

Private Sub Command2_Click()
End
End Sub



Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Refresh
Label1.Caption = (File1.ListCount - 1) & "File(s) Found."
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

