VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9465
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1680
      Top             =   2880
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   240
      Pattern         =   "*.jpg;*.gif;*.bmp"
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   6855
      Left            =   1200
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RandomZ As Boolean
Public ResiZ As Boolean

Public Sub LogicalSize(ContainerObj As Object, ImgObj As Image, ByVal Cushion As Integer)
    Dim VertChg, HorzChg As Integer
    Dim iRatio As Double
    Dim ActualH, ActualW As Integer
    Dim ContH, ContW As Integer
    On Error GoTo LogicErr
    
    'ImgObj.Width = Me.Width + 1000
    'ImgObj.Height = Me.Width + 1000

    With ImgObj 'hide picture While changing size
        .Visible = False
        .Stretch = False 'set actual size
    End With
    VertChg = 0: HorzChg = 0
    
    ActualH = ImgObj.Height 'actual picture height
    ActualW = ImgObj.Width 'actual picture width
    ContH = ContainerObj.Height - Cushion 'set max. picture height
    ContW = ContainerObj.Width - Cushion 'set max. picture width
    CenterCTL ContainerObj, ImgObj 'center picture
    CenterCTL Form1, File1

    If ResiZ = True Then
    If ActualW > ActualH Then
        iRatio = (ActualW / ActualH)
        ActualW = Me.Width
        ActualH = ActualW / iRatio
    ElseIf ActualH > ActualW Then
        iRatio = (ActualH / ActualW)
        ActualH = Me.Height
        ActualW = ActualH / iRatio
    Else
        ActualW = Me.Height
        ActualH = Me.Height
    End If
    





        With ImgObj 'set new height and width
            .Stretch = True
            .Height = ActualH
            .Width = ActualW
        End With
    End If
    
    CenterCTL ContainerObj, ImgObj 'center picture in container
    ImgObj.Visible = True 'show picture
    Exit Sub
LogicErr:
    MsgBox "An Error occured While rescaling this image. Image size maybe invalid.", vbSystemModal + vbExclamation, "Resize Error!"
End Sub












Public Function DoImage()
Dim X As String
Image1.Visible = False
If Right(File1.Path, 1) <> "\" Then X = "\" Else X = ""
Image1.Picture = LoadPicture(File1.Path & X & File1.FileName)
LogicalSize Form1, Image1, 0
'Form1.Cls
'Form1.Print File1.Path & X & File1.FileName
End Function



Public Sub CenterCTL(FRMObj As Object, OBJ As Control)


    With OBJ
        .Top = (FRMObj.Height / 2) - (OBJ.Height / 2)
        .Left = (FRMObj.Width / 2) - (OBJ.Width / 2)
        .ZOrder
    End With
End Sub



Private Sub Form_DblClick()
ShowCursor (bShow = False)
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
'Form1.Cls
'Form1.Print File1.Path & X & File1.FileName
If KeyAscii = 13 Or KeyAscii = 54 Then
    If File1.ListIndex = File1.ListCount - 1 Then
    File1.ListIndex = 0
    Else
        If RandomZ = True Then
            Randomize
            File1.ListIndex = Int(Rnd * (File1.ListCount - 1))
        Else
            File1.ListIndex = File1.ListIndex + 1
        End If
    End If
    DoImage
ElseIf KeyAscii = 52 Then
    If File1.ListIndex = 0 Then
    File1.ListIndex = File1.ListCount - 1
    Else
        If RandomZ = True Then
            Randomize
            File1.ListIndex = Int(Rnd * (File1.ListCount - 1))
        Else
            File1.ListIndex = File1.ListIndex - 1
        End If
    End If
    DoImage
ElseIf KeyAscii = 27 Or KeyAscii = 53 Then
    ShowCursor (bShow = False)
    End
ElseIf KeyAscii = 43 Or KeyAscii = 56 Then
    Image1.Visible = False
    Image1.Height = Image1.Height + 1000
    Image1.Width = Image1.Width + 1000
    CenterCTL Form1, Image1
    Image1.Visible = True
ElseIf KeyAscii = 50 Then
    Image1.Visible = False
    Image1.Height = Image1.Height - 1000
    Image1.Width = Image1.Width - 1000
    CenterCTL Form1, Image1
    Image1.Visible = True
ElseIf KeyAscii = 110 Then
    ShowCursor (bShow = True)
ElseIf KeyAscii = 109 Then
    ShowCursor (bShow = False)
ElseIf KeyAscii = 114 Then
If ResiZ = True Then ResiZ = False Else ResiZ = True
ElseIf KeyAscii = 116 Then
If Timer1.Enabled = True Then Timer1.Enabled = False Else Timer1.Enabled = True
ElseIf KeyAscii = 121 Then
If RandomZ = True Then RandomZ = False Else RandomZ = True
End If
End Sub

Private Sub Form_Load()
ShowCursor (bShow = True)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ShowCursor (bShow = False)
End Sub

Private Sub Form_Resize()
On Error Resume Next
Label1.Width = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowCursor (bShow = False)
End Sub

Private Sub Image1_DblClick()
ShowCursor (bShow = False)
X = MsgBox("Exit??", vbYesNo, "SlideShow Commando!")
If X = vbYes Then
End
Else
ShowCursor (bShow = True)
End If

End Sub

Private Sub Timer1_Timer()
If File1.ListIndex = File1.ListCount - 1 Then
    File1.ListIndex = 0
    Else
        If RandomZ = True Then
            Randomize
            File1.ListIndex = Int(Rnd * (File1.ListCount - 1))
        Else
            File1.ListIndex = File1.ListIndex + 1
        End If
    End If
    DoImage
End Sub
