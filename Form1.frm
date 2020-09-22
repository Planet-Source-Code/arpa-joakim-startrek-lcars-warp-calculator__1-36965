VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LCARS Warp Calculator Version 1.0.0"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   13590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWarp 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF00FF&
      Height          =   285
      HideSelection   =   0   'False
      Left            =   3480
      TabIndex        =   3
      Text            =   "0"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.HScrollBar hsbWarp 
      Height          =   255
      Left            =   2280
      Max             =   9999
      TabIndex        =   2
      Top             =   3720
      Width           =   6075
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   1560
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   2895
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   480
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   1560
      OLEDragMode     =   1  'Automatic
      Picture         =   "Form1.frx":26A9
      ScaleHeight     =   3015
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   480
      Width           =   6495
   End
   Begin VB.Timer Timer1 
      Left            =   10680
      Top             =   2040
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4800
      Left            =   11900
      Picture         =   "Form1.frx":43E7
      ScaleHeight     =   4800
      ScaleWidth      =   1800
      TabIndex        =   6
      Top             =   850
      Width           =   1800
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   9350
      Picture         =   "Form1.frx":5349
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   5
      Top             =   0
      Width           =   4335
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   9350
      Picture         =   "Form1.frx":61EC
      ScaleHeight     =   2475
      ScaleWidth      =   2655
      TabIndex        =   4
      Top             =   3180
      Width           =   2655
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      Picture         =   "Form1.frx":6D89
      ScaleHeight     =   5655
      ScaleWidth      =   9405
      TabIndex        =   7
      Top             =   0
      Width           =   9410
      Begin VB.Label lblMetres 
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   6960
         TabIndex        =   13
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Meters/sec:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004EE0E4&
         Height          =   255
         Left            =   4680
         TabIndex        =   12
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label lblMethod 
         BackColor       =   &H00000000&
         Caption         =   "Label6"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label lblLight 
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   6960
         TabIndex        =   10
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Times the speed of light:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004EE0E4&
         Height          =   255
         Left            =   4680
         TabIndex        =   9
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Warp Speed:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004EE0E4&
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   4320
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X(100), Y(100), Z(100) As Integer
Dim tmpX(100), tmpY(100), tmpZ(100) As Integer
Dim K As Integer
Dim Zoom As Integer
Dim Speed As Integer


Sub subWARP()
Dim l001E As Variant
Dim l0022 As String
Dim l0024 As Variant
    l0022 = txtWarp.Text
    l0024 = Val(l0022)
    l001E = (10 / 3) + (1 / (10 ^ 5 - l0024 ^ 5))
    l001E = l0024 ^ l001E
    If l0024 >= 9.99 Then
        l001E = l001E + 5755 + 190347 * (l0024 - 9.99) / 0.0099
    ElseIf l0024 >= 9.9 Then
        l001E = l001E + 969 + 4786 * (l0024 - 9.9) / 0.09
    ElseIf l0024 >= 9.6 Then
        l001E = l001E + 29 + 940 * (l0024 - 9.6) / 0.3
    ElseIf l0024 >= 9.2 Then
        l001E = l001E + 18 + 11 * (l0024 - 9.2) / 0.4
    ElseIf l0024 >= 9 Then
        l001E = l001E + 18 * (l0024 - 9) / 0.2
    End If
    If l0024 > 9 Then
        lblMethod.ForeColor = RGB(255, 0, 0)
        lblMethod.Caption = "Not recommended."
    Else
        lblMethod.ForeColor = RGB(0, 255, 0)
        lblMethod.Caption = "Safe."
    End If
    lblLight.Caption = Format$(l001E, "standard")
    lblMetres.Caption = Format$(Int(l001E * 2.99792458 * 10 ^ 8), "###,###,###,###")
End Sub

Sub cmdExit_Click()
End
End Sub

Private Sub Exit_Click()
End

End Sub

Sub Form_Load()
' Dim l0040 As Variant
' Dim l0044 As Integer
' Const mc004E = 9 ' &H9%
'   If App.PrevInstance = True Then
'      frmMain.Caption = "bye bye"
'      l0040 = extfn00A1%(0&, "Warp")
'      l0044 = extfn00AF%(l0040, mc004E)
'      Unload Me
'      Exit Sub
'   End If
'   Load frmAbout
End Sub

Sub Form_Unload(Cancel As Integer)
   End
End Sub

Sub hsbWarp_Change()
    Dim A As Double
 A = hsbWarp.Value
Picture2.Width = (1000) + (A * (5435 / 9999))
    If gv000E = False Then
        txtWarp.Text = Str$(hsbWarp.Value / 1000)
    End If
    gv000E = False
    subWARP
    Speed = (txtWarp.Text) * (-1)
    K = 2038
    Zoom = 256
    Timer1.Interval = 1


    For i = 0 To 100
        X(i) = Int(Rnd * 1024) - 512
        Y(i) = Int(Rnd * 1024) - 512
        Z(i) = Int(Rnd * 512) - 256
    Next i
    
    
    
End Sub

Sub hsbWarp_Scroll()
    txtWarp.Text = Str$(hsbWarp.Value / 1000)
End Sub

Private Sub txtWarp_Change()
Dim KeyAscii As Integer
Dim l0060 As Variant
' Const mc0064 = 13 ' &HD%
If KeyAscii = mc0064 Then
    KeyAscii = 0
    If Val(txtWarp.Text) < 0 Then
        txtWarp.Text = 0
    End If
    If Val(txtWarp.Text) > 9.9999 Then
        txtWarp.Text = 9.9999
    End If
    gv000E = True
    l0060 = 1000 * Val(txtWarp.Text)
    If l0060 > 9999 Then
        hsbWarp.Value = 9999
    Else
        hsbWarp.Value = l0060
    End If
    subWARP
End If
End Sub


Private Sub Form_Activate()
    Speed = 0
    K = 2038
    Zoom = 256
    Timer1.Interval = 1


    For i = 0 To 100
        X(i) = Int(Rnd * 1024) - 512
        Y(i) = Int(Rnd * 1024) - 512
        Z(i) = Int(Rnd * 512) - 256
    Next i
End Sub


Private Sub Timer1_Timer()


    For i = 0 To 100
        Circle (tmpX(i), tmpY(i)), 5, BackColor
        Z(i) = Z(i) + Speed
        If Z(i) > 255 Then Z(i) = -255
        If Z(i) < -255 Then Z(i) = 255
        tmpZ(i) = Z(i) + Zoom
        tmpX(i) = (X(i) * K / tmpZ(i)) + (10680)
        tmpY(i) = (Y(i) * K / tmpZ(i)) + (2040)
        Radius = 1
        StarColor = 256 - Z(i)
        Circle (tmpX(i), tmpY(i)), 5, RGB(StarColor, StarColor, StarColor)
    Next i
End Sub

