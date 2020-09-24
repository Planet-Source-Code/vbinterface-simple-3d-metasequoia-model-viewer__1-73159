VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "    3D MQO Model Viewer [basic version]"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framAbout 
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      Height          =   5235
      Left            =   1590
      TabIndex        =   19
      Top             =   450
      Visible         =   0   'False
      Width           =   6795
      Begin VB.Label lblHomePage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "http://bytelogik.wordpress.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4290
         TabIndex        =   26
         Top             =   3780
         Width           =   2325
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblEMail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "vbinterface@gmail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   4290
         TabIndex        =   25
         Top             =   3990
         Width           =   2325
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTwitter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "http://twitter.com/bytelogik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4290
         TabIndex        =   24
         Top             =   3570
         Width           =   2325
      End
      Begin VB.Label lblBottom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "join me :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   4290
         TabIndex        =   23
         Top             =   3270
         Width           =   2325
      End
      Begin VB.Label lblMidTwo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "so friends, never give up. even if the obstacles are innumerable and difficult one."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   555
         Left            =   4290
         TabIndex        =   22
         Top             =   2040
         Width           =   2325
      End
      Begin VB.Label lblMid 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "my mind replies : yes."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4290
         TabIndex        =   21
         Top             =   1620
         Width           =   2325
      End
      Begin VB.Image imgAbout 
         Height          =   3480
         Left            =   180
         Picture         =   "main.frx":0442
         Top             =   720
         Width           =   3885
      End
      Begin VB.Label lblTop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "very often, when i look into the depths of universe, filled with planets and stars, i ask to myself : can i explore all of them ?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   825
         Left            =   4290
         TabIndex        =   20
         Top             =   720
         Width           =   2325
      End
      Begin VB.Image imgBtnAboutClose 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2880
         Picture         =   "main.frx":2C764
         Top             =   4590
         Width           =   1080
      End
   End
   Begin VB.PictureBox picLeaves 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000C0C0&
      Height          =   1125
      Left            =   7620
      Picture         =   "main.frx":2DB0E
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   4
      Top             =   4920
      Width           =   2340
      Begin VB.Label lblContact 
         BackStyle       =   0  'Transparent
         Caption         =   "vbinterface@gmail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   885
         Width           =   2130
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00444444&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5400
      Left            =   120
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   360
      Width           =   6000
      Begin VB.Image imgBtnAbout 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   0
         Picture         =   "main.frx":32C4C
         Top             =   2220
         Width           =   375
      End
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Read more at http://bytelogik.wordpress.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6240
      TabIndex        =   18
      Top             =   1440
      Width           =   3555
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "A simple Metasequoia 3D model viewer."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6240
      TabIndex        =   17
      Top             =   330
      Width           =   3555
   End
   Begin VB.Label Label12 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   $"main.frx":33BB2
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F9F9F&
      Height          =   765
      Left            =   6240
      TabIndex        =   16
      Top             =   570
      Width           =   3555
   End
   Begin VB.Label lblModelName 
      BackStyle       =   0  'Transparent
      Caption         =   " model name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   90
      TabIndex        =   15
      Top             =   120
      Width           =   5985
   End
   Begin VB.Label Label11 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "LightY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F9F9F&
      Height          =   195
      Left            =   6300
      TabIndex        =   14
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "LightX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F9F9F&
      Height          =   195
      Left            =   6300
      TabIndex        =   13
      Top             =   3900
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "Scale"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F9F9F&
      Height          =   195
      Left            =   6300
      TabIndex        =   12
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "TraY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F9F9F&
      Height          =   195
      Left            =   6300
      TabIndex        =   11
      Top             =   3300
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "TraX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F9F9F&
      Height          =   195
      Left            =   6300
      TabIndex        =   10
      Top             =   3030
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "RotZ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F9F9F&
      Height          =   195
      Left            =   6300
      TabIndex        =   9
      Top             =   2730
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "RotY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F9F9F&
      Height          =   195
      Left            =   6300
      TabIndex        =   8
      Top             =   2490
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "RotX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F9F9F&
      Height          =   195
      Left            =   6300
      TabIndex        =   7
      Top             =   2250
      Width           =   495
   End
   Begin VB.Image imgBt 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   7
      Left            =   8040
      Picture         =   "main.frx":33C57
      Top             =   4170
      Width           =   285
   End
   Begin VB.Image imgBt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   8040
      Picture         =   "main.frx":340D1
      Top             =   3870
      Width           =   300
   End
   Begin VB.Image imgBt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   8040
      Picture         =   "main.frx":34587
      Top             =   3540
      Width           =   300
   End
   Begin VB.Image imgBt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   8040
      Picture         =   "main.frx":34A3D
      Top             =   3240
      Width           =   300
   End
   Begin VB.Image imgBt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   8040
      Picture         =   "main.frx":34EF3
      Top             =   2970
      Width           =   300
   End
   Begin VB.Image imgBt 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   0
      Left            =   7620
      Picture         =   "main.frx":353A9
      Top             =   2190
      Width           =   270
   End
   Begin VB.Image imgBt 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   2
      Left            =   8040
      Picture         =   "main.frx":357A3
      Top             =   2670
      Width           =   270
   End
   Begin VB.Image imgBt 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   1
      Left            =   7830
      Picture         =   "main.frx":35B9D
      Top             =   2430
      Width           =   270
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Transformation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006DDACA&
      Height          =   225
      Left            =   7290
      TabIndex        =   6
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00505050&
      Height          =   2805
      Left            =   6240
      Top             =   1890
      Width           =   3645
   End
   Begin VB.Label Label8 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "Coming Next :  Smooth Shading      "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0048D0BC&
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   5670
      Width           =   1305
   End
   Begin VB.Label Label9 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous :       Illumination in 3D."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F9F9F&
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   5040
      Width           =   1305
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   " status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   5820
      Width           =   5985
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00464646&
      X1              =   458
      X2              =   650
      Y1              =   154
      Y2              =   154
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00464646&
      X1              =   458
      X2              =   650
      Y1              =   170
      Y2              =   170
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00464646&
      X1              =   458
      X2              =   650
      Y1              =   186
      Y2              =   186
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00464646&
      X1              =   458
      X2              =   650
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00464646&
      X1              =   458
      X2              =   650
      Y1              =   226
      Y2              =   226
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00464646&
      X1              =   458
      X2              =   650
      Y1              =   246
      Y2              =   246
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00464646&
      X1              =   458
      X2              =   650
      Y1              =   266
      Y2              =   266
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00464646&
      X1              =   458
      X2              =   650
      Y1              =   286
      Y2              =   286
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------
'
'Logic/Code/Graphics only for "Personal" use
'Please contact vbinterface@gmail.com for any commercial usage
'
'---------------------------------------------------------------------------------------------
'Triangulated models only
'No texture/material/shadow/smooth shading support
'Single Light
'
'Included 4 models in the zip file
'More 3D Metasequoia models at http://bytelogik.wordpress.com
'---------------------------------------------------------------------------------------------

Option Explicit
    Private Type POINTAPI
        X As Long
        Y As Long
    End Type
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Dim Pt As POINTAPI
    Dim ScreenStartX As Long, ScreenStartY As Long
    Dim ScreenEndX As Long, ScreenEndY As Long
    Dim dX As Long, dY As Long
    Dim imgLeft As Long, imgTop As Long
    Dim StartMove As Byte
    Dim LimitMin As Long, LimitMax As Long, CurrentPos As Long
    '-----------------------------------------------------------------------------------------
    Dim tX As Single, tY As Single, tZ As Single, ScaleV As Single
    Dim Angle(2) As Single, Trans(1) As Single, LightAngle(1) As Single
    Dim LightMAT As D3DMATRIX
    Dim LXR As Single, LYR As Single
Private Sub Form_Load()
    sW = pic.ScaleWidth: sH = pic.ScaleHeight
    CenX = sW \ 2: CenY = sH \ 2
    LimitMin = 449: LimitMax = 631
    '------------------------------------
    'lblStatus = LoadMQO(App.Path & "\hqsphere.mqo", pic, False): lblModelName = "sphere"
    'lblStatus = LoadMQO(App.Path & "\torus.mqo", pic, False): lblModelName = "torus"
    'lblStatus = LoadMQO(App.Path & "\planet.mqo", pic, False): lblModelName = "planet"
    lblStatus = LoadMQO(App.Path & "\mug.mqo", pic, False): lblModelName = "mug"
    '------------------------------------
    LoadLights
    D3DXMatrixIdentity ScaleMAT
    Angle(0) = 180: Angle(1) = 90: Angle(2) = 0
    imgBt(0).Left = (Angle(0) \ 2) + 450
    imgBt(1).Left = (Angle(1) \ 2) + 450
    imgBt(2).Left = (Angle(2) \ 2) + 450
    '------------------------------------
    imgBt(3).Left = 540: imgBt(4).Left = 540: imgBt(5).Left = 540
    '------------------------------------
    imgBt(6).Left = 540: imgBt(7).Left = 540
    '------------------------------------
    RotateWorld Angle(0), Angle(1), Angle(2)
End Sub

Private Sub imgBt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetCursorPos Pt
    ScreenStartX = Pt.X
    imgLeft = imgBt(Index).Left
    StartMove = 1
End Sub
Private Sub imgBt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'navigation
    If StartMove = 1 Then
        GetCursorPos Pt
        ScreenEndX = Pt.X
        dX = ScreenEndX - ScreenStartX
        CurrentPos = imgLeft + dX
        If CurrentPos > LimitMin And CurrentPos < LimitMax Then
            imgBt(Index).Left = CurrentPos
            If Index < 3 Then '----rotate
                Angle(Index) = (CurrentPos - 450) * 2
                RotateWorld Angle(0), Angle(1), Angle(2)
            ElseIf Index = 3 Or Index = 4 Then  '----translate
                Trans(Index - 3) = ((CurrentPos - 450) - 90) * 4
                RotateWorld Angle(0), Angle(1), Angle(2)
            ElseIf Index = 5 Then '----scale
                ScaleV = 1 + (((CurrentPos - 450) - 90) / 100)
                ScaleMAT.m11 = ScaleV
                ScaleMAT.m22 = ScaleV
                ScaleMAT.m33 = ScaleV
                RotateWorld Angle(0), Angle(1), Angle(2)
            Else '----light rotation
                LightAngle(Index - 6) = ((CurrentPos - 450) - 90) * PIBY180
                D3DXMatrixRotationYawPitchRoll LightMAT, LightAngle(1), LightAngle(0), 0
                D3DXVec4Transform tLightVect, LightVect, LightMAT
                D3DXVec4Normalize NormLight, tLightVect
                pic.Cls
                RenderWorld pic
                pic.Refresh
                DoEvents
            End If
        End If
    End If
End Sub
Private Sub imgBt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    StartMove = 0
End Sub
Sub RotateWorld(ByVal XR As Long, ByVal YR As Long, ByVal ZR As Long)
    TransformWorld XR * PIBY180, YR * PIBY180, ZR * PIBY180, Trans(0), Trans(1), 0
    pic.Cls
    RenderWorld pic
    pic.Refresh
    DoEvents
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ClearObject
    Erase Angle
    Erase Trans
    Erase LightAngle
End Sub
Private Sub imgBtnAbout_Click()
    If framAbout.Visible = False Then
        framAbout.Visible = True
        framAbout.ZOrder 0
    End If
End Sub
Private Sub imgBtnAboutClose_Click()
    framAbout.Visible = False
End Sub
