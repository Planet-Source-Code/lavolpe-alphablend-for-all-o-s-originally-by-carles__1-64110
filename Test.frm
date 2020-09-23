VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form fTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mAlphaBlt test"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9465
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   631
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   8055
      Top             =   5925
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select Test Image"
      Filter          =   "Bitmaps|*.bmp|GIFs|*.gif|JPGs|*.jpg;*.jpeg"
   End
   Begin VB.OptionButton optImage 
      Caption         =   "Any bit depth from file"
      Height          =   225
      Index           =   4
      Left            =   7320
      TabIndex        =   24
      Top             =   1575
      Width           =   1995
   End
   Begin VB.OptionButton optImage 
      Caption         =   "Linux Penguin"
      Height          =   225
      Index           =   3
      Left            =   7320
      TabIndex        =   23
      Top             =   1305
      Width           =   1995
   End
   Begin VB.OptionButton optImage 
      Caption         =   "Red Brush"
      Height          =   225
      Index           =   2
      Left            =   7320
      TabIndex        =   22
      Top             =   1035
      Width           =   1995
   End
   Begin VB.OptionButton optImage 
      Caption         =   "Ice"
      Height          =   225
      Index           =   1
      Left            =   7320
      TabIndex        =   21
      Top             =   780
      Width           =   1995
   End
   Begin VB.OptionButton optImage 
      Caption         =   "Tucan"
      Height          =   225
      Index           =   0
      Left            =   7320
      TabIndex        =   20
      Top             =   510
      Width           =   1995
   End
   Begin VB.TextBox txtOffsetX 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   3
      Left            =   8100
      TabIndex        =   4
      Text            =   "0"
      Top             =   3180
      Width           =   585
   End
   Begin VB.TextBox txtOffsetY 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   3
      Left            =   8760
      TabIndex        =   5
      Text            =   "0"
      Top             =   3180
      Width           =   585
   End
   Begin VB.TextBox txtOffsetY 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   2
      Left            =   8760
      TabIndex        =   9
      Text            =   "0"
      Top             =   4065
      Width           =   585
   End
   Begin VB.TextBox txtOffsetX 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   2
      Left            =   8100
      TabIndex        =   8
      Text            =   "0"
      Top             =   4065
      Width           =   585
   End
   Begin VB.TextBox txtOffsetX 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   1
      Left            =   8100
      TabIndex        =   6
      Text            =   "0"
      Top             =   3735
      Width           =   585
   End
   Begin VB.TextBox txtOffsetY 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   1
      Left            =   8760
      TabIndex        =   7
      Text            =   "0"
      Top             =   3735
      Width           =   585
   End
   Begin VB.TextBox txtOffsetY 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   0
      Left            =   8760
      TabIndex        =   3
      Text            =   "0"
      Top             =   2850
      Width           =   585
   End
   Begin VB.TextBox txtOffsetX 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   0
      Left            =   8100
      TabIndex        =   2
      Text            =   "0"
      Top             =   2850
      Width           =   585
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   135
      ScaleHeight     =   7065
      ScaleWidth      =   6945
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   165
      Width           =   6975
   End
   Begin VB.CheckBox chkAutoRedraw 
      Caption         =   "AutoRedraw Active?"
      Height          =   300
      Left            =   7380
      TabIndex        =   10
      Top             =   5250
      Width           =   1950
   End
   Begin VB.ComboBox cboGlobal 
      Height          =   315
      Left            =   8385
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2385
      Width           =   960
   End
   Begin VB.ComboBox cbScale 
      Height          =   315
      ItemData        =   "Test.frx":0000
      Left            =   8385
      List            =   "Test.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2025
      Width           =   960
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "&Render"
      Default         =   -1  'True
      Height          =   435
      Left            =   7335
      TabIndex        =   11
      Top             =   6795
      Width           =   2025
   End
   Begin VB.Label Label2 
      Caption         =   "Global Alpha"
      Enabled         =   0   'False
      Height          =   240
      Left            =   7365
      TabIndex        =   26
      Top             =   2445
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "When off-screen DCs or AutoRedraw are used, scaling can be much faster."
      Height          =   585
      Index           =   4
      Left            =   7350
      TabIndex        =   25
      Top             =   4575
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "IMG X,Y"
      Height          =   270
      Index           =   3
      Left            =   7335
      TabIndex        =   19
      Top             =   3795
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "IMG W,H"
      Height          =   270
      Index           =   2
      Left            =   7335
      TabIndex        =   18
      Top             =   4125
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Dest W,H"
      Height          =   270
      Index           =   1
      Left            =   7335
      TabIndex        =   17
      Top             =   3255
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Dest X,Y"
      Height          =   270
      Index           =   0
      Left            =   7335
      TabIndex        =   16
      Top             =   2910
      Width           =   690
   End
   Begin VB.Label lblScale 
      Caption         =   "Scale to"
      Enabled         =   0   'False
      Height          =   240
      Left            =   7380
      TabIndex        =   13
      Top             =   2085
      Width           =   690
   End
   Begin VB.Label lblTiming 
      Height          =   840
      Left            =   7410
      TabIndex        =   14
      Top             =   5820
      Width           =   1905
   End
   Begin VB.Label lblTestFunction 
      Caption         =   "Test Image (32 Bit):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7350
      TabIndex        =   12
      Top             =   210
      Width           =   1740
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' this is a scaled down version of the original project posted by Carles P.V.
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=60424&lngWId=1

' See mAlphaBlt module for an overview of the modifications

Private m_oDIB32 As cDIB32
Private m_oTile  As cTile
Private m_oT     As cTiming

' Note: following are declared here. If I were to modify the cDIB32 class to create
' 32bit DIBs from non-32bit sources, we wouldn't need to track these values
' separately cause the class would have them for our references.
Private hSource As Long
Private imgHeight As Long
Private imgWidth As Long
Private samplePic As StdPicture

Private Sub cboGlobal_Click()
    Call cmdRender_Click
End Sub

Private Sub cbScale_Click()
        
    '-- Scale to
    If Left$(cbScale.Text, 1) = "1" Then ' reducing
        txtOffsetX(3) = imgWidth \ Val(Right$(cbScale.Text, 1))
        txtOffsetY(3) = imgHeight \ Val(Right$(cbScale.Text, 1))
    Else    ' enlarging
        txtOffsetX(3) = imgWidth * Val(Left$(cbScale.Text, 1))
        txtOffsetY(3) = imgHeight * Val(Left$(cbScale.Text, 1))
    End If
    
    ' reset source image Width/Height to use as full width/height
    txtOffsetX(2) = imgWidth
    txtOffsetY(2) = imgHeight
    
    Call cmdRender_Click
    
End Sub

Private Sub chkAutoRedraw_Click()
    Picture1.AutoRedraw = (chkAutoRedraw.Value = 1)
End Sub

Private Sub Form_Load()

    Set Me.Icon = Nothing
    Show
    If (App.LogMode <> 1) Then
        Call MsgBox("Absolutely recommended: compile first...")
    End If
    
    
    Set m_oDIB32 = New cDIB32
    Set m_oTile = New cTile
    Set m_oT = New cTiming
    
    With cbScale
        Call .AddItem("1:5")
        Call .AddItem("1:4")
        Call .AddItem("1:3")
        Call .AddItem("1:2")
        Call .AddItem("1:1")
        Call .AddItem("2:1")
        Call .AddItem("3:1")
        Call .AddItem("4:1")
        Call .AddItem("5:1")
        Let .ListIndex = 4
    End With
    
    chkAutoRedraw = 1
    
    Dim I As Integer
    
    '-- Background
    On Error Resume Next
    If m_oTile.CreatePatternFromStdPicture(LoadPicture(App.Path & "\bkdrop.bmp")) Then
        If Err Then
            Err.Clear
            Set m_oTile = Nothing
        End If
    End If
    
    For I = 5 To 255 Step 20
        cboGlobal.AddItem I
    Next
    cboGlobal.AddItem "255"
    cboGlobal.ListIndex = cboGlobal.ListCount - 1
    optImage(2) = True
    
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Call Unload(Me)
End Sub

Private Sub mnuAbout_Click()
    Call MsgBox("Rendering alpha bitmaps" & vbCrLf & _
                "Carles P.V. - 2005" & vbCrLf & vbCrLf & _
                "Thanks to Peter Scale for 'Bilinear resizing' routine." & vbCrLf & _
                "Thanks to Ron van Tilburg for the 'integer maths' version" & vbCrLf & _
                "as well as for the 'Nearest neighbour resizing' routine.")
End Sub

Private Sub cmdRender_Click()
    
    Dim s  As String
      
    If hSource = 0 Then Exit Sub
    Picture1.Cls
            
    '-- Background
    If Not m_oTile Is Nothing Then Call m_oTile.Tile(Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height)
    
    '-- Render...
    Screen.MousePointer = vbArrowHourglass
    
    s = imgWidth & "x" & imgHeight & "-bitmap "
    
    ' following are optional parameters, but fill in so user knows the actual width/height
    If txtOffsetX(2) = 0 Then txtOffsetX(2) = imgWidth
    If txtOffsetY(2) = 0 Then txtOffsetY(2) = imgHeight
    
    Call m_oT.Reset
    
    If AlphaBlendStretch(Picture1.hDC, Val(txtOffsetX(0)), Val(txtOffsetY(0)), _
        Val(txtOffsetX(3)), Val(txtOffsetY(3)), _
        hSource, Val(txtOffsetX(1)), Val(txtOffsetY(1)), _
        Val(txtOffsetX(2)), Val(txtOffsetY(2)), Val(cboGlobal.Text)) = 0 Then
        
            ' error ocurred
            Screen.MousePointer = vbDefault
            
            MsgBox "Couldn't alpha blend the selected image or image is shifted off the screen by your settings." & vbCrLf & _
                "This sample project isn't written to support alpha blending cursors & icons" & vbCrLf & _
                "however; it wouldn't require much more effort.", vbInformation + vbOKOnly
            
    Else
    
        s = s & "scaled and rendered in "
    
'        Debug.Print s & Format$(m_oT.Elapsed / 1000, "0.000 sec.")
        lblTiming.Caption = s & Format$(m_oT.Elapsed / 1000, "0.000 sec.")
    
        Screen.MousePointer = vbDefault
        
    End If
    
End Sub

Private Sub optImage_Click(Index As Integer)
    
    Dim fn As String
    
    Select Case True
    Case optImage(0)
        fn = App.Path & "\tucan32.bmp"
    Case optImage(1)
        fn = App.Path & "\ice32.bmp"
    Case optImage(2)
        fn = App.Path & "\redbrush32.bmp"
    Case optImage(3)
        fn = App.Path & "\linux32.bmp"
    Case Else
        ' most likely won't have alpha values, but proves we can blend any bit depth
        optImage(4) = False
        On Error GoTo ExitRoutine
        dlgFile.Flags = cdlOFNFileMustExist
        dlgFile.ShowOpen
        
        On Error Resume Next
        Set samplePic = LoadPicture(dlgFile.Filename)
        If Err Then
            MsgBox "Invalid picture type, try another.", vbInformation + vbOKOnly
            If Err Then Err.Clear
            Exit Sub
        End If
        If m_oDIB32 Is Nothing Then m_oDIB32.Destroy
        hSource = samplePic.handle
        imgWidth = ScaleX(samplePic.Width, vbHimetric, vbPixels)
        imgHeight = ScaleY(samplePic.Height, vbHimetric, vbPixels)
    End Select
    
    If Len(fn) Then
        If m_oDIB32 Is Nothing Then Set m_oDIB32 = New cDIB32
        hSource = m_oDIB32.CreateFromBitmapFile(fn)
        If hSource = 0 Then
            MsgBox "Unzip the *.bmp files provided for this project.", vbInformation + vbOKOnly
        Else
            imgHeight = m_oDIB32.Height
            imgWidth = m_oDIB32.Width
            Set samplePic = Nothing
        End If
    End If
    
    ' reset source image X,Y offset to 0,0
    txtOffsetX(1) = 0: txtOffsetY(1) = 0
    
    Call cbScale_Click
    
ExitRoutine:
End Sub

Private Sub txtOffsetX_GotFocus(Index As Integer)
    With txtOffsetX(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtOffsetY_GotFocus(Index As Integer)
    With txtOffsetY(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
