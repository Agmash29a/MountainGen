VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MAIN_MountainGen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mountain Generator"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MAIN_MountainGen.frx":0000
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   775
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TEXTBOX_RocksBlue 
      Height          =   285
      Left            =   11160
      TabIndex        =   53
      Text            =   "Text9"
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox TEXTBOX_RocksGreen 
      Height          =   285
      Left            =   10680
      TabIndex        =   52
      Text            =   "Text8"
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox TEXTBOX_RocksRed 
      Height          =   285
      Left            =   10200
      TabIndex        =   51
      Text            =   "Text7"
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox TEXTBOX_SnowRed 
      Height          =   285
      Left            =   10200
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox TEXTBOX_SnowGreen 
      Height          =   285
      Left            =   10680
      TabIndex        =   45
      Text            =   "Text2"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox TEXTBOX_SnowBlue 
      Height          =   285
      Left            =   11160
      TabIndex        =   44
      Text            =   "Text3"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox TEXTBOX_VegitationRed 
      Height          =   285
      Left            =   10200
      TabIndex        =   43
      Text            =   "Text4"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox TEXTBOX_VegitationGreen 
      Height          =   285
      Left            =   10680
      TabIndex        =   42
      Text            =   "Text5"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox TEXTBOX_VegitationBlue 
      Height          =   285
      Left            =   11160
      TabIndex        =   41
      Text            =   "Text6"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox TEXTBOX_WaterRed 
      Height          =   285
      Left            =   10200
      TabIndex        =   40
      Text            =   "Text7"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox TEXTBOX_WaterGreen 
      Height          =   285
      Left            =   10680
      TabIndex        =   39
      Text            =   "Text8"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox TEXTBOX_WaterBlue 
      Height          =   285
      Left            =   11160
      TabIndex        =   38
      Text            =   "Text9"
      Top             =   4080
      Width           =   375
   End
   Begin VB.Frame FRAME_LightOrigin 
      Caption         =   "Light Origin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   31
      Top             =   1920
      Width           =   3015
      Begin VB.CommandButton BUTTON_LightOriginzPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   2640
         TabIndex        =   37
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_LightOriginzMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         TabIndex        =   36
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_LightOriginyMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   35
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_LightOriginyPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_LightOriginxPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_LightOriginxMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame FRAME_LookatPoint 
      Caption         =   "Lookat Point (x,y,z)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   24
      Top             =   1200
      Width           =   3015
      Begin VB.CommandButton BUTTON_LookatzPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_LookatzMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_LookatyMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_LookatyPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_LookatxPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_LookatxMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame FRAME_ViewCoords 
      Caption         =   "Viewing Coordinates (x,y,z)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8520
      TabIndex        =   13
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton BUTTON_dPlus 
         Caption         =   "d +"
         Height          =   255
         Left            =   2520
         TabIndex        =   23
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton BUTTON_dMinus 
         Caption         =   "d -"
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton BUTTON_TwistAnglePlus 
         Caption         =   "Twist +"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton BUTTON_TwistAngleMinus 
         Caption         =   "Twist -"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton BUTTON_zPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_zMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_yPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_yMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_xPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BUTTON_xMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame FRAME_DetailLevel 
      Caption         =   "Detail level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
      Begin VB.TextBox TEXTBOX_DetailLevel 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FRAME_kadr 
      Caption         =   "kadr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   9
      Top             =   5520
      Width           =   1455
      Begin VB.TextBox TEXTBOX_kadr 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox CHECKBOX_Water 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      Top             =   6240
      Width           =   255
   End
   Begin VB.CommandButton BUTTON_ToggleWireframe 
      Caption         =   "Toggle Wireframe"
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton BUTTON_ToggleAxis 
      Caption         =   "Toggle Axis"
      Height          =   375
      Left            =   10080
      TabIndex        =   5
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton BUTTON_NewLandscape 
      BackColor       =   &H00C0FFC0&
      Caption         =   "New Landscape"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   3015
   End
   Begin VB.PictureBox PICTUREBOX_Main 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      Height          =   7695
      Left            =   -120
      Picture         =   "MAIN_MountainGen.frx":CA04A
      ScaleHeight     =   509
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   565
      TabIndex        =   3
      Top             =   0
      Width           =   8535
      Begin MSComDlg.CommonDialog DIALOG_Save 
         Left            =   360
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label LABEL_TriangleCount 
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   50
         Top             =   7320
         Width           =   2775
      End
   End
   Begin VB.CommandButton BUTTON_Save 
      Caption         =   "Save Image"
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   6600
      Width           =   3015
   End
   Begin VB.Frame FRAME_ChangeByAmount 
      Caption         =   "+ or - amount /  Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   0
      Top             =   2640
      Width           =   3015
      Begin VB.TextBox TEXTBOX_ChangeByAmount 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "0"
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Label LABEL_RockColor 
      Caption         =   "Rocks (RGB)"
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
      Left            =   8520
      TabIndex        =   54
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label LABEL_SnowColor 
      Caption         =   "Snow (RGB)"
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
      Left            =   8520
      TabIndex        =   49
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LABEL_VegColor 
      Caption         =   "Vegetation (RGB)"
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
      Left            =   8520
      TabIndex        =   48
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label LABEL_WaterColor 
      Caption         =   "Water (RGB)"
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
      Left            =   8520
      TabIndex        =   47
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label LABEL_ShowWater 
      Caption         =   "Show Water"
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
      Left            =   8880
      TabIndex        =   8
      Top             =   6240
      Width           =   2175
   End
End
Attribute VB_Name = "MAIN_MountainGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'000000000000000000000000000000000000000000000000000000000000000000000
'000000000000000000011000000000011000000000000000000000000000000000000
'000000011111100000101000110000111000011000000000001111110000000000000
'000000100001100001010000110001110000110000000000001000111100000000000
'000011000001100010100001110011010000110000000000001100001100000000000
'000011000011000111000000100010100001110000001110001100111000010000000
'000110000011001110100000100011000000100000110100001111000001111000000
'000110000000001101100111100110000101100000110001111100000010011011000
'000011000000111110111001111001111001100011111110011111001100010100000
'000011111111001100000001100001100001111110001000010001110000111000000
'000000011000000000000000000000000001111000000000010000011100111000000
'000000000000000000000000000000000000000000000000000000000001011000000
'000000000000000000000000000000000000000000000000000000000010110000000
'(c) 2002 by Riley T. Perry - Chillers of Entropy

'-> If the comments below look garbled then change font to COURIER NEW

'                                                 ,  ,
'                                                / \/ \
'                                              (/ //_ \_
'     .-._                                      \||  .  \
'      \  '-._                            _,:__.-"/---\_ \
' ______/___  '.    .--------------------'~-'--.)__( , )\ \
'`'--.___  _\  /    | Main Form               ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 22/5/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * The main form for the application.                     *
' *--------------------------------------------------------*

'Option Explicit --> removed for file flags
'
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.                                          |
' |  /   \          Types and variables            |
' | |\_.  |         -------------------            |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'

'**** Type for imported function ****

Private Type CornerRec
  x As Long
  y As Long
End Type

'**** Imported functions ****

Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As CornerRec, ByVal nCount As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoints As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal rgbColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hndobj As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRegion As Long, ByVal hBrush As Long) As Long
'
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   Form_Load()                            |
' |  /   \  -----------                            |
' | |\_.  | Start the application                  |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Private Sub Form_Load()

    '*------------------------------------------*
    '*          Declare and initialise          *
    '*------------------------------------------*

    '**** x,y, and z axis' ****
    
    Axis_ystart.x = 0
    Axis_ystart.y = -700
    Axis_ystart.z = 0
    
    Axis_yfinish.x = 0
    Axis_yfinish.y = 700
    Axis_yfinish.z = 0
    
    Axis_xstart.x = -700
    Axis_xstart.y = 0
    Axis_xstart.z = 0
    
    Axis_xfinish.x = 700
    Axis_xfinish.y = 0
    Axis_xfinish.z = 0
    
    Axis_zstart.x = 0
    Axis_zstart.y = 0
    Axis_zstart.z = -700
    
    Axis_zfinish.x = 0
    Axis_zfinish.y = 0
    Axis_zfinish.z = 700

    '**** Powers of 2 (for speed) ****

    Powers2(0) = 1
    Powers2(1) = 2
    Powers2(2) = 4
    Powers2(3) = 8
    Powers2(4) = 16
    Powers2(5) = 32
    Powers2(6) = 64
    Powers2(7) = 128
    Powers2(8) = 256
    Powers2(9) = 512
    Powers2(10) = 1024
    Powers2(11) = 2048
    Powers2(12) = 4096
    Powers2(13) = 8192
    Powers2(14) = 16384
    Powers2(15) = 32768
     
    '**** Misc ****
     
    TwistAngle = 1

    kadr = 0.75
    
    d_Scale = 1500
    
    DetailLevel = 4
    
    '**** Starting view coords ****
    
    ViewCoords_x = 2000
    ViewCoords_y = 2000
    ViewCoords_z = 2000
    
    '**** Start by looking at origin ****
    
    Lookat_x = 0
    Lookat_y = 0
    Lookat_z = 0
    
    '**** The sun is close! ****
    
    L.x = 0
    L.y = 400
    L.z = 20000
   
    '**** Frames ****
    
    Wireframe = False
    ShowAxis = True
    ShowWater = True
    
    '**** Resolution / 2 ****
        
    ResolutionDisplacement_x = 732 / 2
    ResolutionDisplacement_y = 550 / 2
    
    '**** Colors ****
    
    Snow_Red = 125
    Snow_Green = 125
    Snow_Blue = 125

    Vegitation_Red = 0
    Vegitation_Green = 125
    Vegitation_Blue = 0

    Water_Red = 0
    Water_Green = 0
    Water_Blue = 125
    
    Rocks_Red = 50
    Rocks_Green = 50
    Rocks_Blue = 50
    
    '**** Fill text boxes ****
    
    TEXTBOX_kadr.Text = CStr(kadr)
    TEXTBOX_DetailLevel.Text = CStr(DetailLevel)
    
    TEXTBOX_SnowRed.Text = CStr(Snow_Red)
    TEXTBOX_SnowGreen.Text = CStr(Snow_Green)
    TEXTBOX_SnowBlue.Text = CStr(Snow_Blue)

    TEXTBOX_VegitationRed.Text = CStr(Vegitation_Red)
    TEXTBOX_VegitationGreen.Text = CStr(Vegitation_Green)
    TEXTBOX_VegitationBlue.Text = CStr(Vegitation_Blue)

    TEXTBOX_WaterRed.Text = CStr(Water_Red)
    TEXTBOX_WaterGreen.Text = CStr(Water_Green)
    TEXTBOX_WaterBlue.Text = CStr(Water_Blue)
    
    TEXTBOX_RocksRed.Text = CStr(Rocks_Red)
    TEXTBOX_RocksGreen.Text = CStr(Rocks_Green)
    TEXTBOX_RocksBlue.Text = CStr(Rocks_Blue)
    
    CHECKBOX_Water.value = Checked
    
    Call NewLandscape
    
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   NewLandscape()                         |
' |  /   \  --------------                         |
' | |\_.  | Generate a new landscape               |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Private Sub NewLandscape()

    '*------------------------------------------*
    '*          Declare and initialise          *
    '*------------------------------------------*
          
    ReDim Triangles(1)
          
    Dim S1() As PointType
    Dim S2() As PointType
    Dim S3() As PointType
  
    Dim TempMAXy As Double
    Dim I As Long
    
    '**** Sides of the triangle ****

    pp1.x = -600
    pp1.y = 0
    pp1.z = -600
    
    pp3.x = 0
    pp3.y = 0
    pp3.z = 600
    
    pp2.x = 600
    pp2.y = 0
    pp2.z = -600

    '**** Number of triangles ****
    
    TriangleCounter = 0
    
    '*------------------------------------------*
    '*            Generate Landscape            *
    '*------------------------------------------*
    
    '**** Start program ****
    
    Call Make_Sides(DetailLevel + 1, Powers2(DetailLevel), S1, S2, S3)
    Call MakeTriangle(DetailLevel, S1, S2, S3)
    
    If ShowWater Then
    
        Call MakeWater
    
    End If
    
    '**** Find Maximum height ****
    
    For I = 1 To UBound(Triangles)
        
        TempMAXy = GetMAXy(I)
        
        If MaximumHeight <= TempMAXy Then
        
            MaximumHeight = TempMAXy
        
        End If
    
    Next
    
    UnitsOfHeight = MaximumHeight / 90
    
    Call Render
  
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   Render()                               |
' |  /   \  --------                               |
' | |\_.  | Transform and draw landscape           |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Sub Render()

    '*------------------------------------------*
    '*          Declare and initialise          *
    '*------------------------------------------*
       
    Dim N As Vector3D
   
    Dim MAXz As Double
    Dim MAXy As Double
    Dim OLDz As Boolean
        
    Dim Brush As Long
    Dim ReturnCode As Long
    Dim Region As Long
    Dim BaseCol As Long
    
    Dim IRGB As Double
    
    Dim box(3) As CornerRec
    
    Dim FinalColor As Long
    
    Dim Basis_x As Vector3D
    
    Basis_x.x = 1
    Basis_x.y = 0
    Basis_x.z = 0
    
    Dim CosTheta As Double
    
    Dim xaxis(1) As CornerRec
    Dim yaxis(1) As CornerRec
    Dim zaxis(1) As CornerRec
   
    Dim I As Long
   
    Dim VegitationChance As Double
   
    '**** Wipe previous picture ****
   
    PICTUREBOX_Main.Cls

    '*------------------------------------------*
    '*       Transform and project axis'        *
    '*------------------------------------------*
    
    If ShowAxis Then
        
        '**** x axis ****
        
        Call ViewTransform(Axis_xstart, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
        Call ViewTransform(Axis_xfinish, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
    
        Call PerspectiveProject(d_Scale, Axis_xstart)
        Call PerspectiveProject(d_Scale, Axis_xfinish)
      
        '**** y axis ****
           
        Call ViewTransform(Axis_ystart, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
        Call ViewTransform(Axis_yfinish, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
    
        Call PerspectiveProject(d_Scale, Axis_ystart)
        Call PerspectiveProject(d_Scale, Axis_yfinish)
      
        '**** z axis ****
            
        Call ViewTransform(Axis_zstart, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
        Call ViewTransform(Axis_zfinish, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
    
        Call PerspectiveProject(d_Scale, Axis_zstart)
        Call PerspectiveProject(d_Scale, Axis_zfinish)
           
        PICTUREBOX_Main.Line (CInt(Axis_ystart.xp) + ResolutionDisplacement_x, CInt(Axis_ystart.yp) + ResolutionDisplacement_y)-(CInt(Axis_yfinish.xp) + ResolutionDisplacement_x, CInt(Axis_yfinish.yp) + ResolutionDisplacement_y), vbRed
        PICTUREBOX_Main.Line (CInt(Axis_xstart.xp) + ResolutionDisplacement_x, CInt(Axis_xstart.yp) + ResolutionDisplacement_y)-(CInt(Axis_xfinish.xp) + ResolutionDisplacement_x, CInt(Axis_xfinish.yp) + ResolutionDisplacement_y), vbBlue
        PICTUREBOX_Main.Line (CInt(Axis_zstart.xp) + ResolutionDisplacement_x, CInt(Axis_zstart.yp) + ResolutionDisplacement_y)-(CInt(Axis_zfinish.xp) + ResolutionDisplacement_x, CInt(Axis_zfinish.yp) + ResolutionDisplacement_y), vbGreen
      
    End If

    '*------------------------------------------*
    '*            Transform Triangles           *
    '*------------------------------------------*
        
    For I = 0 To UBound(Triangles)

        Call ViewTransform(Triangles(I).edge1.Start, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
        Call ViewTransform(Triangles(I).edge1.Finish, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
        Call ViewTransform(Triangles(I).edge2.Start, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
        Call ViewTransform(Triangles(I).edge2.Finish, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
      
        Call ViewTransform(Triangles(I).edge3.Start, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
        Call ViewTransform(Triangles(I).edge3.Finish, ViewCoords_x, ViewCoords_y, ViewCoords_z, Lookat_x, Lookat_y, Lookat_z, TwistAngle)
      
        Call PerspectiveProject(d_Scale, Triangles(I).edge1.Start)
        Call PerspectiveProject(d_Scale, Triangles(I).edge1.Finish)
      
        Call PerspectiveProject(d_Scale, Triangles(I).edge2.Start)
        Call PerspectiveProject(d_Scale, Triangles(I).edge2.Finish)
      
        Call PerspectiveProject(d_Scale, Triangles(I).edge3.Start)
        Call PerspectiveProject(d_Scale, Triangles(I).edge3.Finish)
     
        '**** Find light direction ****
     
        N = CrossProduct(SubtractVectors(Triangles(I).edge1.Start, Triangles(I).edge1.Finish), SubtractVectors(Triangles(I).edge2.Finish, Triangles(I).edge2.Start))
    
        Call NormaliseVector(N)
        Call NormaliseVector(L)
      
        Triangles(I).I = DotProduct(N, L)
      
        '**** Get max z coord for painting ****
      
        MAXz = Triangles(I).edge1.Finish.zv
    
        If MAXz < Triangles(I).edge2.Finish.zv Then
    
            MAXz = Triangles(I).edge2.Finish.zv
    
        End If
    
        If MAXz < Triangles(I).edge3.Finish.zv Then
    
          MAXz = Triangles(I).edge3.Finish.zv
    
        End If
    
        If MAXz < Triangles(I).edge1.Start.zv Then
    
            MAXz = Triangles(I).edge1.Start.zv
    
        End If
      
        Triangles(I).Middle_z = MAXz
      
    Next

    '**** Sort triangles in order to paint ****

    Call QSort_Numeric_Ascending(Triangles(), 0, UBound(Triangles))

    '*------------------------------------------*
    '*              Paint Triangles             *
    '*------------------------------------------*

    For I = UBound(Triangles) To 0 Step -1
        
        '**** Get slope of triangle ****
        
        N = CrossProduct(SubtractVectors(Triangles(I).edge1.Start, Triangles(I).edge1.Finish), SubtractVectors(Triangles(I).edge2.Finish, Triangles(I).edge2.Start))
        
        N.z = Abs(N.z)
        N.x = Abs(N.x)
        N.y = Abs(N.y)
        
        Call NormaliseVector(N)
        
        Basis_x.z = Abs(N.z)
        
        '**** Find max height of trianlge ****
        
        MAXy = GetMAXy(I)
        
        '**** Store in data structure to paint ****
  
        box(0).x = CInt(Triangles(I).edge1.Start.xp) + ResolutionDisplacement_x
        box(0).y = CInt(Triangles(I).edge1.Start.yp) + ResolutionDisplacement_y
  
        box(1).x = CInt(Triangles(I).edge2.Start.xp) + ResolutionDisplacement_x
        box(1).y = CInt(Triangles(I).edge2.Start.yp) + ResolutionDisplacement_y

        box(2).x = CInt(Triangles(I).edge3.Start.xp) + ResolutionDisplacement_x
        box(2).y = CInt(Triangles(I).edge3.Start.yp) + ResolutionDisplacement_y
  
        '**** find edge where z's are same ****
  
        If (((Triangles(I).edge3.Start.z - Triangles(I).edge3.Finish.z) = 0) And ((Triangles(I).edge3.Start.z - Triangles(I).edge3.Finish.z) = 0)) Then
          
            If Triangles(I).edge1.Finish.z > Triangles(I).edge1.Start.z Then
            
                OLDz = False
            
            Else
        
                OLDz = True
            
            End If
          
        End If
  
        If (((Triangles(I).edge2.Start.z - Triangles(I).edge2.Finish.z) = 0) And ((Triangles(I).edge2.Start.z - Triangles(I).edge2.Finish.z) = 0)) Then
  
            If Triangles(I).edge3.Finish.z > Triangles(I).edge3.Start.z Then
    
                OLDz = False
    
            Else

                OLDz = True
    
            End If
  
        End If
  
        If (((Triangles(I).edge1.Start.z - Triangles(I).edge1.Finish.z) = 0) And ((Triangles(I).edge1.Start.z - Triangles(I).edge1.Finish.z) = 0)) Then
  
            If Triangles(I).edge2.Finish.z > Triangles(I).edge2.Start.z Then
    
                OLDz = False
    
            Else

                OLDz = True
    
            End If
  
        End If
  
        '**** Negate coordinate system of triangle is oriented backwards ****
  
        If Not OLDz Then
  
            Triangles(I).I = -Triangles(I).I
    
        End If
  
        '**** No direct illumination ****
  
        If Triangles(I).I < 0 Then
  
            Triangles(I).I = 0
    
        End If

        '*------------------------------------------*
        '*             Determine Color              *
        '*------------------------------------------*
        
        '**** Determine color coefficient ****
        
        IRGB = kadr + kadr * Triangles(I).I
    
        '**** Color with Rocks value ****
        
        FinalColor = RGB(Rocks_Red * IRGB, Rocks_Green * IRGB, Rocks_Blue * IRGB)
        
        '**** Color with snow (first calculate slope) ****
        
        CosTheta = DotProduct(N, Basis_x)
        CosTheta = 90 - (ArcCos(CosTheta) * 180 / 3.141592654)
        
        '**** Bottom part of mountain has % chance of vegitation ****
        
        Randomize
        
        VegitationChance = (Rnd() * 100)
        
        '**** >>>> if statements for later use in vegitation mapping ****
        
        If (MAXy < (20 * UnitsOfHeight)) And (VegitationChance < 5) Then

            FinalColor = RGB(Vegitation_Red * IRGB, Vegitation_Green * IRGB, Vegitation_Blue * IRGB)
           
        End If
        
        If (MAXy < (15 * UnitsOfHeight)) And (VegitationChance < 20) Then

            FinalColor = RGB(Vegitation_Red * IRGB, Vegitation_Green * IRGB, Vegitation_Blue * IRGB)
           
        End If
        
        If MAXy < (5 * UnitsOfHeight) And (VegitationChance < 70) Then

            FinalColor = RGB(Vegitation_Red * IRGB, Vegitation_Green * IRGB, Vegitation_Blue * IRGB)
           
        End If
        
        '**** Top part of mountain always has snow ****
        
        If MAXy > (80 * UnitsOfHeight) Then

            FinalColor = RGB(Snow_Red * IRGB, Snow_Green * IRGB, Snow_Blue * IRGB)
           
        End If
        
        '**** Snow for other parts of terrain ****
       
        If (((MAXy * UnitsOfHeight) / 8) - CosTheta) > 0 Then
        
            FinalColor = RGB(Snow_Red * IRGB, Snow_Green * IRGB, Snow_Blue * IRGB)
        
        End If
       
        '**** Draw water ****
        
        If MAXy = 0 And ShowWater Then
        
            FinalColor = RGB(Water_Red * IRGB, Water_Green * IRGB, Water_Blue * IRGB)
        
        End If

        '*------------------------------------------*
        '*               Draw Triangle              *
        '*------------------------------------------*
             
        If Wireframe Then
            
            '**** Draw wireframe triangle ****
            
            PICTUREBOX_Main.Line (box(0).x, box(0).y)-(box(1).x, box(1).y)
            PICTUREBOX_Main.Line (box(1).x, box(1).y)-(box(2).x, box(2).y)
            PICTUREBOX_Main.Line (box(2).x, box(2).y)-(box(0).x, box(0).y)
            
        Else
            
            '**** Draw coloured traingle ****
            
            Brush = CreateSolidBrush(FinalColor)
            Region = CreatePolygonRgn(box(0), 3, 1)
            
            ReturnCode = FillRgn(PICTUREBOX_Main.hdc, Region, Brush)
            ReturnCode = FillRgn(PICTUREBOX_Main.hdc, Region, Brush)
            ReturnCode = DeleteObject(Region)
            ReturnCode = DeleteObject(Brush)
            
        End If

    Next
  
    LABEL_TriangleCount.Caption = CStr(TriangleCounter) & " Triangles"
    PICTUREBOX_Main.Refresh
    MAIN_MountainGen.Refresh
   
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   MakeWater()                            |
' |  /   \  -----------                            |
' | |\_.  | Convert negative values of y to water  |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Private Sub MakeWater()

    '*------------------------------------------*
    '* Check each point for a negative y value  *
    '*------------------------------------------*
    
    For I = 0 To UBound(Triangles)
        
        If Triangles(I).edge1.Start.y < 0 Then
        
            Triangles(I).edge1.Start.y = 0
        
        End If
        
        If Triangles(I).edge1.Finish.y < 0 Then
        
            Triangles(I).edge1.Finish.y = 0
        
        End If
        
        If Triangles(I).edge2.Start.y < 0 Then
        
            Triangles(I).edge2.Start.y = 0
        
        End If
        
        If Triangles(I).edge2.Finish.y < 0 Then
        
            Triangles(I).edge2.Finish.y = 0
        
        End If
        
        If Triangles(I).edge3.Start.y < 0 Then
        
            Triangles(I).edge3.Start.y = 0
        
        End If
        
        If Triangles(I).edge3.Finish.y < 0 Then
        
            Triangles(I).edge3.Finish.y = 0
        
        End If
        
    Next

End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   ArcCos()                               |
' |  /   \  --------                               |
' | |\_.  | ArcCos of Theta                        |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Theta for ArcCos
'   `---'
'           Returns:
'           1.<< ArcCos of Theta
'
Function ArcCos(Number As Double) As Double

    '*------------------------------------------*
    '*              ArcCos of Theta             *
    '*------------------------------------------*
    
    Select Case Number
    
        Case 1
        
            ArcCos = 0
            
        Case -1
        
            ArcCos = Pi
            
        Case Else
        
            If Abs(Number) < 1 Then _
                ArcCos = Atn(-Number / Sqr(-Number ^ 2 + 1)) + 2 * Atn(1)
    
    End Select
    
End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   GetMAXy()                              |
' |  /   \  ---------                              |
' | |\_.  | Get maximum y value for a vector       |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Index of Triangles array
'   `---'
'           Returns:
'           1.<< Maximum y value for vector
'
Function GetMAXy(ByVal I As Long) As Double

    '*------------------------------------------*
    '*      Determine maximum y for vector      *
    '*------------------------------------------*
    
    Dim TempMAXy As Double
    
    '**** Find max y value ****
      
    If Triangles(I).edge1.Start.y >= Triangles(I).edge2.Start.y Then
      
        TempMAXy = Triangles(I).edge1.Start.y
      
    Else
      
        TempMAXy = Triangles(I).edge2.Start.y
    
    End If
      
    If Triangles(I).edge3.Start.y >= TempMAXy Then
      
        TempMAXy = Triangles(I).edge3.Start.y
      
    End If
        
    GetMAXy = TempMAXy
          
End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   Button functions                       |
' |  /   \  ----------------                       |
' | |\_.  | Various functions for UI buttons       |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Private Sub BUTTON_dMinus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        d_Scale = d_Scale - CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
    
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
   
End Sub

Private Sub BUTTON_dPlus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        d_Scale = d_Scale + CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
     
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
    
End Sub


Private Sub BUTTON_LightOriginxMinus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        L.x = L.x - CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
    
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
          
End Sub

Private Sub BUTTON_LightOriginxPlus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        L.x = L.x + CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
       
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub

Private Sub BUTTON_LightOriginyMinus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        L.y = L.y - CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
     
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
        
End Sub

Private Sub BUTTON_LightOriginyPlus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        L.y = L.y + CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
       
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
    
End Sub

Private Sub BUTTON_LightOriginzMinus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        L.z = L.z - CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
     
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
      
End Sub

Private Sub BUTTON_LightOriginzPlus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        L.z = L.z + CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
       
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
    
End Sub

Private Sub BUTTON_LookatxMinus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        Lookat_x = Lookat_x - CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
     
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub

Private Sub BUTTON_LookatxPlus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        Lookat_x = Lookat_x + CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
      
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
      
End Sub

Private Sub BUTTON_LookatyMinus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        Lookat_y = Lookat_y - CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
        
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
    
End Sub

Private Sub BUTTON_LookatyPlus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
    
        Lookat_y = Lookat_y + CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
    
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
        
End Sub

Private Sub BUTTON_LookatzMinus_Click()
 
    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
       
        Lookat_z = Lookat_z - CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
     
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub

Private Sub BUTTON_LookatzPlus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
        
        Lookat_z = Lookat_z + CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
    
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub

Private Function IsValidForm() As Boolean
 
    '**** Check input values ****
    
    Dim Validated As Boolean
    Validated = True
    
    If IsNumeric(TEXTBOX_SnowRed.Text) Then
    
        If CInt(TEXTBOX_SnowRed.Text) < 0 Or CInt(TEXTBOX_SnowRed.Text) > 255 Then
        
            Validated = False
            
        End If
        
    Else
    
        Validated = False
    
    End If
        
    If IsNumeric(TEXTBOX_SnowGreen.Text) Then
    
        If CInt(TEXTBOX_SnowGreen.Text) < 0 Or CInt(TEXTBOX_SnowGreen.Text) > 255 Then
        
            Validated = False
           
        End If
           
    Else
    
        Validated = False
    
    End If
    
    If IsNumeric(TEXTBOX_SnowBlue.Text) Then
    
        If CInt(TEXTBOX_SnowBlue.Text) < 0 Or CInt(TEXTBOX_SnowBlue.Text) > 255 Then
        
            Validated = False
           
        End If
           
    Else
    
        Validated = False
    
    End If

    If IsNumeric(TEXTBOX_VegitationRed.Text) Then
    
        If CInt(TEXTBOX_VegitationRed.Text) < 0 Or CInt(TEXTBOX_VegitationRed.Text) > 255 Then
               
            Validated = False
            
        End If
     
    Else
    
        Validated = False
    
    End If
        
    If IsNumeric(TEXTBOX_VegitationGreen.Text) Then
    
        If CInt(TEXTBOX_VegitationGreen.Text) < 0 Or CInt(TEXTBOX_VegitationGreen.Text) > 255 Then
               
            Validated = False
            
        End If
     
    Else
    
        Validated = False
    
    End If
    
    If IsNumeric(TEXTBOX_VegitationBlue.Text) Then
    
        If CInt(TEXTBOX_VegitationBlue.Text) < 0 Or CInt(TEXTBOX_VegitationBlue.Text) > 255 Then
               
            Validated = False
            
        End If
     
    Else
    
        Validated = False
    
    End If
    
    If IsNumeric(TEXTBOX_WaterRed.Text) Then
    
        If CInt(TEXTBOX_WaterRed.Text) < 0 Or CInt(TEXTBOX_WaterRed.Text) > 255 Then
               
            Validated = False
            
        End If
     
    Else
    
        Validated = False
    
    End If
        
    If IsNumeric(TEXTBOX_WaterGreen.Text) Then
    
        If CInt(TEXTBOX_WaterGreen.Text) < 0 Or CInt(TEXTBOX_WaterGreen.Text) > 255 Then
               
            Validated = False
            
        End If
     
    Else
    
        Validated = False
    
    End If
    
    If IsNumeric(TEXTBOX_WaterBlue.Text) Then
    
        If CInt(TEXTBOX_WaterBlue.Text) < 0 Or CInt(TEXTBOX_WaterBlue.Text) > 255 Then
             
            Validated = False
            
        End If
     
    Else
    
        Validated = False
    
    End If

    If IsNumeric(TEXTBOX_RocksRed.Text) Then
    
        If CInt(TEXTBOX_RocksRed.Text) < 0 Or CInt(TEXTBOX_RocksRed.Text) > 255 Then
               
            Validated = False
            
        End If
     
    Else
    
        Validated = False
    
    End If
        
    If IsNumeric(TEXTBOX_RocksGreen.Text) Then
    
        If CInt(TEXTBOX_RocksGreen.Text) < 0 Or CInt(TEXTBOX_RocksGreen.Text) > 255 Then
               
            Validated = False
            
        End If
     
    Else
    
        Validated = False
    
    End If
    
    If IsNumeric(TEXTBOX_RocksBlue.Text) Then
    
        If CInt(TEXTBOX_RocksBlue.Text) < 0 Or CInt(TEXTBOX_RocksBlue.Text) > 255 Then
               
            Validated = False
            
        End If
     
    Else
    
        Validated = False
    
    End If
 
    If IsNumeric(TEXTBOX_kadr.Text) Then
    
        If CInt(TEXTBOX_kadr.Text) < 0 Or CInt(TEXTBOX_kadr.Text) > 2 Then
               
            Validated = False
            
        End If
     
    Else
    
        Validated = False
    
    End If
    
    If IsNumeric(TEXTBOX_DetailLevel.Text) Then
    
        If CInt(TEXTBOX_DetailLevel.Text) < 2 Or CInt(TEXTBOX_DetailLevel.Text) > 200 Then
               
            Validated = False
            
        End If
     
    Else
    
        Validated = False
    
    End If
    
    IsValidForm = Validated
    
End Function

Private Sub BUTTON_NewLandscape_Click()
  
    If IsValidForm() Then
        
        '**** Colors ****
    
        Snow_Red = CInt(TEXTBOX_SnowRed.Text)
        Snow_Green = CInt(TEXTBOX_SnowGreen.Text)
        Snow_Blue = CInt(TEXTBOX_SnowBlue.Text)
    
        Vegitation_Red = CInt(TEXTBOX_VegitationRed.Text)
        Vegitation_Green = CInt(TEXTBOX_VegitationGreen.Text)
        Vegitation_Blue = CInt(TEXTBOX_VegitationBlue.Text)
    
        Water_Red = CInt(TEXTBOX_WaterRed.Text)
        Water_Green = CInt(TEXTBOX_WaterGreen.Text)
        Water_Blue = CInt(TEXTBOX_WaterBlue.Text)
        
        Rocks_Red = CInt(TEXTBOX_RocksRed.Text)
        Rocks_Green = CInt(TEXTBOX_RocksGreen.Text)
        Rocks_Blue = CInt(TEXTBOX_RocksBlue.Text)
        
        '**** Check water ****
        
        If CHECKBOX_Water.value = Checked Then
        
            ShowWater = True
            
        Else
        
            ShowWater = False
        
        End If
       
        '**** Grab other values ****
       
        kadr = CDbl(TEXTBOX_kadr.Text)
        DetailLevel = CInt(TEXTBOX_DetailLevel.Text)
    
        '**** Regenerate ****
    
        Call NewLandscape
            
    Else
    
        MsgBox ("Values must be numeric, colors from 0-255, kadr from 0 to 2, and detail level >= 2")
        
    End If

End Sub

Private Sub BUTTON_Save_Click()
  
    '*------------------------------------------*
    '*               Save Image                 *
    '*------------------------------------------*
 
    '**** create and set cancelled bool ****
 
    Dim cancelled As Boolean
 
    cancelled = True
 
    '**** Trap File Error ****
    
    On Error GoTo Error_Handler
    
    '**** Dialog box options ****
    
    DIALOG_Save.DefaultExt = "bmp"
    DIALOG_Save.Filter = "Bitmap files|*.bmp"
    DIALOG_Save.FilterIndex = 1
    DIALOG_Save.Flags = cdlOHideReadOnly Or cdlOFNPathMustExist Or _
        cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
    DIALOG_Save.DialogTitle = "Select the image file"
    DIALOG_Save.CancelError = True
    
    '**** Show save dialog box ****
    
    DIALOG_Save.ShowSave
    
    '**** Box was not cancelled ****
    
    cancelled = False
    
    '**** Save file ****
    
    SavePicture PICTUREBOX_Main.Image, DIALOG_Save.FileName
 
    Exit Sub
    
Error_Handler:
    
    '**** show error message box if not cancelled ****
    
    If Not cancelled Then
    
        Dim result As VbMsgBoxResult
        result = MsgBox("Invalid File Operation", , "File Error")
        
    End If
    
End Sub

Private Sub BUTTON_ToggleAxis_Click()

    ShowAxis = Not ShowAxis
    
    Call Render

End Sub

Private Sub BUTTON_ToggleWireframe_Click()

    Wireframe = Not Wireframe
    
    Call Render

End Sub

Private Sub BUTTON_TwistAngleMinus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
       
        TwistAngle = TwistAngle - CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
     
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub

Private Sub BUTTON_TwistAnglePlus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
       
        TwistAngle = TwistAngle + CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
     
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub

Private Sub BUTTON_xMinus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
       
        ViewCoords_x = ViewCoords_x - CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
     
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub

Private Sub BUTTON_xPlus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
       
        ViewCoords_x = ViewCoords_x + CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
     
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub

Private Sub BUTTON_yMinus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
       
        ViewCoords_y = ViewCoords_y - CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
     
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub

Private Sub BUTTON_yPlus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
       
        ViewCoords_y = ViewCoords_y + CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
   
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub

Private Sub BUTTON_zMinus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
       
        ViewCoords_z = ViewCoords_z - CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
   
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub

Private Sub BUTTON_zPlus_Click()

    If IsNumeric(TEXTBOX_ChangeByAmount.Text) Then
       
        ViewCoords_z = ViewCoords_z + CDbl(TEXTBOX_ChangeByAmount.Text)
    
        Call Render
   
    Else
    
        MsgBox ("+ or - amount / Value must be numeric")
        
    End If
       
End Sub
