Attribute VB_Name = "MODULE_MountainGenGlobals"
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
'`'--.___  _\  /    | Globals (Module)        ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 27/3/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * Globals.                                               *
' *--------------------------------------------------------*

'**** From edges to the world ****

Public Type edge

    Start  As Vector3D
    Finish As Vector3D

End Type

Public Type Triangle

    edge1 As edge
    edge2 As edge
    edge3 As edge
    
    Middle_z As Double
    I As Double
    
End Type

Public Triangles() As Triangle

'**** Origin of view coordinates ****

Public ViewCoords_x As Double
Public ViewCoords_y As Double
Public ViewCoords_z As Double

'**** Origin of lookat point ****

Public Lookat_x As Double
Public Lookat_y As Double
Public Lookat_z As Double

'**** Axis' ****

Public Axis_ystart As Vector3D
Public Axis_yfinish As Vector3D

Public Axis_xstart As Vector3D
Public Axis_xfinish As Vector3D

Public Axis_zstart As Vector3D
Public Axis_zfinish As Vector3D

'**** d ****

Public d_Scale As Integer

'**** Twist angle ****

Public TwistAngle As Double

'**** L ****

Public L As Vector3D

'**** kdr & kar ****

Public kadr As Double

'**** Maximum Height ****

Public MaximumHeight As Double
Public UnitsOfHeight As Double

'**** Wireframe ****

Public Wireframe As Boolean

'**** ShowAxis ****

Public ShowAxis As Boolean

'**** Show Water ****

Public ShowWater As Boolean

'**** Resolution Displacement ****

Public ResolutionDisplacement_x As Integer
Public ResolutionDisplacement_y As Integer

'**** Colors ****

Public Snow_Red As Integer
Public Snow_Green As Integer
Public Snow_Blue As Integer

Public Vegitation_Red As Integer
Public Vegitation_Green As Integer
Public Vegitation_Blue As Integer

Public Rocks_Red As Integer
Public Rocks_Green As Integer
Public Rocks_Blue As Integer

Public Water_Red As Integer
Public Water_Green As Integer
Public Water_Blue As Integer
