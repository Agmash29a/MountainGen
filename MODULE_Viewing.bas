Attribute VB_Name = "MODULE_Viewing"
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
'`'--.___  _\  /    | ViewTransform (Module)  ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 22/5/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * Viewing and Projection Transformations.                *
' *--------------------------------------------------------*
'
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   ViewTransform()                        |
' |  /   \  ---------------                        |
' | |\_.  | 8 paramater viewing transformation     |
' |\|  | /| by Kevin Suffern                      /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Vector3D
'   `---'   2. - 4.>> a,b, and c - origin of viewing coordinates
'           5. - 7.>> e,f, and g - Lookat point
'           8. >> Twist angle of viewing coordinates
'
Public Sub ViewTransform(ByRef v1 As Vector3D, ByVal a As Double, ByVal b As Double, _
                         ByVal c As Double, ByVal e As Double, ByVal f As Double, ByVal g As Double, ByVal Twist As Double)

    '*------------------------------------------*
    '*          Declare and initialise          *
    '*------------------------------------------*
    
    Dim r As Double
    
    Dim sin_theta As Double
    Dim cos_theta As Double
    Dim sin_phi As Double
    Dim cos_phi As Double
    
    Dim sin_twist As Double
    Dim cos_twist As Double
    
    '*------------------------------------------*
    '*           PreCalculate values            *
    '*------------------------------------------*
     
    '**** Calculate r, sins, and cos' ****
    
    r = Sqr((a - e) ^ 2 + (b - f) ^ 2 + (c - g) ^ 2)
    
    sin_theta = (b - f) / Sqr((a - e) ^ 2 + (b - f) ^ 2)
    cos_theta = (a - e) / Sqr((a - e) ^ 2 + (b - f) ^ 2)
    sin_phi = Sqr((a - e) ^ 2 + (b - f) ^ 2) / r
    cos_phi = (c - g) / r
    
    sin_twist = Sin(Twist)
    cos_twist = Cos(Twist)
    
    '*------------------------------------------*
    '*         Apply vector to matrix           *
    '*------------------------------------------*
    
    v1.xv = v1.x * ((-cos_twist * sin_theta) - (sin_twist * cos_theta * cos_phi)) + _
            v1.y * ((cos_twist * cos_theta) - (sin_twist * sin_theta * cos_phi)) + _
            v1.z * (sin_twist * sin_phi) + _
            ((cos_twist * (a * sin_theta - b * cos_theta)) + (sin_twist * (a * cos_theta + b * sin_theta) * cos_phi) - (c * sin_twist * sin_phi))
            
    v1.yv = v1.x * ((sin_twist * sin_theta) - (cos_twist * cos_theta * cos_phi)) + _
            v1.y * ((-sin_twist * cos_theta) - (cos_twist * sin_theta * cos_phi)) + _
            v1.z * (cos_twist * sin_phi) + _
            ((-sin_twist * (a * sin_theta - b * cos_theta)) + (cos_twist * (a * cos_theta + b * sin_theta) * cos_phi) - (c * cos_twist * sin_phi))
            
    v1.zv = v1.x * (-cos_theta * sin_phi) + _
            v1.y * (-sin_theta * sin_phi) + _
            v1.z * (-cos_phi) + _
            (((a * cos_theta + b * sin_theta) * sin_phi) + (c * cos_phi))
            
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   PerspectiveProject()                   |
' |  /   \  -------------------                    |
' | |\_.  | Project 3d coords onto 2d plane        |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> d - Scale factor
'   `---'   2.>> v1 - Vector to transform
'
Public Sub PerspectiveProject(ByVal d As Double, ByRef v1 As Vector3D)

    '*------------------------------------------*
    '*       Divide x and y elements by z       *
    '*------------------------------------------*
     
    v1.xp = (v1.xv / v1.zv) * d
    v1.yp = (v1.yv / v1.zv) * d

End Sub


