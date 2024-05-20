Attribute VB_Name = "MODULE_Terrain"
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
'`'--.___  _\  /    | Terrain (Module)    ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 22/5/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * Methods for generating random terrain.                 *
' *--------------------------------------------------------*
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

'**** Powers of 16 ****

Public Powers2(16) As Long

'**** Detail level ****

Public DetailLevel As Integer

'**** Point Type ****

Public Type PointType

        x As Integer
        y As Integer
        z As Integer
        
End Type

'**** TriangleCounter *****

Public TriangleCounter As Long

'**** Temporary Points ****

Public pp3 As PointType
Public pp2 As PointType
Public pp1 As PointType
'
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   DivideLine()                           |
' |  /   \  ------------                           |
' | |\_.  | Divide a line at its midpoints         |
' |\|  | /| recursively                           /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Current detail level
'   `---'   2.>> Array of points
'           3.>> Index to start from
'           4.>> Index to finish on
'
Public Sub DivideLine(ByVal Level As Integer, ByRef PointTypes() As PointType, ByVal Start As Integer, ByVal Finish As Integer)

    '*------------------------------------------*
    '*          Declare and initialise          *
    '*------------------------------------------*
    
    Dim pos As Integer
    
    Dim dx As Integer   '-> Change in x
    Dim dy As Integer   '-> Change in y
    Dim dz As Integer   '-> Change in z
    
    Dim DisplacementAmount As Integer
    Dim UpOrDown As Integer
    
    Dim point_ As PointType
    
    Dim p1 As PointType
    Dim p2 As PointType

    '**** determine point range ****
        
    p1 = PointTypes(Start)
    p2 = PointTypes(Finish)
    
    '**** work out change values for x,y, and z ****
    
    dx = p2.x - p1.x
    dy = p2.y - p1.y
    dz = p2.z - p1.z

    '*------------------------------------------*
    '*      Determine random displacement       *
    '*------------------------------------------*
    
    Randomize
    
    '**** calculate displacement value depending on level ****
    
    If Level = DetailLevel - 1 Then
     
        '**** Large displacement for initial level ****
     
        DisplacementAmount = (Int(Rnd * 250) + 1)
    
    Else
    
        '**** Smaller displacements for subsequent levels ****
        
        DisplacementAmount = (Int(Rnd * (Level ^ 2)) + 1)
    
    End If

    '**** Determine whether to displace up or down ****
   
    UpOrDown = (Int(Rnd * 3) + 1)
    
    If UpOrDown = 3 Then
    
        '**** Displace down (up otherwise) ****
    
        DisplacementAmount = -DisplacementAmount
    
    End If

    '*------------------------------------------*
    '*      Displace y and find new points      *
    '*------------------------------------------*
     
    point_.x = p1.x + (dx + 1) / 2
    point_.y = p1.y + (dy + 1) / 2 + CInt(DisplacementAmount)
    point_.z = p1.z + (dz + 1) / 2
    
    pos = (Start + Finish) / 2
    
    PointTypes(pos) = point_
    
    '*------------------------------------------*
    '* Call self again if need more subdivision *
    '*------------------------------------------*
    
    If Level > 1 Then
    
        Call DivideLine(Level - 1, PointTypes, Start, pos)
        Call DivideLine(Level - 1, PointTypes, pos, Finish)
    
    End If

End Sub
'
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   MakeTriangle()                         |
' |  /   \  --------------                         |
' | |\_.  | Make individual triangles              |
' |\|  | /| recursively                           /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Current detail level
'   `---'   2. - 4.>> Sides as arrays of points
'
Public Sub MakeTriangle(ByVal Level As Integer, ByRef Side1() As PointType, ByRef Side2() As PointType, ByRef Side3() As PointType)

    '*------------------------------------------*
    '*          Declare and initialise          *
    '*------------------------------------------*
    
    Dim I As Integer
    
    Dim S1() As PointType
    Dim S2() As PointType
    Dim S3() As PointType
    
    Dim Side1Plus()  As PointType
    Dim Side2Plus()  As PointType
    Dim Side3Plus()  As PointType
    
    Dim SideLength As Integer
    
    SideLength = Powers2(Level - 1)
    
    '**** determine new points of triangle ****
    
    pp1 = Side1(SideLength)
    pp2 = Side2(SideLength)
    pp3 = Side3(SideLength)

    '*------------------------------------------*
    '*       Divide lines for all sides         *
    '*------------------------------------------*
    
    Call Make_Sides(Level, SideLength, S1, S2, S3)

    '*------------------------------------------*
    '*             Reassign sides               *
    '*------------------------------------------*
        
    '**** Side1 + SideLength ****
    
    ReDim Side1Plus(UBound(Side1) - slen)
    
    For I = SideLength To UBound(Side1)
        
        Side1Plus(I - SideLength) = Side1(I)
    
    Next
    
    '**** Side2 + SideLength ****
    
    ReDim Side2Plus(UBound(Side2) - slen)
    
    For I = SideLength To UBound(Side2)
        
        Side2Plus(I - SideLength) = Side2(I)
    
    Next
    
    '**** Side3 + SideLength ****
    
    ReDim Side3Plus(UBound(Side3) - slen)
    
    For I = SideLength To UBound(Side3)
        
        Side3Plus(I - SideLength) = Side3(I)
    
    Next

    '*------------------------------------------*
    '*     Make larger trianlges or store       *
    '*------------------------------------------*
      
    If Level > 1 Then
    
        '**** Make larger triangles ****
    
        Call MakeTriangle(Level - 1, Side1Plus, Side2, S3)
        Call MakeTriangle(Level - 1, Side2Plus, Side3, S2)
        Call MakeTriangle(Level - 1, Side1, S1, Side3Plus)
        Call MakeTriangle(Level - 1, S1, S2, S3)
       
    Else
       
        '**** Lowest level reached, store as world coordinates ****
       
        ReDim Preserve Triangles(TriangleCounter)
    
        Triangles(TriangleCounter).edge1.Start.x = Side1(0).x
        Triangles(TriangleCounter).edge1.Start.y = Side1(0).y
        Triangles(TriangleCounter).edge1.Start.z = Side1(0).z
        
        Triangles(TriangleCounter).edge1.Finish.x = Side1(2).x
        Triangles(TriangleCounter).edge1.Finish.y = Side1(2).y
        Triangles(TriangleCounter).edge1.Finish.z = Side1(2).z
         
        Triangles(TriangleCounter).edge2.Start.x = Side2(0).x
        Triangles(TriangleCounter).edge2.Start.y = Side2(0).y
        Triangles(TriangleCounter).edge2.Start.z = Side2(0).z
        
        Triangles(TriangleCounter).edge2.Finish.x = Side2(2).x
        Triangles(TriangleCounter).edge2.Finish.y = Side2(2).y
        Triangles(TriangleCounter).edge2.Finish.z = Side2(2).z
         
        Triangles(TriangleCounter).edge3.Start.x = Side3(0).x
        Triangles(TriangleCounter).edge3.Start.y = Side3(0).y
        Triangles(TriangleCounter).edge3.Start.z = Side3(0).z
            
        Triangles(TriangleCounter).edge3.Finish.x = Side3(2).x
        Triangles(TriangleCounter).edge3.Finish.y = Side3(2).y
        Triangles(TriangleCounter).edge3.Finish.z = Side3(2).z
         
        TriangleCounter = TriangleCounter + 1
        
    End If

End Sub
'
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   MakeSides()                            |
' |  /   \  -----------                            |
' | |\_.  | Divide lines for each side of          |
' |\|  | /| triangle                              /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Current detail level
'   `---'   2.>> Length of sides
'           3. - 5.>> Sides as arrays of points
'
Public Sub Make_Sides(ByVal Level As Integer, ByVal SideLength As Integer, ByRef S1() As PointType, ByRef S2() As PointType, ByRef S3() As PointType)

    '*------------------------------------------*
    '*          Declare and initialise          *
    '*------------------------------------------*
    
    ReDim S1(SideLength + 1)
    ReDim S2(SideLength + 1)
    ReDim S3(SideLength + 1)
    
    '**** assign points to boundries ****
    
    S1(0) = pp1
    S1(SideLength) = pp3
    S2(0) = pp3
    S2(SideLength) = pp2
    S3(0) = pp2
    S3(SideLength) = pp1

    '*------------------------------------------*
    '*     Divide each side and store points    *
    '*------------------------------------------*
    
    If Level > 1 Then
    
        Call DivideLine(Level - 1, S1, 0, SideLength)
        Call DivideLine(Level - 1, S2, 0, SideLength)
        Call DivideLine(Level - 1, S3, 0, SideLength)
    
    End If

End Sub
