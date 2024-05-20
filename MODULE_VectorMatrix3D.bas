Attribute VB_Name = "MODULE_VectorMatrix3D"
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
'`'--.___  _\  /    | VectorMatrix3D (Module) ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 22/5/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * Vector and Matrix Types and Operations. Based on a C++ *
' * class by Kevin Suffern                                 *
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

'**** A vector type ****

Public Type Vector3D

    x As Double
    y As Double
    z As Double
    
    xv As Double
    yv As Double
    zv As Double
    
    xp As Double
    yp As Double

End Type

'**** A matrix type ****

Public Type Matrix3D

    Elements(1 To 4, 1 To 4) As Double

End Type
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   ZeroVector()                           |
' |  /   \  ------------                           |
' | |\_.  | Generates a zero vector                |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Vector3D
'   `---'
'           Returns:
'           1.<< v1 = Vector3D with x,y and z set to 0
'
Public Sub ZeroVector(ByRef v1 As Vector3D)

    '*------------------------------------------*
    '*           Generate Zero Vector           *
    '*------------------------------------------*
    
    v1.x = 0#
    v1.y = 0#
    v1.z = 0#
 
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   EqualVector()                          |
' |  /   \  ------------                           |
' | |\_.  | Generates a vector with x=y=z=n        |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Vector3D
'   `---'   2.>> n
'
'           Returns:
'           1.<< v1 = Vector3D with x,y and z set to n
'
Public Sub EqualVector(ByRef v1 As Vector3D, ByVal N As Double)

    '*------------------------------------------*
    '*           Generate Equal Vector          *
    '*------------------------------------------*
    
    v1.x = N
    v1.y = N
    v1.z = N
   
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   AddVectors()                           |
' |  /   \  ------------                           |
' | |\_.  | Adds 2 Vectors                         |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> v1
'   `---'   2.>> v2
'
'           Returns:
'           1.<< v1 + v2
'
Public Function AddVectors(ByRef v1 As Vector3D, ByRef v2 As Vector3D) As Vector3D

    '*------------------------------------------*
    '*              Adds 2 Vectors              *
    '*------------------------------------------*
    
    Dim v3 As Vector3D
    
    v3.x = v1.x + v2.x
    v3.y = v1.y + v2.y
    v3.z = v1.z + v2.z
    
    AddVectors = v3
    
End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   SubtractVectors()                      |
' |  /   \  -----------------                      |
' | |\_.  | Subtraction on 2 Vectors               |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> v1
'   `---'   2.>> v2
'
'           Returns:
'           1.<< v1 - v2
'
Public Function SubtractVectors(ByRef v1 As Vector3D, ByRef v2 As Vector3D) As Vector3D

    '*------------------------------------------*
    '*            Subtract 2 Vectors            *
    '*------------------------------------------*
    
    Dim v3 As Vector3D
    
    v3.x = v1.x - v2.x
    v3.y = v1.y - v2.y
    v3.z = v1.z - v2.z
    
    SubtractVectors = v3
    
End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   NegateVector()                         |
' |  /   \  --------------                         |
' | |\_.  | Negate a Vector                        |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> v1
'   `---'
'           Returns:
'           1.<< v1 = -v1
'
Public Sub NegateVector(ByRef v1 As Vector3D)

    '*------------------------------------------*
    '*              Negate a Vector             *
    '*------------------------------------------*
    
    v1.x = -v1.x
    v1.y = -v1.y
    v1.z = -v1.z
    
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   AddScalarToVector()                    |
' |  /   \  -------------------                    |
' | |\_.  | Add Scalar to a Vector                 |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> v1
'   `---'   2.>> n
'
'           Returns:
'           1.<< v1 = n + (v1)
'
Public Sub AddScalarToVector(ByRef v1 As Vector3D, ByVal N As Double)

    '*------------------------------------------*
    '*            Add Scalar To Vector          *
    '*------------------------------------------*
    
    v1.x = v1.x + N
    v1.y = v1.y + N
    v1.z = v1.z + N
    
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   MultiplyVectorByScalar()               |
' |  /   \  ------------------------               |
' | |\_.  | Multiply vector by a scalar            |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> v1
'   `---'   2.>> n
'
'           Returns:
'           1.<< v1 = n * (v1)
'
Public Sub MultiplyVectorByScalar(ByRef v1 As Vector3D, ByVal N As Double)

    '*------------------------------------------*
    '*       Multiply vector by a scalar        *
    '*------------------------------------------*
    
    v1.x = v1.x * N
    v1.y = v1.y * N
    v1.z = v1.z * N
    
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   DivideVectorByScalar()                 |
' |  /   \  ----------------------                 |
' | |\_.  | Divide vector by a scalar              |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> v1
'   `---'   2.>> n
'
'           Returns:
'           1.<< v1 = (v1)/n
'
Public Sub DivideVectorByScalar(ByRef v1 As Vector3D, ByVal N As Double)

    '*------------------------------------------*
    '*        Divide vector by a scalar         *
    '*------------------------------------------*
    
    v1.x = v1.x / N
    v1.y = v1.y / N
    v1.z = v1.z / N
    
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   CrossProduct()                         |
' |  /   \  --------------                         |
' | |\_.  | Cross product of 2 Vectors             |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> v1
'   `---'   2.>> v2
'
'           Returns:
'           1.<< v1 x v2
'
Public Function CrossProduct(ByRef v1 As Vector3D, ByRef v2 As Vector3D) As Vector3D

    '*------------------------------------------*
    '*        Cross product of 2 vectors        *
    '*------------------------------------------*
    
    Dim v3 As Vector3D
    
    v3.x = v1.y * v2.z - v1.z * v2.y
    v3.y = v1.z * v2.x - v1.x * v2.z
    v3.z = v1.x * v2.y - v1.y * v2.x
    
    CrossProduct = v3
    
End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   DistanceBetweenTwoVectors()            |
' |  /   \  ---------------------------            |
' | |\_.  | Distance between 2 vectors             |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> v1
'   `---'   2.>> v2
'
'           Returns:
'           1.<< Distance between v1 and v2
'
Public Function DistanceBetweenTwoVectors(ByRef v1 As Vector3D, ByRef v2 As Vector3D) As Double

    '*------------------------------------------*
    '*       Distance between v1 and v2         *
    '*------------------------------------------*
    
    DistanceBetweenTwoVectors = Sqr((v1.x - v2.x) * (v1.x - v2.x) + _
                                (v1.y - v2.y) * (v1.y - v2.y) + _
                                (v1.z - v2.z) * (v1.z - v2.z))
    
End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   NormaliseVector()                      |
' |  /   \  -----------------                      |
' | |\_.  | Convert to a unit vector               |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> v1
'   `---'
'           Returns:
'           1.<< v1 = Unit vector of v1
'
Public Sub NormaliseVector(ByRef v1 As Vector3D)

    '*------------------------------------------*
    '*               Normalise v1               *
    '*------------------------------------------*
    
    Dim Length As Double
    
    Length = Sqr((v1.x * v1.x) + (v1.y * v1.y) + (v1.z * v1.z))
   
    v1.x = v1.x / Length
    v1.y = v1.y / Length
    v1.z = v1.z / Length
    
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   VectorLength()                         |
' |  /   \  --------------                         |
' | |\_.  | Get the length of a vector             |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> v1
'   `---'
'           Returns:
'           1.<< Length of v1
'
Public Function VectorLength(ByRef v1 As Vector3D) As Double

    '*------------------------------------------*
    '*               Normalise v1               *
    '*------------------------------------------*
    
    VectorLength = Sqr(v1.x * v1.x + v1.y * v1.y + v1.z * v1.z)
    
End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   DotProduct()                           |
' |  /   \  ------------                           |
' | |\_.  | Dot product of 2 Vectors               |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> v1
'   `---'   2.>> v2
'
'           Returns:
'           1.<< v1 . v2
'
Public Function DotProduct(ByRef v1 As Vector3D, ByRef v2 As Vector3D) As Double

    '*------------------------------------------*
    '*         Dot product of 2 vectors         *
    '*------------------------------------------*

    DotProduct = v1.x * v2.x + v1.y * v2.y + v1.z * v2.z
    
End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   MakeIdentityMatrix()                   |
' |  /   \  -------------------                    |
' | |\_.  | Make an identity matrix                |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Matrix3D
'   `---'
'           Returns:
'           1.<< m1 = m1 as identity matrix
'
Public Sub MakeIdentityMatrix(ByRef m1 As Matrix3D)

    '*------------------------------------------*
    '*          Generate Identity Matrix        *
    '*------------------------------------------*
    
    Dim I As Integer
    Dim j As Integer
    
    '**** Iterate through all matrix elements ****

    For I = 1 To 4
        
        For j = 1 To 4
                
            '**** check for an identity element ****
            
            If I = j Then
                
                m1.Elements(I, j) = 1
                    
            Else
                
                m1.Elements(I, j) = 0
                    
            End If
                
        Next j
            
    Next I
 
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   MultiplyMatrixByMatrix()               |
' |  /   \  ------------------------               |
' | |\_.  | Make an identity matrix                |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> m1 (Matrix3D)
'   `---'   2.>> m2 (Matrix3D)
'
'           Returns:
'           1.<< m1 * m2
'
Public Function MultiplyMatrixByMatrix(ByRef m1 As Matrix3D, ByRef m2 As Matrix3D) As Matrix3D

    '*------------------------------------------*
    '*      Multiply a matrix by a matrix       *
    '*------------------------------------------*
    
    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim value As Double

    Dim m3 As Matrix3D
    
    '**** Iterate through all matrix elements ****

    For I = 1 To 4
    
        For j = 1 To 4
        
            value = 0#
            
            For k = 1 To 4
            
                '**** multiply individual elements ****
            
                value = value + m1.Elements(I, k) * m2.Elements(k, j)
                
            Next k
            
            m3.Elements(I, j) = value
            
        Next j
        
    Next I
    
    MultiplyMatrixByMatrix = m3
    
End Function


