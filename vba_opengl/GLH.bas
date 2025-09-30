Attribute VB_Name = "GLh"
Option Explicit
Public Enum Glenum
    GL_AMBIENT_AND_DIFFUSE = &H1602
    GL_COLOR_BUFFER_BIT = &H4000
    GL_DEPTH_BUFFER_BIT = &H100
    GL_DEPTH_TEST = &HB71
    GL_LIGHT0 = &H4000
    GL_LIGHTING = &HB50
    GL_MODELVIEW = &H1700
    GL_PROJECTION = &H1701
    GL_TRIANGLES = &H4
    GL_FRONT = 1028
    GL_VERTEX_ARRAY = 32884
    GL_NORMAL_ARRAY = 32885
    GL_DOUBLE = &H140A
    PFD_DOUBLEBUFFER = 1
    PFD_DRAW_TO_WINDOW = 4
    PFD_SUPPORT_OPENGL = 32
End Enum
Public Type Vector2
    X As Single
    Y As Single
End Type
Public Type Vector2d
    X As Double
    Y As Double
End Type
Public Type Vector3d
    X As Double
    Y As Double
    z As Double
End Type
Public Type Color4
    R As Single
    G As Single
    b As Single
    a As Single
End Type
Public Type PIXELFORMATDESCRIPTOR
    nSize As Long
    nVersion As Long
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type
Public Type GLYPHMETRICSFLOAT
    gmfBlackBoxX As Single
    gmfBlackBoxY As Single
    gmfptGlyphOrigin As Vector2
    gmfCellIncX As Single
    gmfCellIncY As Single
End Type
Public Type B4
    b(3) As Byte
End Type
Public Type S1
    s As Single
End Type
Public Type L1
    l As Long
End Type
Public Type STLFile
    Vertex() As Vector3d
    Normal() As Vector3d
End Type
Public Function B2Single(ByRef B1, ByRef B2, ByRef B3, ByRef B4) As Single
    Dim X As B4, Y As S1
    With X: .b(0) = B1: .b(1) = B2: .b(2) = B3: .b(3) = B4: End With
    LSet Y = X
    B2Single = Y.s
End Function
Public Function B2Long(ByRef B1, ByRef B2, ByRef B3, ByRef B4) As Long
    Dim X As B4, Y As L1
    With X: .b(0) = B1: .b(1) = B2: .b(2) = B3: .b(3) = B4: End With
    LSet Y = X
    B2Long = Y.l
End Function
Public Function Vector3d(ByVal X As Double, ByVal Y As Double, ByVal z As Double) As Vector3d
    With Vector3d: .X = X: .Y = Y: .z = z: End With
End Function
Public Function Vector2d(ByVal X As Double, ByVal Y As Double) As Vector2d
    With Vector2d: .X = X: .Y = Y: End With
End Function
Public Function Color4(ByVal R As Single, ByVal G As Single, ByVal b As Single, ByVal a As Single) As Color4
    With Color4: .a = a: .b = b: .G = G: .R = R: End With
End Function
Public Function GetIdentityMatrix() As Double()
    Dim tmparr(3, 3) As Double
    tmparr(0, 0) = 1: tmparr(1, 1) = 1: tmparr(2, 2) = 1: tmparr(3, 3) = 1
    GetIdentityMatrix = tmparr
End Function
Public Function LoadSTL(FilePath) As STLFile
    Dim Vertex() As Vector3d, Normal() As Vector3d
    Dim inputFn As Long, Fcnt As Long, i As Long, k As Long, b() As Byte, l As Long, m As Long, n As Long
    inputFn = FreeFile
    Open FilePath For Binary As #inputFn
        ReDim b(LOF(inputFn))
        Get #inputFn, , b
    Close #inputFn
    Fcnt = -1 + 3 * B2Long(b(80), b(81), b(82), b(83))
    ReDim Vertex(Fcnt) As Vector3d
    ReDim Normal(Fcnt) As Vector3d
    i = 84
    Do
        l = k + 0
        m = k + 1
        n = k + 2
        Normal(l) = Vector3d(B2Single(b(i + 0), b(i + 1), b(i + 2), b(i + 3)), B2Single(b(i + 4), b(i + 5), b(i + 6), b(i + 7)), B2Single(b(i + 8), b(i + 9), b(i + 10), b(i + 11)))
        Normal(m) = Normal(l)
        Normal(n) = Normal(l)
        Vertex(l) = Vector3d(B2Single(b(i + 12), b(i + 13), b(i + 14), b(i + 15)), B2Single(b(i + 16), b(i + 17), b(i + 18), b(i + 19)), B2Single(b(i + 20), b(i + 21), b(i + 22), b(i + 23)))
        Vertex(m) = Vector3d(B2Single(b(i + 24), b(i + 25), b(i + 26), b(i + 27)), B2Single(b(i + 28), b(i + 29), b(i + 30), b(i + 31)), B2Single(b(i + 32), b(i + 33), b(i + 34), b(i + 35)))
        Vertex(n) = Vector3d(B2Single(b(i + 36), b(i + 37), b(i + 38), b(i + 39)), B2Single(b(i + 40), b(i + 41), b(i + 42), b(i + 43)), B2Single(b(i + 44), b(i + 45), b(i + 46), b(i + 47)))
        k = k + 3
        i = i + 50
    Loop Until k >= Fcnt
    LoadSTL.Vertex = Vertex
    LoadSTL.Normal = Normal
End Function
