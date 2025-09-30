VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm2"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11310
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents glControl As GLFrame
Private GL As GL
Private stl As STLFile
Private Roll As Double, Pitch As Double, Mat() As Double
Private Material As Color4
Private Sub CommandButton1_Click()
    stl = LoadSTL(Application.GetOpenFilename)
    glControl.Refresh
End Sub
Private Sub UserForm_Initialize()
    Set glControl = New GLFrame
    Set GL = New GL
    Mat = GetIdentityMatrix
End Sub
Private Sub UserForm_Activate()
    Material = Color4(0.3, 0.7, 0.3, 1)
    glControl.Init Me.Frame1, GL
End Sub
Private Sub GLControl_Load()
    With GL
        .Enable GL_LIGHT0
        .Enable GL_LIGHTING
        .Viewport 0, 0, glControl.Width, glControl.Height
        .Materialfv GL_FRONT, GL_AMBIENT_AND_DIFFUSE, VarPtr(Material)
    End With
    glControl.Refresh
End Sub
Private Sub ScrollBar1_Change()
    glControl.Refresh
End Sub
Private Sub GLControl_DragDelta(ByVal DeltaX As Double, DeltaY As Double, Button As Integer)
    If Button <> 1 Then Exit Sub
    Roll = Roll + DeltaX * 0.005
    Pitch = Pitch + DeltaY * 0.005
    Mat(0, 0) = Cos(Roll): Mat(1, 0) = -Sin(Roll): Mat(0, 1) = Sin(Roll): Mat(1, 1) = Cos(Roll)
    glControl.Refresh
End Sub
Private Sub GLControl_Paint()
    With GL
        .Clear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        
        .MatrixMode GL_PROJECTION
        .LoadIdentity
        .Perspective Me.ScrollBar1.value, glControl.Width / glControl.Height, 10, 3000
        .LookAt 0, -1000, 1000, 0, 0, 50, 0, 0, 1

        .MatrixMode GL_MODELVIEW
        .LoadIdentity
        
        .MultMatrixd (VarPtr(Mat(0, 0)))
        If Not Not stl.Vertex Then
            .PushMatrix
                .EnableClientState GL_VERTEX_ARRAY
                .EnableClientState GL_NORMAL_ARRAY
                .NormalPointer GL_DOUBLE, 0, VarPtr(stl.Normal(0))
                .VertexPointer 3, GL_DOUBLE, 0, VarPtr(stl.Vertex(0))
                .DrawArrays GL_TRIANGLES, 0, UBound(stl.Vertex) + 1
                .DisableClientState GL_NORMAL_ARRAY
                .DisableClientState GL_VERTEX_ARRAY
            .PopMatrix
        End If
        .SwapBuffers
    End With
End Sub
