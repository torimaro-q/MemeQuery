VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'■関数
'User32
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
'Gdi32
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByRef pfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare PtrSafe Function SetPixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByVal format As Long, ByRef ppfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare PtrSafe Function SwapBuffersA Lib "gdi32" Alias "SwapBuffers" (ByVal hdc As LongPtr) As Long
'/Glu拡張
Private Declare PtrSafe Sub gluPerspective Lib "glu32.dll" (ByVal FovY As Double, ByVal aspect As Double, ByVal zNear As Double, ByVal zfar As Double)
Private Declare PtrSafe Sub gluLookAt Lib "glu32.dll" (ByVal eyeX As Double, ByVal eyeY As Double, ByVal eyeZ As Double, _
                                                       ByVal centerX As Double, ByVal centerY As Double, ByVal centerZ As Double, _
                                                       ByVal upx As Double, ByVal upy As Double, ByVal upz As Double)
'OpenGL
'開始/終了処理
Private Declare PtrSafe Function wglCreateContext Lib "opengl32.dll" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function wglMakeCurrent Lib "opengl32.dll" (ByVal hdc As LongPtr, ByVal hGLRC As LongPtr) As Long
Private Declare PtrSafe Function wglDeleteContext Lib "opengl32.dll" (ByVal hdc As LongPtr) As Long
'座標系
Private Declare PtrSafe Sub glMultMatrixd Lib "opengl32.dll" (ByVal Ptr As LongPtr)
Private Declare PtrSafe Sub glMatrixMode Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glPushMatrix Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glPopMatrix Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glLoadIdentity Lib "opengl32.dll" ()
Private Declare PtrSafe Sub glViewport Lib "opengl32.dll" (ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
'ON/OFF
Private Declare PtrSafe Sub glEnable Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glDisable Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glEnableClientState Lib "opengl32.dll" (ByVal Glenum As Long)
Private Declare PtrSafe Sub glDisableClientState Lib "opengl32.dll" (ByVal Glenum As Long)
'色
Private Declare PtrSafe Sub glMaterialfv Lib "opengl32.dll" (ByVal Face As Long, ByVal pname As Long, ByVal param As LongPtr)
Private Declare PtrSafe Sub glClear Lib "opengl32.dll" (ByVal Mask As Long)
Private Declare PtrSafe Sub glClearColor Lib "opengl32.dll" (ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
'ポイント描画
Private Declare PtrSafe Sub glNormalPointer Lib "opengl32.dll" (ByVal Glenum As Long, ByVal Stride As Long, ByVal Ptr As LongPtr)
Private Declare PtrSafe Sub glVertexPointer Lib "opengl32.dll" (ByVal Size As Long, ByVal Glenum As Long, ByVal Stride As Long, ByVal Ptr As LongPtr)
Private Declare PtrSafe Sub glDrawArrays Lib "opengl32.dll" (ByVal Glenum As Long, ByVal First As Long, ByVal Size As Long)

'■変数
Private pfd As PIXELFORMATDESCRIPTOR
Public hformat As Long, hGLRC As LongPtr, hdc As LongPtr
Private myhwnd As LongPtr
Public Sub EnableClientState(ByVal Gle As Glenum)
    Call glEnableClientState(Gle)
End Sub
Public Sub DisableClientState(ByVal Gle As Glenum)
    Call glDisableClientState(Gle)
End Sub
Public Sub VertexPointer(ByVal Size As Long, ByVal Gle As Glenum, ByVal Stride As Long, ByVal Ptr As LongPtr)
    Call glVertexPointer(Size, Gle, Stride, Ptr)
End Sub
Public Sub NormalPointer(ByVal Gle As Glenum, ByVal Stride As Long, ByVal Ptr As LongPtr)
    Call glNormalPointer(Gle, Stride, Ptr)
End Sub
Public Sub DrawArrays(ByVal Mode As Glenum, ByVal First As Long, ByVal Size As Long)
    Call glDrawArrays(Mode, First, Size)
End Sub
Public Sub MultMatrixd(ByVal Ptr As LongPtr)
    Call glMultMatrixd(Ptr)
End Sub
Public Sub PaintStart()
    Call wglMakeCurrent(hdc, hGLRC)
End Sub
Public Sub PaintEnd()
    Call SwapBuffersA(hdc)
    Call wglMakeCurrent(0, 0)
End Sub
Public Sub MakeCurrent(ByVal hdc As LongPtr, ByVal hGLRC As LongPtr)
    Call wglMakeCurrent(hdc, hGLRC)
End Sub
Public Sub SwapBuffers()
    Call SwapBuffersA(hdc)
End Sub
Public Sub Viewport(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
    Call glViewport(X, Y, Width, Height)
End Sub
Public Sub Enable(ByVal Gle As Glenum)
    Call glEnable(Gle)
End Sub
Public Sub Disable(ByVal Gle As Glenum)
    Call glDisable(Gle)
End Sub
Public Sub MatrixMode(ByVal Gle As Glenum)
    Call glMatrixMode(Gle)
End Sub
Public Sub Materialfv(ByVal GLe1 As Glenum, ByVal GLe2 As Glenum, ByVal param As LongPtr)
    Call glMaterialfv(GLe1, GLe2, param)
End Sub
Public Sub LoadIdentity()
    Call glLoadIdentity
End Sub
Public Sub PushMatrix()
    Call glPushMatrix
End Sub
Public Sub PopMatrix()
    Call glPopMatrix
End Sub
Public Sub Clear(Optional ByVal Mask As Long = 16384)
    Call glClear(Mask)
End Sub
Public Sub ClearColor(ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
    Call glClearColor(red, green, blue, alpha)
End Sub
Public Sub Perspective(ByVal FovY As Double, ByVal aspect As Double, ByVal zNear As Double, ByVal zfar As Double)
    Call gluPerspective(FovY, aspect, zNear, zfar)
End Sub
Public Sub LookAt(ByVal eyeX As Double, ByVal eyeY As Double, ByVal eyeZ As Double, ByVal centerX As Double, ByVal centerY As Double, ByVal centerZ As Double, ByVal upx As Double, ByVal upy As Double, ByVal upz As Double)
    Call gluLookAt(eyeX, eyeY, eyeZ, centerX, centerY, centerZ, upx, upy, upz)
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Let hWnd(newPtr As LongPtr)
    myhwnd = newPtr
    Call Inithwnd
End Property
Public Property Get hWnd() As LongPtr
    hWnd = myhwnd
End Property
Private Sub Inithwnd() '前処理
    If hdc <> 0 Then Releasehwnd
    hdc = GetDC(myhwnd)
    hformat = ChoosePixelFormat(hdc, pfd)
    Call SetPixelFormat(hdc, hformat, pfd)
    hGLRC = wglCreateContext(hdc)
End Sub
Private Sub Releasehwnd()
    Call wglMakeCurrent(0, 0)
    Call wglDeleteContext(hGLRC)
    Call ReleaseDC(myhwnd, hdc)
End Sub
Private Sub Class_Initialize()
    With pfd
        .nSize = 40 ' LenB(pfd)
        .nVersion = 1
        .dwFlags = (PFD_DOUBLEBUFFER Or PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL)
        .cColorBits = 32
        .cRedBits = 0
        .cRedShift = 0
        .cGreenBits = 0
        .cGreenShift = 0
        .cBlueBits = 0
        .cBlueShift = 0
        .cAlphaBits = 0
        .cAlphaShift = 0
        .cAccumBits = 0
        .cAccumRedBits = 0
        .cAccumGreenBits = 0
        .cAccumBlueBits = 0
        .cAccumAlphaBits = 0
        .cDepthBits = 32
        .cStencilBits = 8
        .cAuxBuffers = 0
        .iLayerType = 0 'PFD_MAIN_PLANE
        .bReserved = 0
        .dwLayerMask = 0
        .dwVisibleMask = 0
        .dwDamageMask = 0
    End With
End Sub
Private Sub Class_Terminate()
    Call Releasehwnd
End Sub


