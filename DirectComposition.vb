Imports System
Imports System.Runtime.InteropServices
Imports GlobalStructures
Imports DXGI

Namespace DirectComposition

    Public Class DirectCompositionTools
        <DllImport("Dcomp.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function DCompositionCreateDevice(dxgiDevice As IDXGIDevice, <MarshalAs(UnmanagedType.LPStruct)> iid As Guid, ByRef dcompositionDevice As IntPtr) As HRESULT
        End Function

        <DllImport("Dcomp.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function DCompositionCreateDevice2(renderingDevice As IDXGIDevice, <MarshalAs(UnmanagedType.LPStruct)> iid As Guid, ByRef dcompositionDevice As IntPtr) As HRESULT
        End Function

        <DllImport("Dcomp.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function DCompositionCreateSurfaceHandle(desiredAccess As UInteger, securityAttributes As IntPtr, ByRef surfaceHandle As IntPtr) As HRESULT
        End Function
    End Class

    <ComImport()>
    <Guid("C37EA93A-E7AA-450D-B16F-9746CB0407F3")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDCompositionDevice
        <PreserveSig()>
        Function Commit() As HRESULT
        <PreserveSig()>
        Function WaitForCommitCompletion() As HRESULT
        <PreserveSig()>
        Function GetFrameStatistics(ByRef statistics As DCOMPOSITION_FRAME_STATISTICS) As HRESULT
        <PreserveSig()>
        Function CreateTargetForHwnd(hwnd As IntPtr, topmost As Boolean, ByRef target As IDCompositionTarget) As HRESULT
        <PreserveSig()>
        Function CreateVisual(ByRef visual As IDCompositionVisual) As HRESULT
        <PreserveSig()>
        Function CreateSurface(width As UInteger, height As UInteger, pixelFormat As DXGI_FORMAT, alphaMode As DXGI_ALPHA_MODE, ByRef surface As IDCompositionSurface) As HRESULT
        <PreserveSig()>
        Function CreateVirtualSurface(initialWidth As UInteger, initialHeight As UInteger, pixelFormat As DXGI_FORMAT, alphaMode As DXGI_ALPHA_MODE, ByRef virtualSurface As IDCompositionVirtualSurface) As HRESULT
        <PreserveSig()>
        Function CreateSurfaceFromHandle(ByVal handle As IntPtr, ByRef surface As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateSurfaceFromHwnd(ByVal hwnd As IntPtr, ByRef surface As IntPtr) As HRESULT

        ' Note: many transform/effect creation methods are returned as raw IntPtr here
        <PreserveSig()>
        Function CreateTranslateTransform(ByRef translateTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateScaleTransform(ByRef scaleTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateRotateTransform(ByRef rotateTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateSkewTransform(ByRef skewTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateMatrixTransform(ByRef matrixTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateTransformGroup(transforms As IntPtr, elements As UInteger, ByRef transformGroup As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateTranslateTransform3D(ByRef translateTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateScaleTransform3D(ByRef scaleTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateRotateTransform3D(ByRef rotateTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateMatrixTransform3D(ByRef matrixTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateTransform3DGroup(transforms3D As IntPtr, elements As UInteger, ByRef transform3DGroup As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateEffectGroup(ByRef effectGroup As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateRectangleClip(ByRef clip As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateAnimation(ByRef animation As IntPtr) As HRESULT
        <PreserveSig()>
        Function CheckDeviceState(ByRef pfValid As Boolean) As HRESULT
    End Interface

    <ComImport()>
    <Guid("75F6468D-1B8E-447C-9BC6-75FEA80B5B25")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDCompositionDevice2
        <PreserveSig()>
        Function Commit() As HRESULT
        <PreserveSig()>
        Function WaitForCommitCompletion() As HRESULT
        <PreserveSig()>
        Function GetFrameStatistics(ByRef statistics As DCOMPOSITION_FRAME_STATISTICS) As HRESULT
        <PreserveSig()>
        Function CreateVisual(ByRef visual As IDCompositionVisual2) As HRESULT
        <PreserveSig()>
        Function CreateSurfaceFactory(renderingDevice As IntPtr, ByRef surfaceFactory As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateSurface(width As UInteger, height As UInteger, pixelFormat As DXGI_FORMAT, alphaMode As DXGI_ALPHA_MODE, ByRef surface As IDCompositionSurface) As HRESULT
        <PreserveSig()>
        Function CreateVirtualSurface(initialWidth As UInteger, initialHeight As UInteger, pixelFormat As DXGI_FORMAT, alphaMode As DXGI_ALPHA_MODE, ByRef virtualSurface As IDCompositionVirtualSurface) As HRESULT
        <PreserveSig()>
        Function CreateTranslateTransform(ByRef translateTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateScaleTransform(ByRef scaleTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateRotateTransform(ByRef rotateTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateSkewTransform(ByRef skewTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateMatrixTransform(ByRef matrixTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateTransformGroup(transforms As IntPtr, elements As UInteger, ByRef transformGroup As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateTranslateTransform3D(ByRef translateTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateScaleTransform3D(ByRef scaleTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateRotateTransform3D(ByRef rotateTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateMatrixTransform3D(ByRef matrixTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateTransform3DGroup(transforms3D As IntPtr, elements As UInteger, ByRef transform3DGroup As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateEffectGroup(ByRef effectGroup As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateRectangleClip(ByRef clip As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateAnimation(ByRef animation As IntPtr) As HRESULT
    End Interface

    <ComImport()>
    <Guid("5F4633FE-1E08-4CB8-8C75-CE24333F5602")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDCompositionDesktopDevice
        Inherits IDCompositionDevice2

        <PreserveSig()>
        Shadows Function Commit() As HRESULT
        <PreserveSig()>
        Shadows Function WaitForCommitCompletion() As HRESULT
        <PreserveSig()>
        Shadows Function GetFrameStatistics(ByRef statistics As DCOMPOSITION_FRAME_STATISTICS) As HRESULT
        <PreserveSig()>
        Shadows Function CreateVisual(ByRef visual As IDCompositionVisual2) As HRESULT
        <PreserveSig()>
        Shadows Function CreateSurfaceFactory(renderingDevice As IntPtr, ByRef surfaceFactory As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateSurface(width As UInteger, height As UInteger, pixelFormat As DXGI_FORMAT, alphaMode As DXGI_ALPHA_MODE, ByRef surface As IDCompositionSurface) As HRESULT
        <PreserveSig()>
        Shadows Function CreateVirtualSurface(initialWidth As UInteger, initialHeight As UInteger, pixelFormat As DXGI_FORMAT, alphaMode As DXGI_ALPHA_MODE, ByRef virtualSurface As IDCompositionVirtualSurface) As HRESULT
        <PreserveSig()>
        Shadows Function CreateTranslateTransform(ByRef translateTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateScaleTransform(ByRef scaleTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateRotateTransform(ByRef rotateTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateSkewTransform(ByRef skewTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateMatrixTransform(ByRef matrixTransform As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateTransformGroup(transforms As IntPtr, elements As UInteger, ByRef transformGroup As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateTranslateTransform3D(ByRef translateTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateScaleTransform3D(ByRef scaleTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateRotateTransform3D(ByRef rotateTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateMatrixTransform3D(ByRef matrixTransform3D As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateTransform3DGroup(transforms3D As IntPtr, elements As UInteger, ByRef transform3DGroup As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateEffectGroup(ByRef effectGroup As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateRectangleClip(ByRef clip As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function CreateAnimation(ByRef animation As IntPtr) As HRESULT

        <PreserveSig()>
        Function CreateTargetForHwnd(hwnd As IntPtr, topmost As Boolean, ByRef target As IDCompositionTarget) As HRESULT
        <PreserveSig()>
        Function CreateSurfaceFromHandle(handle As IntPtr, ByRef surface As IntPtr) As HRESULT
        <PreserveSig()>
        Function CreateSurfaceFromHwnd(hwnd As IntPtr, ByRef surface As IntPtr) As HRESULT
    End Interface

    <ComImport()>
    <Guid("eacdd04c-117e-4e17-88f4-d1b12b0e3d89")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDCompositionTarget
        <PreserveSig()>
        Function SetRoot(visual As IDCompositionVisual) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DCOMPOSITION_FRAME_STATISTICS
        Public lastFrameTime As LARGE_INTEGER
        Public currentCompositionRate As DXGI_RATIONAL
        Public currentTime As LARGE_INTEGER
        Public timeFrequency As LARGE_INTEGER
        Public nextEstimatedFrameTime As LARGE_INTEGER
    End Structure

    <ComImport()>
    <Guid("4d93059d-097b-4651-9a60-f0f25116e2f3")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDCompositionVisual
        <PreserveSig()>
        Function _SetOffsetX(animation As IntPtr) As HRESULT
        <PreserveSig()>
        Function SetOffsetX(offsetX As Single) As HRESULT

        <PreserveSig()>
        Function _SetOffsetY(animation As IntPtr) As HRESULT
        <PreserveSig()>
        Function SetOffsetY(offsetY As Single) As HRESULT

        <PreserveSig()>
        Function _SetTransform(transform As IntPtr) As HRESULT
        <PreserveSig()>
        Function SetTransform(ByRef matrix As D2D_MATRIX_3X2_F) As HRESULT

        <PreserveSig()>
        Function SetTransformParent(visual As IDCompositionVisual) As HRESULT

        <PreserveSig()>
        Function SetEffect(effect As IntPtr) As HRESULT

        <PreserveSig()>
        Function SetBitmapInterpolationMode(interpolationMode As DCOMPOSITION_BITMAP_INTERPOLATION_MODE) As HRESULT

        <PreserveSig()>
        Function SetBorderMode(borderMode As DCOMPOSITION_BORDER_MODE) As HRESULT

        <PreserveSig()>
        Function _SetClip(clip As IntPtr) As HRESULT
        <PreserveSig()>
        Function SetClip(ByRef rect As D2D_RECT_F) As HRESULT

        <PreserveSig()>
        Function SetContent(content As IntPtr) As HRESULT

        <PreserveSig()>
        Function AddVisual(visual As IDCompositionVisual, insertAbove As Boolean, referenceVisual As IDCompositionVisual) As HRESULT

        <PreserveSig()>
        Function RemoveVisual(visual As IDCompositionVisual) As HRESULT

        <PreserveSig()>
        Function RemoveAllVisuals() As HRESULT

        <PreserveSig()>
        Function SetCompositeMode(compositeMode As DCOMPOSITION_COMPOSITE_MODE) As HRESULT
    End Interface

    <ComImport()>
    <Guid("E8DE1639-4331-4B26-BC5F-6A321D347A85")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDCompositionVisual2
        Inherits IDCompositionVisual

        <PreserveSig()>
        Shadows Function _SetOffsetX(animation As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function SetOffsetX(offsetX As Single) As HRESULT

        <PreserveSig()>
        Shadows Function _SetOffsetY(animation As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function SetOffsetY(offsetY As Single) As HRESULT

        <PreserveSig()>
        Shadows Function _SetTransform(transform As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function SetTransform(ByRef matrix As D2D_MATRIX_3X2_F) As HRESULT

        <PreserveSig()>
        Shadows Function SetTransformParent(visual As IDCompositionVisual) As HRESULT

        <PreserveSig()>
        Shadows Function SetEffect(effect As IntPtr) As HRESULT

        <PreserveSig()>
        Shadows Function SetBitmapInterpolationMode(interpolationMode As DCOMPOSITION_BITMAP_INTERPOLATION_MODE) As HRESULT

        <PreserveSig()>
        Shadows Function SetBorderMode(borderMode As DCOMPOSITION_BORDER_MODE) As HRESULT

        <PreserveSig()>
        Shadows Function _SetClip(clip As IntPtr) As HRESULT
        <PreserveSig()>
        Shadows Function SetClip(ByRef rect As D2D_RECT_F) As HRESULT

        <PreserveSig()>
        Shadows Function SetContent(content As IntPtr) As HRESULT

        <PreserveSig()>
        Shadows Function AddVisual(visual As IDCompositionVisual, insertAbove As Boolean, referenceVisual As IDCompositionVisual) As HRESULT

        <PreserveSig()>
        Shadows Function RemoveVisual(visual As IDCompositionVisual) As HRESULT

        <PreserveSig()>
        Shadows Function RemoveAllVisuals() As HRESULT

        <PreserveSig()>
        Shadows Function SetCompositeMode(compositeMode As DCOMPOSITION_COMPOSITE_MODE) As HRESULT

        <PreserveSig()>
        Function SetOpacityMode(mode As DCOMPOSITION_OPACITY_MODE) As HRESULT

        <PreserveSig()>
        Function SetBackFaceVisibility(visibility As DCOMPOSITION_BACKFACE_VISIBILITY) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D_MATRIX_3X2_F
        Public m11 As Single
        Public m12 As Single
        Public m21 As Single
        Public m22 As Single
        Public dx As Single
        Public dy As Single
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure D2D_RECT_F
        Public left As Single
        Public top As Single
        Public right As Single
        Public bottom As Single
    End Structure

    Public Enum DCOMPOSITION_BITMAP_INTERPOLATION_MODE As UInteger
        DCOMPOSITION_BITMAP_INTERPOLATION_MODE_NEAREST_NEIGHBOR = 0
        DCOMPOSITION_BITMAP_INTERPOLATION_MODE_LINEAR = 1
        DCOMPOSITION_BITMAP_INTERPOLATION_MODE_INHERIT = &HFFFFFFFFUI
    End Enum

    Public Enum DCOMPOSITION_BORDER_MODE As UInteger
        DCOMPOSITION_BORDER_MODE_SOFT = 0
        DCOMPOSITION_BORDER_MODE_HARD = 1
        DCOMPOSITION_BORDER_MODE_INHERIT = &HFFFFFFFFUI
    End Enum

    Public Enum DCOMPOSITION_COMPOSITE_MODE As UInteger
        DCOMPOSITION_COMPOSITE_MODE_SOURCE_OVER = 0
        DCOMPOSITION_COMPOSITE_MODE_DESTINATION_INVERT = 1
        DCOMPOSITION_COMPOSITE_MODE_MIN_BLEND = 2
        DCOMPOSITION_COMPOSITE_MODE_INHERIT = &HFFFFFFFFUI
    End Enum

    Public Enum DCOMPOSITION_OPACITY_MODE As UInteger
        DCOMPOSITION_OPACITY_MODE_LAYER = 0
        DCOMPOSITION_OPACITY_MODE_MULTIPLY = 1
        DCOMPOSITION_OPACITY_MODE_INHERIT = &HFFFFFFFFUI
    End Enum

    Public Enum DCOMPOSITION_BACKFACE_VISIBILITY As UInteger
        DCOMPOSITION_BACKFACE_VISIBILITY_VISIBLE = 0
        DCOMPOSITION_BACKFACE_VISIBILITY_HIDDEN = 1
        DCOMPOSITION_BACKFACE_VISIBILITY_INHERIT = &HFFFFFFFFUI
    End Enum

    <ComImport()>
    <Guid("BB8A4953-2C99-4F5A-96F5-4819027FA3AC")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDCompositionSurface
        <PreserveSig()>
        Function BeginDraw(updateRect As IntPtr, ByRef iid As Guid, ByRef updateObject As IntPtr, ByRef updateOffset As POINT) As HRESULT
        <PreserveSig()>
        Function EndDraw() As HRESULT
        <PreserveSig()>
        Function SuspendDraw() As HRESULT
        <PreserveSig()>
        Function ResumeDraw() As HRESULT
        <PreserveSig()>
        Function Scroll(ByRef scrollRect As RECT, ByRef clipRect As RECT, offsetX As Integer, offsetY As Integer) As HRESULT
    End Interface

    <ComImport()>
    <Guid("AE471C51-5F53-4A24-8D3E-D0C39C30B3F0")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDCompositionVirtualSurface
        Inherits IDCompositionSurface

        <PreserveSig()>
        Shadows Function BeginDraw(updateRect As IntPtr, ByRef iid As Guid, ByRef updateObject As IntPtr, ByRef updateOffset As POINT) As HRESULT

        <PreserveSig()>
        Shadows Function EndDraw() As HRESULT
        <PreserveSig()>
        Shadows Function SuspendDraw() As HRESULT
        <PreserveSig()>
        Shadows Function ResumeDraw() As HRESULT
        <PreserveSig()>
        Shadows Function Scroll(ByRef scrollRect As RECT, ByRef clipRect As RECT, offsetX As Integer, offsetY As Integer) As HRESULT

        <PreserveSig()>
        Function Resize(width As UInteger, height As UInteger) As HRESULT

        <PreserveSig()>
        Function Trim(ByRef rectangles As RECT, count As UInteger) As HRESULT
    End Interface

End Namespace
