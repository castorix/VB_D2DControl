Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports Direct2D
Imports Direct2D.D2DTools
Imports DWrite
Imports DXGI
Imports DXGI.DXGITools
Imports GlobalStructures
Imports GlobalStructures.GlobalTools
Imports VB_D2DControl.D2DControl
Imports VB_D2DControl.DirectComposition
Imports WIC

Public Class D2DControl

    Public Const WS_EX_LAYERED As Integer = &H80000

    Public Const WM_CREATE = &H1
    Public Const WM_DESTROY = &H2
    Public Const WM_SIZE = &H5
    Public Const WM_PAINT = &HF
    Public Const WM_ERASEBKGND = &H14
    Public Const WM_TIMER = &H113
    Public Const WM_NCHITTEST = &H84

    Public Const WS_OVERLAPPED As Integer = &H0
    Public Const WS_POPUP As Integer = &H80000000
    Public Const WS_CHILD As Integer = &H40000000
    Public Const WS_MINIMIZE As Integer = &H20000000
    Public Const WS_VISIBLE As Integer = &H10000000
    Public Const WS_DISABLED As Integer = &H8000000
    Public Const WS_CLIPSIBLINGS As Integer = &H4000000
    Public Const WS_CLIPCHILDREN As Integer = &H2000000
    Public Const WS_MAXIMIZE As Integer = &H1000000
    Public Const WS_CAPTION As Integer = &HC00000
    Public Const WS_BORDER As Integer = &H800000
    Public Const WS_DLGFRAME As Integer = &H400000
    Public Const WS_VSCROLL As Integer = &H200000
    Public Const WS_HSCROLL As Integer = &H100000
    Public Const WS_SYSMENU As Integer = &H80000
    Public Const WS_THICKFRAME As Integer = &H40000
    Public Const WS_TABSTOP As Integer = &H10000
    Public Const WS_MINIMIZEBOX As Integer = &H20000
    Public Const WS_MAXIMIZEBOX As Integer = &H10000

    Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)

    Public Enum HT
        HTERROR = -2
        HTTRANSPARENT = -1
        HTNOWHERE = 0
        HTCLIENT = 1
        HTCAPTION = 2
        HTSYSMENU = 3
        HTGROWBOX = 4
        HTMENU = 5
        HTHSCROLL = 6
        HTVSCROLL = 7
        HTMINBUTTON = 8
        HTMAXBUTTON = 9
        HTLEFT = 10
        HTRIGHT = 11
        HTTOP = 12
        HTTOPLEFT = 13
        HTTOPRIGHT = 14
        HTBOTTOM = 15
        HTBOTTOMLEFT = 16
        HTBOTTOMRIGHT = 17
        HTBORDER = 18
        HTOBJECT = 19
        HTCLOSE = 20
        HTHELP = 21
    End Enum

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function DefWindowProc(hWnd As IntPtr, uMsg As UInteger, wParam As Integer, lParam As IntPtr) As Integer
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function BeginPaint(ByVal hWnd As IntPtr, ByRef lpPaint As PAINTSTRUCT) As IntPtr
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function EndPaint(ByVal hWnd As IntPtr, ByRef lpPaint As PAINTSTRUCT) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PAINTSTRUCT
        Public hdc As IntPtr
        Public fErase As Boolean
        Public rcPaint_left As Integer
        Public rcPaint_top As Integer
        Public rcPaint_right As Integer
        Public rcPaint_bottom As Integer
        Public fRestore As Boolean
        Public fIncUpdate As Boolean
        Public reserved1 As Integer
        Public reserved2 As Integer
        Public reserved3 As Integer
        Public reserved4 As Integer
        Public reserved5 As Integer
        Public reserved6 As Integer
        Public reserved7 As Integer
        Public reserved8 As Integer
    End Structure

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function InvalidateRect(hWnd As IntPtr, ByRef lpRect As RECT, bErase As Boolean) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function InvalidateRect(hWnd As IntPtr, lpRect As IntPtr, bErase As Boolean) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function SetTimer(hWnd As IntPtr, nIDEvent As IntPtr, uElapse As UInteger, lpTimerFunc As IntPtr) As IntPtr
    End Function

    <DllImport("User32.dll", SetLastError:=True, ExactSpelling:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function RedrawWindow(hWnd As IntPtr, ByRef lprcUpdate As RECT, hrgnUpdate As IntPtr, flags As UInteger) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True, ExactSpelling:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function RedrawWindow(hWnd As IntPtr, lprcUpdate As IntPtr, hrgnUpdate As IntPtr, flags As UInteger) As Boolean
    End Function

    Public Const RDW_INVALIDATE As Integer = &H1, RDW_INTERNALPAINT As Integer = &H2, RDW_ERASE As Integer = &H4,
        RDW_VALIDATE As Integer = &H8, RDW_NOINTERNALPAINT As Integer = &H10, RDW_NOERASE As Integer = &H20,
        RDW_NOCHILDREN As Integer = &H40, RDW_ALLCHILDREN As Integer = &H80, RDW_UPDATENOW As Integer = &H100,
        RDW_ERASENOW As Integer = &H200, RDW_FRAME As Integer = &H400, RDW_NOFRAME As Integer = &H800

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function SetLayeredWindowAttributes(hwnd As IntPtr, crKey As UInteger, bAlpha As Byte, dwFlags As UInteger) As Boolean
    End Function

    Public Const LWA_COLORKEY As UInteger = &H1UI
    Public Const LWA_ALPHA As UInteger = &H2UI


    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function UpdateLayeredWindow(hwnd As IntPtr, hdcDst As IntPtr, pptDst As IntPtr, psize As IntPtr, hdcSrc As IntPtr, pprSrc As IntPtr, crKey As Integer, ByRef pblend As BLENDFUNCTION, dwFlags As Integer) As Boolean
    End Function

    Public Const ULW_COLORKEY As Integer = &H1
    Public Const ULW_ALPHA As Integer = &H2
    Public Const ULW_OPAQUE As Integer = &H4

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Public Structure BLENDFUNCTION
        Public BlendOp As Byte
        Public BlendFlags As Byte
        Public SourceConstantAlpha As Byte
        Public AlphaFormat As Byte
    End Structure

    Public Const AC_SRC_OVER = &H0
    Public Const AC_SRC_ALPHA = &H1

    <DllImport("Gdi32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function GetObject(hFont As IntPtr, nSize As Integer, <Out> ByRef bm As BITMAP) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure BITMAP
        Public bmType As Integer
        Public bmWidth As Integer
        Public bmHeight As Integer
        Public bmWidthBytes As Integer
        Public bmPlanes As Short
        Public bmBitsPixel As Short
        Public bmBits As IntPtr
    End Structure

    <DllImport("Gdi32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function CreateCompatibleDC(hDC As IntPtr) As IntPtr
    End Function

    <DllImport("Gdi32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function SelectObject(hDC As IntPtr, hObject As IntPtr) As IntPtr
    End Function

    <DllImport("Gdi32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function DeleteObject(hObject As IntPtr) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function GetDC(hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function ReleaseDC(hWnd As IntPtr, hDC As IntPtr) As Integer
    End Function

    <DllImport("Gdi32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function DeleteDC(hdc As IntPtr) As Boolean
    End Function

    <DllImport("Gdi32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function CreateDIBSection(hdc As IntPtr, ByRef pbmi As BITMAPINFO, usage As UInteger,
                                            ByRef ppvBits As IntPtr, hSection As IntPtr, offset As Integer) As IntPtr
    End Function

    <DllImport("Gdi32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function CreateDIBSection(hdc As IntPtr, ByRef pbmi As BITMAPV5HEADER, usage As UInteger,
                                            ByRef ppvBits As IntPtr, hSection As IntPtr, offset As Integer) As IntPtr
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure BITMAPINFOHEADER
        Public biSize As Integer
        Public biWidth As Integer
        Public biHeight As Integer
        Public biPlanes As Short
        Public biBitCount As Short
        Public biCompression As Integer
        Public biSizeImage As Integer
        Public biXPelsPerMeter As Integer
        Public biYPelsPerMeter As Integer
        Public biClrUsed As Integer
        Public biClrImportant As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure BITMAPINFO
        <MarshalAs(UnmanagedType.Struct)>
        Public bmiHeader As BITMAPINFOHEADER

        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1024)>
        Public bmiColors As Integer()
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure BITMAPV5HEADER
        Public bV5Size As Integer
        Public bV5Width As Integer
        Public bV5Height As Integer
        Public bV5Planes As Short
        Public bV5BitCount As Short
        Public bV5Compression As Integer
        Public bV5SizeImage As Integer
        Public bV5XPelsPerMeter As Integer
        Public bV5YPelsPerMeter As Integer
        Public bV5ClrUsed As Integer
        Public bV5ClrImportant As Integer
        Public bV5RedMask As Integer
        Public bV5GreenMask As Integer
        Public bV5BlueMask As Integer
        Public bV5AlphaMask As Integer
        Public bV5CSType As Integer
        Public bV5Endpoints As CIEXYZTRIPLE
        Public bV5GammaRed As Integer
        Public bV5GammaGreen As Integer
        Public bV5GammaBlue As Integer
        Public bV5Intent As Integer
        Public bV5ProfileData As Integer
        Public bV5ProfileSize As Integer
        Public bV5Reserved As Integer
    End Structure


    <StructLayout(LayoutKind.Sequential)>
    Public Structure CIEXYZTRIPLE
        Public ciexyzRed As CIEXYZ
        Public ciexyzGreen As CIEXYZ
        Public ciexyzBlue As CIEXYZ
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure CIEXYZ
        Public ciexyzX As Integer
        Public ciexyzY As Integer
        Public ciexyzZ As Integer
    End Structure

    Public Const BI_RGB = 0
    Public Const BI_RLE8 = 1
    Public Const BI_RLE4 = 2
    Public Const BI_BITFIELDS = 3
    Public Const BI_JPEG = 4
    Public Const BI_PNG = 5

    Public Const DIB_RGB_COLORS = 0
    Public Const DIB_PAL_COLORS = 1

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function GetWindowRect(hWnd As IntPtr, <Out> ByRef lpRect As RECT) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function GetClientRect(hWnd As IntPtr, <Out> ByRef lpRect As RECT) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function ScreenToClient(hWnd As IntPtr, ByRef lpPoint As System.Drawing.Point) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function ClientToScreen(hWnd As IntPtr, ByRef lpPoint As System.Drawing.Point) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function GetParent(hWnd As IntPtr) As IntPtr
    End Function

    <ComImport>
    <Guid("db6f6ddb-ac77-4e88-8253-819df9bbf140")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ID3D11Device
        <PreserveSig>
        Function CreateBuffer(ByRef pDesc As D3D11_BUFFER_DESC, pInitialData As IntPtr, ByRef ppBuffer As ID3D11Buffer) As HRESULT
        <PreserveSig>
        Function CreateTexture1D(ByRef pDesc As IntPtr, pInitialData As IntPtr, ByRef ppTexture1D As IntPtr) As HRESULT
        'Function CreateTexture1D(ByRef pDesc As D3D11_TEXTURE1D_DESC, pInitialData As IntPtr, ByRef ppTexture1D As ID3D11Texture1D) As HRESULT
        <PreserveSig>
        Function CreateTexture2D(ByRef pDesc As DXGI.D3D11_TEXTURE2D_DESC, pInitialData As IntPtr, ByRef ppTexture2D As DXGI.ID3D11Texture2D) As HRESULT
    End Interface

    Public Enum D3D11_CPU_ACCESS_FLAG
        D3D11_CPU_ACCESS_WRITE = &H10000
        D3D11_CPU_ACCESS_READ = &H20000
    End Enum

    Private m_pD2DFactory1 As ID2D1Factory1 = Nothing
    Private m_pDWriteFactory As IDWriteFactory = Nothing
    Private m_pWICImagingFactory As IWICImagingFactory = Nothing

    Private m_pD2DMainBrush As ID2D1SolidColorBrush = Nothing

    Private m_pD3D11DevicePtr As IntPtr = IntPtr.Zero
    Private m_pD3D11DeviceContextPtr As IntPtr = IntPtr.Zero

    Private m_pD3D11Device As ID3D11Device = Nothing
    Private m_pD3D11DeviceContext As ID3D11DeviceContext = Nothing

    Private m_pDXGIDevice As IDXGIDevice1 = Nothing
    Private m_pD2DTargetBitmap As ID2D1Bitmap1 = Nothing
    Private m_pDXGISwapChain1 As IDXGISwapChain1 = Nothing

    Private m_pD2DDevice As ID2D1Device = Nothing
    Private m_pD2DDeviceContext As ID2D1DeviceContext = Nothing

    Private m_pD2DHwndRenderTarget As ID2D1HwndRenderTarget = Nothing

    Public m_pD2DBitmap As ID2D1Bitmap1 = Nothing

    Private m_pD2DDCRenderTarget As ID2D1DCRenderTarget = Nothing


    Public Sub New()
        InitializeComponent()
        Me.SetStyle(ControlStyles.DoubleBuffer, True)
    End Sub

    'Public Sub Initialize(pD2DFactory1 As ID2D1Factory1, pDWriteFactory As IDWriteFactory, pWICImagingFactory As IWICImagingFactory)
    Public Sub Initialize(pD2DFactory1 As ID2D1Factory1, pWICImagingFactory As IWICImagingFactory)
        Dim hr As HRESULT = HRESULT.S_OK
        m_pD2DFactory1 = pD2DFactory1

        Dim rtp = RenderTargetProperties(D2D1_RENDER_TARGET_TYPE.D2D1_RENDER_TARGET_TYPE_DEFAULT,
                               PixelFormat(DXGI_FORMAT.DXGI_FORMAT_UNKNOWN, D2D1_ALPHA_MODE.D2D1_ALPHA_MODE_PREMULTIPLIED))
        Dim hrtp = HwndRenderTargetProperties(Me.Handle, SizeU(Me.ClientSize.Width, Me.ClientSize.Height), D2D1_PRESENT_OPTIONS.D2D1_PRESENT_OPTIONS_IMMEDIATELY)

        hr = m_pD2DFactory1.CreateHwndRenderTarget(rtp, hrtp, m_pD2DHwndRenderTarget)

        rtp = RenderTargetProperties(D2D1_RENDER_TARGET_TYPE.D2D1_RENDER_TARGET_TYPE_DEFAULT,
                               PixelFormat(DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE.D2D1_ALPHA_MODE_PREMULTIPLIED))
        ' -2003292288 &H88982F80 UNSUPPORTED_PIXEL_FORMAT
        hr = m_pD2DFactory1.CreateDCRenderTarget(rtp, m_pD2DDCRenderTarget)

        'm_pDWriteFactory = pDWriteFactory
        m_pWICImagingFactory = pWICImagingFactory
        hr = CreateD3D11Device()
        hr = CreateDeviceResources()
        hr = CreateSwapChain(IntPtr.Zero)
        If (SUCCEEDED(hr)) Then
            hr = ConfigureSwapChain()
            hr = CreateDirectComposition(Me.Handle)
        End If
        SetTimer(Me.Handle, 1, 15, IntPtr.Zero)
    End Sub

    Public Function CreateD3D11Device() As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        Dim creationFlags = D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_BGRA_SUPPORT
#If DEBUG Then
        creationFlags = creationFlags Or D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_DEBUG
#End If

        Dim aD3D_FEATURE_LEVEL = New Integer() {D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_11_1, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_11_0, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_10_1, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_10_0, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_9_3, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_9_2, D3D_FEATURE_LEVEL.D3D_FEATURE_LEVEL_9_1}

        Dim featureLevel As D3D_FEATURE_LEVEL
        hr = DXGITools.D3D11CreateDevice(Nothing, D3D_DRIVER_TYPE.D3D_DRIVER_TYPE_HARDWARE, IntPtr.Zero, creationFlags, aD3D_FEATURE_LEVEL, aD3D_FEATURE_LEVEL.Length, DXGITools.D3D11_SDK_VERSION, m_pD3D11DevicePtr, featureLevel, m_pD3D11DeviceContextPtr)    ' specify null to use the default adapter
        If hr = HRESULT.S_OK Then

            m_pD3D11Device = TryCast(Marshal.GetObjectForIUnknown(m_pD3D11DevicePtr), ID3D11Device)
            m_pD3D11DeviceContext = TryCast(Marshal.GetObjectForIUnknown(m_pD3D11DeviceContextPtr), ID3D11DeviceContext)

            m_pDXGIDevice = TryCast(Marshal.GetObjectForIUnknown(m_pD3D11DevicePtr), IDXGIDevice1)
            hr = m_pD2DFactory1.CreateDevice(m_pDXGIDevice, m_pD2DDevice)
            If hr = HRESULT.S_OK Then
                hr = m_pD2DDevice.CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS.D2D1_DEVICE_CONTEXT_OPTIONS_NONE, m_pD2DDeviceContext)
                Marshal.ReleaseComObject(m_pD2DDevice)
            End If
            Marshal.Release(m_pD3D11DevicePtr)
            Marshal.Release(m_pD3D11DeviceContextPtr)
        End If
        Return hr
    End Function

    Private Function CreateSwapChain(hWnd As IntPtr) As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        Dim swapChainDesc As DXGI_SWAP_CHAIN_DESC1 = New DXGI_SWAP_CHAIN_DESC1()
        'swapChainDesc.Width = 1
        'swapChainDesc.Height = 1
        swapChainDesc.Width = Me.ClientSize.Width
        swapChainDesc.Height = Me.ClientSize.Height
        swapChainDesc.Format = DXGI.DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM ' this is the most common swapchain format
        swapChainDesc.Stereo = False
        swapChainDesc.SampleDesc.Count = 1                ' don't use multi-sampling
        swapChainDesc.SampleDesc.Quality = 0
        swapChainDesc.BufferUsage = D2DTools.DXGI_USAGE_RENDER_TARGET_OUTPUT
        swapChainDesc.BufferCount = 2                     ' use double buffering to enable flip
        swapChainDesc.Scaling = If(hWnd <> IntPtr.Zero, DXGI_SCALING.DXGI_SCALING_NONE, DXGI_SCALING.DXGI_SCALING_STRETCH)
        swapChainDesc.SwapEffect = DXGI_SWAP_EFFECT.DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL ' all apps must use this SwapEffect       
        swapChainDesc.Flags = 0

        Dim pDXGIAdapter As IDXGIAdapter = Nothing
        hr = m_pDXGIDevice.GetAdapter(pDXGIAdapter)
        If hr = HRESULT.S_OK Then
            Dim pDXGIFactory2Ptr As IntPtr
            hr = pDXGIAdapter.GetParent(GetType(IDXGIFactory2).GUID, pDXGIFactory2Ptr)
            If hr = HRESULT.S_OK Then
                Dim pDXGIFactory2 As IDXGIFactory2 = TryCast(Marshal.GetObjectForIUnknown(pDXGIFactory2Ptr), IDXGIFactory2)
                If hWnd <> IntPtr.Zero Then
                    hr = pDXGIFactory2.CreateSwapChainForHwnd(m_pD3D11DevicePtr, hWnd, swapChainDesc, IntPtr.Zero, Nothing, m_pDXGISwapChain1)
                Else
                    ' For Composition SwapChain
                    swapChainDesc.AlphaMode = DXGI_ALPHA_MODE.DXGI_ALPHA_MODE_PREMULTIPLIED
                    hr = pDXGIFactory2.CreateSwapChainForComposition(m_pD3D11DevicePtr, swapChainDesc, Nothing, m_pDXGISwapChain1)
                End If
                hr = m_pDXGIDevice.SetMaximumFrameLatency(1)
                SafeRelease(pDXGIFactory2)
                Marshal.Release(pDXGIFactory2Ptr)
            End If
            SafeRelease(pDXGIAdapter)
        End If
        Return hr
    End Function

    Private Function ConfigureSwapChain() As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK

        Dim bitmapProperties As D2D1_BITMAP_PROPERTIES1 = New D2D1_BITMAP_PROPERTIES1()
        bitmapProperties.bitmapOptions = D2D1_BITMAP_OPTIONS.D2D1_BITMAP_OPTIONS_TARGET Or D2D1_BITMAP_OPTIONS.D2D1_BITMAP_OPTIONS_CANNOT_DRAW
        bitmapProperties.pixelFormat = D2DTools.PixelFormat(DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE.D2D1_ALPHA_MODE_PREMULTIPLIED)

        Dim nDPI As UInteger = GetDpiForWindow(Me.Handle)
        bitmapProperties.dpiX = nDPI
        bitmapProperties.dpiY = nDPI

        Dim pDXGISurfacePtr = IntPtr.Zero
        hr = m_pDXGISwapChain1.GetBuffer(0, GetType(IDXGISurface).GUID, pDXGISurfacePtr)
        If hr = HRESULT.S_OK Then
            Dim pDXGISurface As IDXGISurface = TryCast(Marshal.GetObjectForIUnknown(pDXGISurfacePtr), IDXGISurface)
            hr = m_pD2DDeviceContext.CreateBitmapFromDxgiSurface(pDXGISurface, bitmapProperties, m_pD2DTargetBitmap)
            If hr = HRESULT.S_OK Then
                m_pD2DDeviceContext.SetTarget(m_pD2DTargetBitmap)
            End If
            SafeRelease(pDXGISurface)
            Marshal.Release(pDXGISurfacePtr)
        End If
        Return hr
    End Function

    Private m_pDCompositionDevice As IDCompositionDevice = Nothing
    Private m_pDCompositionTarget As IDCompositionTarget = Nothing
    Private m_pDCompositionVisual As IDCompositionVisual = Nothing

    Private Function CreateDirectComposition(hWnd As IntPtr) As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        Dim pDCompositionDevicePtr As IntPtr = IntPtr.Zero
        hr = DirectCompositionTools.DCompositionCreateDevice2(m_pDXGIDevice, GetType(IDCompositionDevice).GUID, pDCompositionDevicePtr)
        If hr = HRESULT.S_OK Then
            m_pDCompositionDevice = CType(Marshal.GetObjectForIUnknown(pDCompositionDevicePtr), IDCompositionDevice)
            hr = m_pDCompositionDevice.CreateTargetForHwnd(hWnd, True, m_pDCompositionTarget)
            If hr = HRESULT.S_OK Then
                hr = m_pDCompositionDevice.CreateVisual(m_pDCompositionVisual)
                If hr = HRESULT.S_OK Then
                    Dim pDXGISwapChain1Ptr As IntPtr = Marshal.GetIUnknownForObject(m_pDXGISwapChain1)

                    'Dim pSurface As IntPtr
                    'hr = m_pDCompositionDevice.CreateSurfaceFromHwnd(Me.Handle, pSurface)

                    hr = m_pDCompositionVisual.SetContent(pDXGISwapChain1Ptr)
                    'hr = m_pDCompositionVisual.SetContent(pSurface)

                    hr = m_pDCompositionTarget.SetRoot(m_pDCompositionVisual)
                    Marshal.Release(pDXGISwapChain1Ptr)
                End If
            End If
            Marshal.Release(pDCompositionDevicePtr)
            hr = m_pDCompositionDevice.Commit()
        End If
        Return hr
    End Function


    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            'Dim bDesignMode As Boolean = (System.ComponentModel.LicenseManager.UsageMode = System.ComponentModel.LicenseUsageMode.Designtime)
            Dim bDesignMode = Me.DesignMode
            If Not bDesignMode Then
                cp.ExStyle = cp.ExStyle Or WS_EX_LAYERED
                'cp.Style = cp.Style - WS_CHILD
            End If
            Return cp
        End Get
    End Property

    Private m_bUpdated = False

    Protected Overrides Sub WndProc(ByRef m As Message)
        If (m.Msg = WM_CREATE) Then
            'Me.BackColor = Color.FromArgb(255, 0, 255, 1)
            'Me.BackColor = Color.FromArgb(255, 255, 0, 255)
            'MakeWindowTransparent(m.HWnd)
            'Dim bReturn = SetLayeredWindowAttributes(m.HWnd, RGB(0, 0, 0), 255, LWA_COLORKEY)
            'Dim bReturn = SetLayeredWindowAttributes(m.HWnd, RGB(255, 0, 255), 255, LWA_COLORKEY)

            ' Marche avec Backcolor = Red
            'Dim bReturn = SetLayeredWindowAttributes(m.HWnd, RGB(255, 0, 255), 255, LWA_COLORKEY)

            'Dim bReturn = SetLayeredWindowAttributes(m.HWnd, RGB(255, 0, 0), 255, LWA_COLORKEY)

            'Dim hBitmap As IntPtr = LoadImage("E:\Sources\VB_D2DControl\Assets\Butterfly.png", System.Drawing.Color.FromArgb(0))
            'Me.SetBitmap(hBitmap)

            'Dim bReturn = SetLayeredWindowAttributes(m.HWnd, RGB(255, 255, 255), 255, LWA_COLORKEY)
            'SetLayeredWindowAttributes(m.HWnd, 0, 255, LWA_ALPHA)
            'SetOpacity(50)
            Me.BringToFront()
            Return
        ElseIf (m.Msg = WM_PAINT) Then
            Dim hr As HRESULT = OnPaintProc(m.HWnd)
            m.Result = CType(hr, IntPtr)
            'ElseIf (m.Msg = WM_ERASEBKGND) Then
            '    m.Result = CType(1, IntPtr)
        ElseIf (m.Msg = WM_SIZE) Then
            Dim hr As HRESULT = OnResizeProc(m)
            m.Result = IntPtr.Zero
        ElseIf (m.Msg = WM_NCHITTEST) Then
            Dim nRet = DefWindowProc(m.HWnd, WM_NCHITTEST, 0, m.LParam)
            Select Case nRet
                Case HT.HTLEFT
                Case HT.HTTOP
                Case HT.HTBOTTOM
                Case HT.HTRIGHT
                Case HT.HTBOTTOMLEFT
                Case HT.HTBOTTOMRIGHT
                Case HT.HTTOPLEFT
                Case HT.HTTOPRIGHT
                    m.Result = nRet
                Case Else
                    m.Result = HT.HTCAPTION
            End Select
            Return
        ElseIf (m.Msg = WM_TIMER) Then
            RedrawWindow(m.HWnd, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE)
        ElseIf (m.Msg = WM_DESTROY) Then
            Clean()
        Else
            MyBase.WndProc(m)
        End If
    End Sub

    Public Function OnPaintProc(hWnd As IntPtr) As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        Dim ps As PAINTSTRUCT = New PAINTSTRUCT
        If (BeginPaint(hWnd, ps) <> IntPtr.Zero) Then

            If (m_pD2DDeviceContext IsNot Nothing) Then
                m_pD2DDeviceContext.BeginDraw()

                Dim size As D2D1_SIZE_F
                m_pD2DDeviceContext.GetSize(size)

                'm_pD2DDeviceContext.Clear(New ColorF(ColorF.Enum.Red, 0.0F))
                'm_pD2DDeviceContext.Clear(New ColorF(ColorF.Enum.Red, 1.0F))
                'm_pD2DDeviceContext.Clear(New ColorF(ColorF.Enum.Black, 0.0F))

                'm_pD2DDeviceContext.Clear(New ColorF(ColorF.Enum.Black, 1.0F))

                'm_pD2DDeviceContext.Clear(New ColorF(1.0F, 0F, 1.0F, 1.0F))

                m_pD2DDeviceContext.Clear(New ColorF(1.0F, 0F, 0.0F, 0.0F))

                'Dim pTransformIdentity As D2D1_MATRIX_3X2_F = New D2D1_MATRIX_3X2_F()
                'pTransformIdentity = Matrix3x2F.Identity()
                'm_pD2DDeviceContext.SetTransform(pTransformIdentity)

                Dim centerX As Single = size.width / 2.0F
                Dim centerY As Single = size.height / 2.0F

                ' Ellipse radii (adjust as you want)
                Dim radiusX As Single = 200.0F
                Dim radiusY As Single = 200.0F

                m_pD2DDeviceContext.FillEllipse(Ellipse(New Direct2D.D2D1_POINT_2F(centerX, centerY), radiusX, radiusY), m_pD2DMainBrush)

                'm_pD2DHwndRenderTarget.FillEllipse(Ellipse(New Direct2D.D2D1_POINT_2F(centerX, centerY), radiusX, radiusY), m_pD2DMainBrush)
                If m_pD2DBitmap IsNot Nothing Then

                    If (Not m_bUpdated) Then
                        'SetD2DBitmapToLayeredWindow(hWnd, m_pD2DBitmap)

                        UpdateLayeredFromD2DBitmap(hWnd, m_pD2DBitmap, m_pD2DFactory1, m_pWICImagingFactory)

                        m_bUpdated = True
                    End If

                    Dim sizeBitmap As D2D1_SIZE_F
                    m_pD2DBitmap.GetSize(sizeBitmap)
                    Dim destRectBitmap As D2D1_RECT_F = New D2D1_RECT_F(0.0F, 0.0F, size.width, size.height)
                    'Dim destRectBitmap As D2D1_RECT_F = New D2D1_RECT_F(0.0F, 0.0F, sizeBitmap.width, sizeBitmap.height)
                    Dim sourceRectBitmap As D2D1_RECT_F = New D2D1_RECT_F(0.0F, 0.0F, sizeBitmap.width, sizeBitmap.height)
                    m_pD2DDeviceContext.DrawBitmap(m_pD2DBitmap, destRectBitmap, 1.0F, D2D1_BITMAP_INTERPOLATION_MODE.D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, sourceRectBitmap)
                End If

                'Dim rect = New D2D1_RECT_F(0, 0, size.width, size.height)
                'm_pD2DDeviceContext.FillRectangle(rect, m_pD2DMainBrush)

                Dim tag1 As ULong, tag2 As ULong = 0
                hr = m_pD2DDeviceContext.EndDraw(tag1, tag2)
                If CUInt(hr) = D2DTools.D2DERR_RECREATE_TARGET Then
                    m_pD2DDeviceContext.SetTarget(Nothing)
                    SafeRelease(m_pD2DDeviceContext)
                    hr = CreateD3D11Device()
                    hr = CreateDeviceResources()
                    hr = CreateSwapChain(IntPtr.Zero)
                    If (SUCCEEDED(hr)) Then
                        hr = ConfigureSwapChain()
                        hr = CreateDirectComposition(Me.Handle)
                    End If
                End If
                hr = m_pDXGISwapChain1.Present(1, 0)

            End If

            'If (m_pD2DHwndRenderTarget IsNot Nothing) Then
            '    m_pD2DHwndRenderTarget.BeginDraw()

            '    'm_pD2DHwndRenderTarget.Clear(New ColorF(ColorF.Enum.Blue, 1.0F))
            '    'm_pD2DHwndRenderTarget.Clear(New ColorF(ColorF.Enum.Magenta, 1.0F))

            '    'm_pD2DHwndRenderTarget.Clear(New ColorF(ColorF.Enum.Red, 1.0F))

            '    'm_pD2DHwndRenderTarget.Clear(New ColorF(1.0F, 0F, 0.0F, 1.0F))

            '    'm_pD2DHwndRenderTarget.Clear(New ColorF(1.0F, 0F, 0.0F, 1.0F))

            '    ' Transparent but with Red halo
            '    m_pD2DHwndRenderTarget.Clear(New ColorF(1.0F, 0F, 0.0F, 1.0F))

            '    'm_pD2DHwndRenderTarget.Clear(New ColorF(0.0F, 0F, 1.0F, 1.0F))

            '    'm_pD2DHwndRenderTarget.Clear(New ColorF(0F, 0F, 0F, 1.0F))


            '    Dim nOldAntialiasMode As D2D1_ANTIALIAS_MODE
            '    m_pD2DHwndRenderTarget.GetAntialiasMode(nOldAntialiasMode)
            '    'm_pD2DHwndRenderTarget.SetAntialiasMode(D2D1_ANTIALIAS_MODE.D2D1_ANTIALIAS_MODE_ALIASED)
            '    m_pD2DHwndRenderTarget.SetAntialiasMode(D2D1_ANTIALIAS_MODE.D2D1_ANTIALIAS_MODE_PER_PRIMITIVE)

            '    Dim centerX As Single = Size.Width / 2.0F
            '    Dim centerY As Single = Size.Height / 2.0F

            '    ' Ellipse radii (adjust as you want)
            '    Dim radiusX As Single = 200.0F
            '    Dim radiusY As Single = 200.0F

            '    'm_pD2DHwndRenderTarget.FillEllipse(Ellipse(New Direct2D.D2D1_POINT_2F(centerX, centerY), radiusX, radiusY), m_pD2DMainBrush)
            '    If m_pD2DBitmap IsNot Nothing Then
            '        Dim sizeBitmap As D2D1_SIZE_F
            '        m_pD2DBitmap.GetSize(sizeBitmap)
            '        Dim destRectBitmap As D2D1_RECT_F = New D2D1_RECT_F(0.0F, 0.0F, Size.Width, Size.Height)
            '        'Dim destRectBitmap As D2D1_RECT_F = New D2D1_RECT_F(0.0F, 0.0F, sizeBitmap.width, sizeBitmap.height)
            '        Dim sourceRectBitmap As D2D1_RECT_F = New D2D1_RECT_F(0.0F, 0.0F, sizeBitmap.width, sizeBitmap.height)
            '        m_pD2DHwndRenderTarget.DrawBitmap(m_pD2DBitmap, destRectBitmap, 1.0F, D2D1_BITMAP_INTERPOLATION_MODE.D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, sourceRectBitmap)
            '    End If

            '    m_pD2DHwndRenderTarget.SetAntialiasMode(nOldAntialiasMode)

            '    Dim tag1 As ULong, tag2 As ULong = 0
            '    hr = m_pD2DHwndRenderTarget.EndDraw(tag1, tag2)
            'End If

            EndPaint(hWnd, ps)
        End If
        'InvalidateRect(Me.Handle, IntPtr.Zero, False)
        Return (hr)
    End Function

    Public Function OnResizeProc(ByRef m As Message) As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        If (m_pDXGISwapChain1 IsNot Nothing) Then

            If (m_pD2DDeviceContext IsNot Nothing) Then
                m_pD2DDeviceContext.SetTarget(Nothing)
            End If

            If (m_pD2DTargetBitmap IsNot Nothing) Then
                SafeRelease(m_pD2DTargetBitmap)
            End If

            '// 0, 0 => HRESULT 0x80070057 (E_INVALIDARG) if Not CreateSwapChainForHwnd
            'hr = m_pDXGISwapChain1.ResizeBuffers(
            ' 2,
            ' 0,
            ' 0,
            ' DXGI.DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM,
            ' 0
            ' )
            hr = m_pDXGISwapChain1.ResizeBuffers(
             2,
             LOWORD(m.LParam),
             HIWORD(m.LParam),
             DXGI.DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM,
             0
             )
            ConfigureSwapChain()
        End If
        Return hr
    End Function

    Public Sub SetOpacity(nOpacity As Integer)
        SetLayeredWindowAttributes(Me.Handle, 0, CByte(255 * nOpacity / 100), LWA_ALPHA)
    End Sub

    Public Sub MakeWindowTransparent(hWnd As IntPtr)
        Dim width As Integer = Me.Width
        Dim height As Integer = Me.Height
        Dim sizeBitmap As System.Drawing.Size = New System.Drawing.Size(width, height)

        Dim hdcScreen As IntPtr = GetDC(IntPtr.Zero)
        Dim hdcMem As IntPtr = CreateCompatibleDC(hdcScreen)

        ' Create a 32-bit transparent bitmap
        Dim bmi As New BITMAPINFO()
        bmi.bmiHeader.biSize = Marshal.SizeOf(GetType(BITMAPINFOHEADER))
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = -height ' top-down bitmap
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = BI_RGB

        Dim pBits As IntPtr = IntPtr.Zero
        Dim hBitmap As IntPtr = CreateDIBSection(hdcScreen, bmi, DIB_RGB_COLORS, pBits, IntPtr.Zero, 0)

        ' Fill with transparent pixels (A = 0)
        Dim totalBytes As Integer = width * height * 4
        Dim managed() As Byte = New Byte(totalBytes - 1) {}  ' already all zeros
        Marshal.Copy(managed, 0, pBits, totalBytes)

        Dim hOld As IntPtr = SelectObject(hdcMem, hBitmap)

        Dim rectWnd As RECT
        GetWindowRect(hWnd, rectWnd)
        Dim ptClient As Drawing.Point
        ptClient.X = rectWnd.left
        ptClient.Y = rectWnd.top
        ScreenToClient(GetParent(hWnd), ptClient)
        'rectWnd.left = Me.Left
        'rectWnd.top = Me.Top

        Dim ptSrc As System.Drawing.Point = New System.Drawing.Point()
        'Dim ptDest As System.Drawing.Point = New System.Drawing.Point(rectWnd.left, rectWnd.top)
        Dim ptDest As System.Drawing.Point = New System.Drawing.Point(ptClient.X, ptClient.Y)

        Dim pptSrc = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(System.Drawing.Point)))
        Marshal.StructureToPtr(ptSrc, pptSrc, False)

        Dim pptDest = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(System.Drawing.Point)))
        Marshal.StructureToPtr(ptDest, pptDest, False)

        Dim psizeBitmap = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(System.Drawing.Size)))
        Marshal.StructureToPtr(sizeBitmap, psizeBitmap, False)

        Dim blend As New BLENDFUNCTION With {
            .BlendOp = AC_SRC_OVER,
            .SourceConstantAlpha = 255,
            .AlphaFormat = AC_SRC_ALPHA
        }

        UpdateLayeredWindow(hWnd, hdcScreen, pptDest, psizeBitmap, hdcMem, pptSrc, 0, blend, ULW_ALPHA)

        SelectObject(hdcMem, hOld)
        DeleteObject(hBitmap)
        DeleteDC(hdcMem)
        ReleaseDC(IntPtr.Zero, hdcScreen)
    End Sub

    Public Function GetD2DBitmapFromStream(ms As IO.MemoryStream) As ID2D1Bitmap1
        Dim hr As HRESULT = HRESULT.S_OK
        Dim pBitmap As ID2D1Bitmap = Nothing
        Dim byteArray = ms.ToArray()
        Dim wicStream As IWICStream = Nothing
        hr = m_pWICImagingFactory.CreateStream(wicStream)
        If hr = HRESULT.S_OK Then
            hr = wicStream.InitializeFromMemory(byteArray, byteArray.Length)
            If hr = HRESULT.S_OK Then
                Dim pDecoder As IWICBitmapDecoder = Nothing
                hr = m_pWICImagingFactory.CreateDecoderFromStream(wicStream, Guid.Empty, WICDecodeOptions.WICDecodeMetadataCacheOnDemand, pDecoder)
                If hr = HRESULT.S_OK Then
                    Dim pFrame As IWICBitmapFrameDecode = Nothing
                    hr = pDecoder.GetFrame(0, pFrame)
                    If hr = HRESULT.S_OK Then
                        Dim pConvertedSourceBitmap As IWICFormatConverter = Nothing
                        hr = m_pWICImagingFactory.CreateFormatConverter(pConvertedSourceBitmap)
                        If hr = HRESULT.S_OK Then
                            hr = pConvertedSourceBitmap.Initialize(pFrame, WICTools.GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherType.WICBitmapDitherTypeNone, Nothing, 0, WICBitmapPaletteType.WICBitmapPaletteTypeCustom)
                            If hr = HRESULT.S_OK Then
                                Dim bp As D2D1_BITMAP_PROPERTIES = New D2D1_BITMAP_PROPERTIES()

                                'hr = m_pD2DHwndRenderTarget.CreateBitmapFromWicBitmap(pConvertedSourceBitmap, bp, pBitmap)
                                hr = m_pD2DDeviceContext.CreateBitmapFromWicBitmap(pConvertedSourceBitmap, bp, pBitmap)

                                If hr = HRESULT.S_OK Then
                                    Return pBitmap
                                End If
                            End If
                            SafeRelease(pConvertedSourceBitmap)
                        End If
                        SafeRelease(pFrame)
                    End If
                    SafeRelease(pDecoder)
                End If
            End If
            SafeRelease(wicStream)
        End If
        Return Nothing
    End Function

    Private Sub SetPictureToLayeredWindow(hWnd As IntPtr, hBitmap As IntPtr)
        Dim bm As BITMAP
        GetObject(hBitmap, Marshal.SizeOf(GetType(BITMAP)), bm)
        Dim sizeBitmap As System.Drawing.Size = New System.Drawing.Size(bm.bmWidth, bm.bmHeight)

        Dim hDCScreen As IntPtr = GetDC(IntPtr.Zero)
        Dim hDCMem As IntPtr = CreateCompatibleDC(hDCScreen)
        Dim hBitmapOld As IntPtr = SelectObject(hDCMem, hBitmap)

        Dim bf As BLENDFUNCTION = New BLENDFUNCTION()
        bf.BlendOp = AC_SRC_OVER
        bf.SourceConstantAlpha = 255
        bf.AlphaFormat = AC_SRC_ALPHA

        Dim rectWnd As RECT
        GetWindowRect(hWnd, rectWnd)
        Dim ptClient As Drawing.Point
        ptClient.X = rectWnd.left
        ptClient.Y = rectWnd.top
        ScreenToClient(GetParent(hWnd), ptClient)

        Dim ptSrc As System.Drawing.Point = New System.Drawing.Point()
        Dim ptDest As System.Drawing.Point = New System.Drawing.Point(ptClient.X, ptClient.Y)

        Dim pptSrc = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(System.Drawing.Point)))
        Marshal.StructureToPtr(ptSrc, pptSrc, False)

        Dim pptDest = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(System.Drawing.Point)))
        Marshal.StructureToPtr(ptDest, pptDest, False)

        Dim psizeBitmap = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(System.Drawing.Size)))
        Marshal.StructureToPtr(sizeBitmap, psizeBitmap, False)

        Dim bRet As Boolean = UpdateLayeredWindow(hWnd, hDCScreen, pptDest, psizeBitmap, hDCMem, pptSrc, 0, bf, ULW_ALPHA)

        Marshal.FreeHGlobal(pptSrc)
        Marshal.FreeHGlobal(pptDest)
        Marshal.FreeHGlobal(psizeBitmap)

        SelectObject(hDCMem, hBitmapOld)
        DeleteDC(hDCMem)
        ReleaseDC(IntPtr.Zero, hDCScreen)
    End Sub

    Public Sub UpdateLayeredFromD2DBitmap(hWnd As IntPtr, gpuBitmap As ID2D1Bitmap1,
                                                 d2dFactory As ID2D1Factory, wicFactory As IWICImagingFactory)

        Dim hr As HRESULT = HRESULT.S_OK

        Dim dxgiSurface As IDXGISurface = Nothing
        hr = gpuBitmap.GetSurface(dxgiSurface)

        'Dim pTex As IntPtr
        'Dim riid As Guid = GetType(DXGI.ID3D11Texture2D).GUID
        'Marshal.QueryInterface(Marshal.GetIUnknownForObject(dxgiSurface), riid, pTex)
        'Dim srcTexture As DXGI.ID3D11Texture2D = CType(Marshal.GetObjectForIUnknown(pTex), DXGI.ID3D11Texture2D)

        Dim srcTexture As DXGI.ID3D11Texture2D = dxgiSurface

        Dim desc As DXGI.D3D11_TEXTURE2D_DESC
        srcTexture.GetDesc(desc)

        Dim stagingDesc As DXGI.D3D11_TEXTURE2D_DESC
        stagingDesc.Width = desc.Width
        stagingDesc.Height = desc.Height
        stagingDesc.MipLevels = 1
        stagingDesc.ArraySize = 1
        stagingDesc.Format = desc.Format
        stagingDesc.SampleDesc = desc.SampleDesc
        stagingDesc.Usage = DXGI.D3D11_USAGE.D3D11_USAGE_STAGING
        stagingDesc.BindFlags = 0
        stagingDesc.CPUAccessFlags = D3D11_CPU_ACCESS_FLAG.D3D11_CPU_ACCESS_READ
        stagingDesc.MiscFlags = 0

        Dim stagingTexture As DXGI.ID3D11Texture2D = Nothing
        hr = m_pD3D11Device.CreateTexture2D(stagingDesc, IntPtr.Zero, stagingTexture)

        m_pD3D11DeviceContext.CopyResource(stagingTexture, srcTexture)

        ' --- 5. Map staging for CPU read ---
        Dim mapped As D3D11_MAPPED_SUBRESOURCE
        hr = m_pD3D11DeviceContext.Map(stagingTexture, 0, D3D11_MAP.D3D11_MAP_READ, 0, mapped)

        Dim width As Integer = desc.Width
        Dim height As Integer = desc.Height
        Dim rowPitch As Integer = mapped.RowPitch
        Dim totalBytes As Integer = rowPitch * height
        Dim pixels(totalBytes - 1) As Byte

        Dim ptr As IntPtr = mapped.pData
        For y As Integer = 0 To height - 1
            Marshal.Copy(IntPtr.Add(ptr, y * rowPitch), pixels, y * rowPitch, rowPitch)
        Next

        m_pD3D11DeviceContext.Unmap(stagingTexture, 0)

        ' 1. Get the bitmap size
        'Dim size As D2D1_SIZE_F
        'gpuBitmap.GetSize(size)
        'Dim width As Integer = CInt(size.width)
        'Dim height As Integer = CInt(size.height)

        '' 2. Create a WIC bitmap (CPU memory)
        'Dim wicBitmap As IWICBitmap = Nothing
        'wicFactory.CreateBitmap(
        'width,
        'height,
        'WICTools.GUID_WICPixelFormat32bppPBGRA,
        'WICBitmapCreateCacheOption.WICBitmapCacheOnLoad,
        'wicBitmap)

        'Dim rtp = RenderTargetProperties(D2D1_RENDER_TARGET_TYPE.D2D1_RENDER_TARGET_TYPE_DEFAULT,
        '                       PixelFormat(DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE.D2D1_ALPHA_MODE_PREMULTIPLIED))

        'Dim wicRenderTarget As ID2D1RenderTarget = Nothing
        'hr = d2dFactory.CreateWicBitmapRenderTarget(wicBitmap, rtp, wicRenderTarget)

        '' 4. Render the GPU bitmap into the WIC bitmap (CPU)
        'wicRenderTarget.BeginDraw()
        'Dim dest As New D2D1_RECT_F(0, 0, width, height)

        ''D2D DEBUG ERROR - An operation failed because a device-dependent resource is associated with the wrong ID2D1Device (resource domain).
        'Dim sourceRectBitmap As D2D1_RECT_F = New D2D1_RECT_F(0.0F, 0.0F, width, height)
        'wicRenderTarget.DrawBitmap(gpuBitmap, dest, 1.0F, D2D1_BITMAP_INTERPOLATION_MODE.D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, sourceRectBitmap)

        'Dim tag1 As ULong, tag2 As ULong = 0
        'hr = wicRenderTarget.EndDraw(tag1, tag2)

        ' 5. Lock WIC bitmap to get pixel pointer
        'Dim rc As New WICRect(0, 0, width, height)
        'Dim wicLock As IWICBitmapLock = Nothing
        'wicBitmap.Lock(rc, WICBitmapLockFlags.WICBitmapLockRead, wicLock)

        'Dim pixelPtr As IntPtr = IntPtr.Zero
        'Dim bufferSize As UInteger = 0
        'wicLock.GetDataPointer(bufferSize, pixelPtr)

        'Dim stride As UInteger = 0
        'wicLock.GetStride(stride)

        Dim bmi As New BITMAPINFO
        bmi.bmiHeader.biSize = Marshal.SizeOf(GetType(BITMAPINFOHEADER))
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = -height
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = BI_RGB

        Dim pBits As IntPtr
        Dim hScreenDC As IntPtr = GetDC(IntPtr.Zero)
        Dim hBitmap As IntPtr = CreateDIBSection(hScreenDC, bmi, DIB_RGB_COLORS, pBits, IntPtr.Zero, 0)

        If hBitmap = IntPtr.Zero Then Exit Sub

        ' --- 7. Copy pixel buffer into DIB ---
        Marshal.Copy(pixels, 0, pBits, pixels.Length)

        ' --- 8. Apply DIBSection to layered window ---
        Dim memDC As IntPtr = CreateCompatibleDC(IntPtr.Zero)
        Dim oldBmp As IntPtr = SelectObject(memDC, hBitmap)

        Dim blend As New BLENDFUNCTION With {
            .BlendOp = AC_SRC_OVER,
            .BlendFlags = 0,
            .SourceConstantAlpha = 255,
            .AlphaFormat = AC_SRC_ALPHA
        }

        Dim wndSize As New System.Drawing.Size(width, height)

        Dim rectWnd As RECT
        GetWindowRect(hWnd, rectWnd)
        Dim ptClient As Drawing.Point
        ptClient.X = rectWnd.left
        ptClient.Y = rectWnd.top
        ScreenToClient(GetParent(hWnd), ptClient)

        Dim ptSrc As System.Drawing.Point = New System.Drawing.Point()
        Dim ptDest As System.Drawing.Point = New System.Drawing.Point(ptClient.X, ptClient.Y)

        Dim pptSrc = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(System.Drawing.Point)))
        Marshal.StructureToPtr(ptSrc, pptSrc, False)

        Dim pptDest = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(System.Drawing.Point)))
        Marshal.StructureToPtr(ptDest, pptDest, False)

        Dim pSize = Marshal.AllocHGlobal(Marshal.SizeOf(wndSize))
        Marshal.StructureToPtr(wndSize, pSize, False)

        Dim bRet = UpdateLayeredWindow(hWnd, hScreenDC, pptDest, pSize, memDC, pptSrc, 0, blend, ULW_ALPHA)

        Marshal.FreeHGlobal(pptSrc)
        Marshal.FreeHGlobal(pptDest)
        Marshal.FreeHGlobal(pSize)

        SelectObject(memDC, oldBmp)
        DeleteObject(hBitmap)
        DeleteDC(memDC)
        ReleaseDC(IntPtr.Zero, hScreenDC)

    End Sub


    'Private Sub SetD2DBitmapToLayeredWindow(hWnd As IntPtr, pBitmap As ID2D1Bitmap1)
    '    If (pBitmap IsNot Nothing) Then
    '        Dim sizeBitmap As D2D1_SIZE_F
    '        pBitmap.GetSize(sizeBitmap)

    '        Dim hDCScreen As IntPtr = GetDC(IntPtr.Zero)
    '        Dim bmi As New BITMAPINFO()
    '        bmi.bmiHeader.biSize = Marshal.SizeOf(GetType(BITMAPINFOHEADER))
    '        bmi.bmiHeader.biWidth = sizeBitmap.width
    '        bmi.bmiHeader.biHeight = -sizeBitmap.height ' top-down bitmap
    '        bmi.bmiHeader.biPlanes = 1
    '        bmi.bmiHeader.biBitCount = 32
    '        bmi.bmiHeader.biCompression = BI_RGB

    '        Dim pBits As IntPtr = IntPtr.Zero
    '        Dim hBitmap As IntPtr = CreateDIBSection(hDCScreen, bmi, DIB_RGB_COLORS, pBits, IntPtr.Zero, 0)
    '        Dim hDCMem As IntPtr = CreateCompatibleDC(hDCScreen)
    '        Dim hBitmapOld As IntPtr = SelectObject(hDCMem, hBitmap)

    '        Dim rectClientWnd As RECT
    '        GetClientRect(hWnd, rectClientWnd)
    '        Dim rc As RECT = New RECT(0, 0, rectClientWnd.right, rectClientWnd.bottom)
    '        Dim hr As HRESULT = m_pD2DDCRenderTarget.BindDC(hDCMem, rc)
    '        If SUCCEEDED(hr) And m_pD2DDCRenderTarget IsNot Nothing Then
    '            m_pD2DDCRenderTarget.BeginDraw()
    '            If pBitmap IsNot Nothing Then
    '                m_pD2DDCRenderTarget.Clear(New ColorF(ColorF.Enum.Red, 1.0F))
    '                Dim destRectBitmap As D2D1_RECT_F = New D2D1_RECT_F(0.0F, 0.0F, sizeBitmap.width, sizeBitmap.height)
    '                Dim sourceRectBitmap As D2D1_RECT_F = New D2D1_RECT_F(0.0F, 0.0F, sizeBitmap.width, sizeBitmap.height)
    '                ' D2D DEBUG ERROR - An operation failed because a device-dependent resource is associated with the wrong ID2D1Device (resource domain).
    '                m_pD2DDCRenderTarget.DrawBitmap(pBitmap, destRectBitmap, 1.0F, D2D1_BITMAP_INTERPOLATION_MODE.D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, sourceRectBitmap)
    '            End If
    '            Dim tag1 As ULong, tag2 As ULong = 0
    '            hr = m_pD2DDCRenderTarget.EndDraw(tag1, tag2)

    '            Dim bf As BLENDFUNCTION = New BLENDFUNCTION()
    '            bf.BlendOp = AC_SRC_OVER
    '            bf.SourceConstantAlpha = 255
    '            bf.AlphaFormat = AC_SRC_ALPHA

    '            Dim rectWnd As RECT
    '            GetWindowRect(hWnd, rectWnd)
    '            Dim ptClient As Drawing.Point
    '            ptClient.X = rectWnd.left
    '            ptClient.Y = rectWnd.top
    '            ScreenToClient(GetParent(hWnd), ptClient)

    '            Dim ptSrc As System.Drawing.Point = New System.Drawing.Point()
    '            Dim ptDest As System.Drawing.Point = New System.Drawing.Point(ptClient.X, ptClient.Y)

    '            Dim pptSrc = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(System.Drawing.Point)))
    '            Marshal.StructureToPtr(ptSrc, pptSrc, False)

    '            Dim pptDest = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(System.Drawing.Point)))
    '            Marshal.StructureToPtr(ptDest, pptDest, False)

    '            Dim psizeBitmap = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(System.Drawing.Size)))
    '            Marshal.StructureToPtr(sizeBitmap, psizeBitmap, False)

    '            Dim bRet As Boolean = UpdateLayeredWindow(hWnd, hDCScreen, pptDest, psizeBitmap, hDCMem, pptSrc, 0, bf, ULW_ALPHA)

    '            Marshal.FreeHGlobal(pptSrc)
    '            Marshal.FreeHGlobal(pptDest)
    '            Marshal.FreeHGlobal(psizeBitmap)
    '        End If

    '        SelectObject(hDCMem, hBitmapOld)
    '        DeleteDC(hDCMem)
    '        ReleaseDC(IntPtr.Zero, hDCScreen)
    '    End If
    'End Sub

    Public Function LoadImage(sFile As String, color As Color) As IntPtr
        Dim hBitmap = Nothing
        Try
            Using img As Image = Image.FromFile(sFile)
                Using bmp As New System.Drawing.Bitmap(img)
                    hBitmap = bmp.GetHbitmap(color)
                End Using
            End Using
        Catch ex As Exception
            Dim nError As Integer = Marshal.GetLastWin32Error()
            Dim win32Exception As Win32Exception = New Win32Exception(nError)
            MessageBox.Show(String.Format("Error : {0}{1}{2}{1}({3})", nError.ToString(), Environment.NewLine, win32Exception.Message, ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return hBitmap
    End Function

    Public Sub SetBitmap(hBitmap As IntPtr)
        SetPictureToLayeredWindow(Me.Handle, hBitmap)
    End Sub

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function MoveWindow(hWnd As IntPtr, x As Integer, y As Integer, cx As Integer, cy As Integer, repaint As Boolean) As Boolean
    End Function

    Public Sub SetControlSizeToBitmap(pBitmap As ID2D1Bitmap)
        Dim sizeBitmap As D2D1_SIZE_F
        pBitmap.GetSize(sizeBitmap)
        'Me.Width = sizeBitmap.width
        'Me.Height = sizeBitmap.height
        MoveWindow(Me.Handle, Me.Left, Me.Top, sizeBitmap.width, sizeBitmap.height, False)
        Dim sz As Direct2D.D2D1_SIZE_U = New Direct2D.D2D1_SIZE_U(sizeBitmap.width, sizeBitmap.height)
        m_pD2DHwndRenderTarget.Resize(sz)
    End Sub

    Public Function CreateDeviceResources() As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        If m_pD2DMainBrush Is Nothing Then hr = m_pD2DDeviceContext.CreateSolidColorBrush(New ColorF(ColorF.Enum.Blue, 1.0F), Nothing, m_pD2DMainBrush)
        'If m_pD2DMainBrush Is Nothing Then hr = m_pD2DHwndRenderTarget.CreateSolidColorBrush(New ColorF(ColorF.Enum.Blue, 1.0F), Nothing, m_pD2DMainBrush)
        Return hr
    End Function

    Private Sub CleanDeviceResources()
        SafeRelease(m_pD2DMainBrush)
    End Sub

    Private Sub Clean()
        CleanDeviceResources()

        SafeRelease(m_pD2DHwndRenderTarget)

        SafeRelease(m_pD2DDeviceContext)
        SafeRelease(m_pD2DTargetBitmap)
        SafeRelease(m_pDXGISwapChain1)
        SafeRelease(m_pDXGIDevice)
        SafeRelease(m_pWICImagingFactory)
        SafeRelease(m_pDWriteFactory)
        SafeRelease(m_pD2DFactory1)

        SafeRelease(m_pDCompositionVisual)
        SafeRelease(m_pDCompositionTarget)
        SafeRelease(m_pDCompositionDevice)
    End Sub
End Class

