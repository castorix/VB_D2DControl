Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrayNotify
Imports Direct2D
Imports DWrite
Imports GlobalStructures
Imports GlobalStructures.GlobalTools
Imports WIC

Public Class Form1

    Dim m_pD2DFactory As ID2D1Factory = Nothing
    Dim m_pD2DFactory1 As ID2D1Factory1 = Nothing
    'Dim m_pDWriteFactory As IDWriteFactory = Nothing
    Dim m_pWICImagingFactory As IWICImagingFactory = Nothing

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hr As HRESULT = HRESULT.S_OK
        m_pWICImagingFactory = CType(Activator.CreateInstance(Type.GetTypeFromCLSID(WICTools.CLSID_WICImagingFactory)), IWICImagingFactory)
        hr = CreateD2D1Factory()
        If SUCCEEDED(hr) Then
            D2DControl1.Initialize(m_pD2DFactory1, m_pWICImagingFactory)
            D2DControl2.Initialize(m_pD2DFactory1, m_pWICImagingFactory)
            D2DControl3.Initialize(m_pD2DFactory1, m_pWICImagingFactory)
            D2DControl4.Initialize(m_pD2DFactory1, m_pWICImagingFactory)
        End If

        Dim img As Image = CType(My.Resources.ResourceManager.GetObject("Octopus_11x6"), Image)
        If (img IsNot Nothing) Then
            D2DControl1.SetBitmap(img, 11, 6)
            D2DControl1.SetControlSizeToBitmap()
        End If

        'Dim img2 As Image = CType(My.Resources.ResourceManager.GetObject("hulk"), Image)
        Dim img2 As Image = CType(My.Resources.ResourceManager.GetObject("Butterfly"), Image)
        If (img2 IsNot Nothing) Then
            D2DControl2.SetBitmap(img2, 1, 1, 0, Nothing, 0.5)
            'D2DControl2.SetBitmap(img2, 1, 1, 0, New ColorF(ColorF.Enum.Red, 1.0F), 0.5)
            'D2DControl2.SetBitmap(img2, 1, 1)
            D2DControl2.SetControlSizeToBitmap()
        End If

        'Dim img3 As Image = CType(My.Resources.ResourceManager.GetObject("Explode8x8"), Image)
        'If (img3 IsNot Nothing) Then
        '    D2DControl3.SetBitmap(img3, 8, 8, 63)
        '    D2DControl3.SetControlSizeToBitmap()
        '    D2DControl3.FrameDuration = 16
        'End If

        Dim img3 As Image = CType(My.Resources.ResourceManager.GetObject("Walking_girl_6x5"), Image)
        If (img3 IsNot Nothing) Then
            D2DControl3.SetBitmap(img3, 6, 5, 29)
            D2DControl3.SetControlSizeToBitmap()
            D2DControl3.FrameDuration = 40
        End If

        Dim img4 As Image = CType(My.Resources.ResourceManager.GetObject("Dragon8x8"), Image)
        If (img4 IsNot Nothing) Then
            D2DControl4.SetBitmap(img4, 8, 8, 62, Nothing, 0.5)
            D2DControl4.SetControlSizeToBitmap()
            D2DControl4.FrameDuration = 30
        End If

        PictureBox1.Image = My.Resources.ResourceManager.GetObject("Caribbean_Sea")
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

        Me.CenterToScreen()
    End Sub

    Private Function CreateD2D1Factory() As HRESULT
        Dim hr As HRESULT = HRESULT.S_OK
        Dim options As D2D1_FACTORY_OPTIONS = New D2D1_FACTORY_OPTIONS()
#If DEBUG Then
        options.debugLevel = D2D1_DEBUG_LEVEL.D2D1_DEBUG_LEVEL_INFORMATION
#End If
        hr = D2DTools.D2D1CreateFactory(D2D1_FACTORY_TYPE.D2D1_FACTORY_TYPE_SINGLE_THREADED, D2DTools.CLSID_D2D1Factory, options, m_pD2DFactory)
        m_pD2DFactory1 = CType(m_pD2DFactory, ID2D1Factory1)
        Return hr
    End Function

    'Public Function CreateDWriteFactory() As HRESULT
    '    Dim hr As HRESULT = HRESULT.S_OK
    '    hr = DWriteCreateFactory(D2DTools.DWRITE_FACTORY_TYPE.DWRITE_FACTORY_TYPE_SHARED, CLSID_DWriteFactory, m_pDWriteFactory)
    '    Return hr
    'End Function

    Private Sub cbOctopus_CheckedChanged(sender As Object, e As EventArgs) Handles cbOctopus.CheckedChanged
        D2DControl1.SetDraggable(cbOctopus.Checked)
    End Sub

    Private Sub cbButterfly_CheckedChanged(sender As Object, e As EventArgs) Handles cbButterfly.CheckedChanged
        D2DControl2.SetDraggable(cbButterfly.Checked)
    End Sub

    Private Sub cbGirl_CheckedChanged(sender As Object, e As EventArgs) Handles cbGirl.CheckedChanged
        D2DControl3.SetDraggable(cbGirl.Checked)
    End Sub

    Private Sub cbDragon_CheckedChanged(sender As Object, e As EventArgs) Handles cbDragon.CheckedChanged
        D2DControl4.SetDraggable(cbDragon.Checked)
    End Sub

    Private Sub Clean()
        'SafeRelease(m_pDWriteFactory)
        SafeRelease(m_pD2DFactory1)
        SafeRelease(m_pD2DFactory)
        SafeRelease(m_pWICImagingFactory)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Clean()
    End Sub
End Class


