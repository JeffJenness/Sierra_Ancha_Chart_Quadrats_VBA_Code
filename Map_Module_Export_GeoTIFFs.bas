Attribute VB_Name = "Map_Module_Export_GeoTIFFs"
Option Explicit

Public Sub ExportActiveView2(strNameRoot As String, strOutputDir As String)

    Dim pMxDoc As IMxDocument
    Dim pActiveView As IActiveView
    Dim pExport As IExport
    Dim iPrevOutputImageQuality As Long
    Dim pOutputRasterSettings As IOutputRasterSettings
    Dim pPixelBoundsEnv As IEnvelope
    Dim exportRECT As tagRECT
    Dim DisplayBounds As tagRECT
    Dim pDisplayTransformation As IDisplayTransformation
    Dim pPageLayout As IPageLayout
    Dim pMapExtEnv As IEnvelope
    Dim hdc As Long
    Dim tmpDC As Long
    Dim iOutputResolution As Long
    Dim iScreenResolution As Long
    Dim bContinue As Boolean
    Dim msg As String
    Dim pTrackCancel As ITrackCancel
    Dim pGraphicsExtentEnv As IEnvelope
    Dim bClipToGraphicsExtent As Boolean
    Dim pUnitConvertor As IUnitConverter

    Set pMxDoc = Application.Document
    Set pActiveView = pMxDoc.ActiveView
    Set pTrackCancel = New CancelTracker

    Set pExport = New ExportTIFF
    Dim pExportTiff As IExportTIFF
    Dim pWorldFileSettings As IWorldFileSettings
    Set pExportTiff = pExport
    Set pWorldFileSettings = pExport
    If TypeOf pActiveView Is IMap Then
      pExportTiff.GeoTiff = True
      pWorldFileSettings.MapExtent = pActiveView.Extent
    End If
    Set pExportTiff = Nothing
    Set pWorldFileSettings = Nothing
    iOutputResolution = 200
    bClipToGraphicsExtent = True

    Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
    iPrevOutputImageQuality = pOutputRasterSettings.ResampleRatio
    If TypeOf pExport Is IExportImage Then
        SetOutputQuality2 pActiveView, 1
    ElseIf TypeOf pExport Is IOutputRasterSettings Then
        Set pOutputRasterSettings = pExport
        pOutputRasterSettings.ResampleRatio = 1
        Set pOutputRasterSettings = Nothing
    End If

    If Right(strOutputDir, 1) <> "\" Then strOutputDir = strOutputDir & "\"

    pExport.ExportFileName = strOutputDir & strNameRoot & "." & Right(Split(pExport.Filter, "|")(1), _
                             Len(Split(pExport.Filter, "|")(1)) - 2)
    tmpDC = GetDC(0)
    iScreenResolution = GetDeviceCaps(tmpDC, 88) '88 is the win32 const for Logical pixels/inch in X)
    ReleaseDC 0, tmpDC
    pExport.Resolution = iOutputResolution

    If TypeOf pActiveView Is IPageLayout Then
        DisplayBounds = pActiveView.ExportFrame
        Set pMapExtEnv = pGraphicsExtentEnv
    Else
        Set pDisplayTransformation = pActiveView.ScreenDisplay.DisplayTransformation
        DisplayBounds.Left = 0
        DisplayBounds.Top = 0
        DisplayBounds.Right = pDisplayTransformation.DeviceFrame.Right
        DisplayBounds.bottom = pDisplayTransformation.DeviceFrame.bottom
        Set pMapExtEnv = New Envelope
        Set pMapExtEnv = pDisplayTransformation.FittedBounds
    End If

    Set pPixelBoundsEnv = New Envelope
    If bClipToGraphicsExtent And (TypeOf pActiveView Is IPageLayout) Then
        Set pGraphicsExtentEnv = GetGraphicsExtent(pActiveView)
        Set pPageLayout = pActiveView
        Set pUnitConvertor = New UnitConverter
        pPixelBoundsEnv.XMin = 0
        pPixelBoundsEnv.YMin = 0
        pPixelBoundsEnv.XMax = pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.XMax, pPageLayout.Page.Units, esriInches) * pExport.Resolution _
                               - pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.XMin, pPageLayout.Page.Units, esriInches) * pExport.Resolution
        pPixelBoundsEnv.YMax = pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.YMax, pPageLayout.Page.Units, esriInches) * pExport.Resolution _
                               - pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.YMin, pPageLayout.Page.Units, esriInches) * pExport.Resolution

        With exportRECT
            .bottom = Fix(pPixelBoundsEnv.YMax) + 1
            .Left = Fix(pPixelBoundsEnv.XMin)
            .Top = Fix(pPixelBoundsEnv.YMin)
            .Right = Fix(pPixelBoundsEnv.XMax) + 1
        End With

        Set pMapExtEnv = pGraphicsExtentEnv
    Else
        With exportRECT
            .bottom = DisplayBounds.bottom * (iOutputResolution / iScreenResolution)
            .Left = DisplayBounds.Left * (iOutputResolution / iScreenResolution)
            .Top = DisplayBounds.Top * (iOutputResolution / iScreenResolution)
            .Right = DisplayBounds.Right * (iOutputResolution / iScreenResolution)
        End With
        pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
    End If

    pExport.PixelBounds = pPixelBoundsEnv

    Set pExport.TrackCancel = pTrackCancel
    Set pExport.StepProgressor = Application.StatusBar.ProgressBar
    pTrackCancel.Reset
    pTrackCancel.CancelOnClick = False
    pTrackCancel.CancelOnKeyPress = True
    bContinue = pTrackCancel.Continue()

    hdc = pExport.StartExporting

    pActiveView.Output hdc, pExport.Resolution, exportRECT, pMapExtEnv, pTrackCancel

    bContinue = pTrackCancel.Continue()
    If bContinue Then
        msg = "Writing export file..."
        Application.StatusBar.Message(0) = msg
        pExport.FinishExporting
        pExport.Cleanup
    Else
        pExport.Cleanup
    End If
    pTrackCancel.CancelOnClick = False
    pTrackCancel.CancelOnKeyPress = True

    bContinue = pTrackCancel.Continue()
    If bContinue Then
        msg = "Finished exporting '" & pExport.ExportFileName & "'"
        Application.StatusBar.Message(0) = msg
    End If

    SetOutputQuality2 pActiveView, iPrevOutputImageQuality
    Set pTrackCancel = Nothing
    Set pMapExtEnv = Nothing
    Set pPixelBoundsEnv = Nothing
End Sub

Function GetGraphicsExtent2(pActiveView As IActiveView) As IEnvelope
    Dim pBounds As IEnvelope
    Dim pEnv As IEnvelope
    Dim pGraphicsContainer As IGraphicsContainer
    Dim pPageLayout As IPageLayout
    Dim pDisplay As IDisplay
    Dim pElement As IElement

    Set pBounds = New Envelope
    Set pEnv = New Envelope
    Set pPageLayout = pActiveView
    Set pDisplay = pActiveView.ScreenDisplay
    Set pGraphicsContainer = pActiveView
    pGraphicsContainer.Reset

    Set pElement = pGraphicsContainer.Next
    Do While Not pElement Is Nothing
        pElement.QueryBounds pDisplay, pEnv
        pBounds.Union pEnv
        DoEvents
        Set pElement = pGraphicsContainer.Next
    Loop

    Set GetGraphicsExtent2 = pBounds

    Set pBounds = Nothing
    Set pEnv = Nothing
    Set pGraphicsContainer = Nothing
    Set pPageLayout = Nothing
    Set pDisplay = Nothing
    Set pElement = Nothing

End Function


