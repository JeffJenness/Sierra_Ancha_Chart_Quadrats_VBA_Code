Attribute VB_Name = "Map_Module"
Option Explicit

Public Sub ExportMaps()

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

End Sub

Public Sub ConvertLayoutGraphics()

  Dim lngTransparency As Long
  Dim dblOutlineWidth As Double
  lngTransparency = 30
  dblOutlineWidth = 2

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument

  Dim pGeom As IGeometry
  Dim pArray As esriSystem.IArray
  Dim lngIndex As Long
  Dim pPolygon As IPolygon

  Dim pActiveView As IActiveView
  Dim pDisplay As IScreenDisplay
  Dim pDisplayTransform As IDisplayTransformation

  Dim pPoint As IPoint
  Dim pPtColl As IPointCollection
  Dim pNewPtColl As IPointCollection
  Dim pNewPolygon As IPolygon
  Dim lngX As Long
  Dim lngY As Long
  Dim lngIndex2 As Long
  Dim pNewPoint As IPoint
  Dim pSpRef As ISpatialReference
  Dim pMapView As IActiveView
  Set pMapView = pMxDoc.FocusMap
  Set pSpRef = pMxDoc.FocusMap.SpatialReference

  Dim pNewFClass As IFeatureClass
  Dim pWS As IWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Springs_Stewardship_Institute\Range_Maps\Map_Boxes.gdb", 0)
  Dim strName As String
  Dim pWS2 As IWorkspace2
  Set pWS2 = pWS
  Dim pEnv As IEnvelope
  Set pEnv = New Envelope
  Set pEnv.SpatialReference = pSpRef
  Dim pNewPolys As esriSystem.IArray
  Dim pNewBuff As IFeatureBuffer
  Dim pNewCursor As IFeatureCursor
  Dim lngIDIndex As Long
  Dim pNewFlayer As IFeatureLayer
  Dim pDataset As IDataset
  Dim pRender As ISimpleRenderer
  Dim pFillSymbol As ISimpleFillSymbol
  Dim pLineSymbol As ISimpleLineSymbol
  Dim pWhite As IRgbColor
  Dim pBlack As IRgbColor
  Dim pLyr As IGeoFeatureLayer
  Dim hx As IRendererPropertyPage
  Dim pLayerEffects As ILayerEffects
  Dim pNewFLayer2 As IFeatureLayer
  Dim pGroupLayer As IGroupLayer

  strName = MyGeneralOperations.MakeUniqueGDBFeatureClassName(pWS2, "Map_Boxes")

  Set pArray = MyGeneralOperations.ReturnGraphicsByNameFromLayout(pMxDoc, "Box", False)
  If pArray.Count > 0 Then
    Set pNewPolys = New esriSystem.Array
    For lngIndex = 0 To pArray.Count - 1
      Set pGeom = pArray.Element(lngIndex)

      If pGeom.GeometryType = esriGeometryPolygon Then
        Set pActiveView = pMxDoc.PageLayout
        Set pDisplay = pActiveView.ScreenDisplay
        Set pDisplayTransform = pDisplay.DisplayTransformation

        Set pPtColl = pGeom
        Set pNewPolygon = New Polygon
        Set pNewPolygon.SpatialReference = pSpRef
        Set pNewPtColl = pNewPolygon

        For lngIndex2 = 0 To pPtColl.PointCount - 1
          Set pPoint = pPtColl.Point(lngIndex2)
          pDisplayTransform.FromMapPoint pPoint, lngX, lngY
          Set pNewPoint = pMapView.ScreenDisplay.DisplayTransformation.ToMapPoint(lngX, lngY)
          Set pNewPoint.SpatialReference = pSpRef
          pNewPtColl.AddPoint pNewPoint
        Next lngIndex2
        pNewPolygon.Close
        pNewPolygon.SimplifyPreserveFromTo
        pNewPolys.Add pNewPolygon
        pEnv.Union pNewPolygon.Envelope
      End If
    Next lngIndex

    Set pNewFClass = MyGeneralOperations.CreateGDBFeatureClass(pWS, strName, esriFTSimple, pSpRef, esriGeometryPolygon, _
      , , , , True, ENUM_FileGDB, pEnv, pArray.Count)
    Set pNewBuff = pNewFClass.CreateFeatureBuffer
    Set pNewCursor = pNewFClass.Insert(True)
    lngIDIndex = pNewFClass.FindField("Unique_ID")

    For lngIndex = 0 To pNewPolys.Count - 1
      Set pNewPolygon = pNewPolys.Element(lngIndex)
      Set pNewBuff.Shape = pNewPolygon
      pNewCursor.InsertFeature pNewBuff
    Next lngIndex

    Set pWhite = MyGeneralOperations.MakeColorRGB(255, 255, 255)
    Set pBlack = MyGeneralOperations.MakeColorRGB(0, 0, 0)

    Set pNewFlayer = New FeatureLayer
    Set pNewFlayer.FeatureClass = pNewFClass
    Set pDataset = pNewFClass
    pNewFlayer.Name = pDataset.BrowseName & " Fill"
    Set pLyr = pNewFlayer
    Set pRender = New SimpleRenderer
    Set pLineSymbol = New SimpleLineSymbol
    Set pFillSymbol = New SimpleFillSymbol
    pLineSymbol.Width = 0
    pLineSymbol.Style = esriSLSNull
    pFillSymbol.Color = pWhite
    pFillSymbol.Outline = pLineSymbol
    pFillSymbol.Style = esriSFSSolid
    Set pRender.Symbol = pFillSymbol
    pRender.Label = "Fill"
    Set pLyr.Renderer = pRender
    Set hx = New SingleSymbolPropertyPage
    pLyr.RendererPropertyPageClassID = hx.ClassID
    Set pLayerEffects = pNewFlayer
    pLayerEffects.Transparency = lngTransparency

    Set pNewFLayer2 = New FeatureLayer
    Set pNewFLayer2.FeatureClass = pNewFClass
    Set pDataset = pNewFClass
    pNewFLayer2.Name = pDataset.BrowseName & " Outline"
    Set pLyr = pNewFLayer2
    Set pRender = New SimpleRenderer
    Set pLineSymbol = New SimpleLineSymbol
    Set pFillSymbol = New SimpleFillSymbol
    pLineSymbol.Width = 2
    pLineSymbol.Style = esriSLSSolid
    pLineSymbol.Color = pBlack
    pFillSymbol.Outline = pLineSymbol
    pFillSymbol.Style = esriSFSHollow
    Set pRender.Symbol = pFillSymbol
    pRender.Label = "Outline"
    Set pLyr.Renderer = pRender
    Set hx = New SingleSymbolPropertyPage
    pLyr.RendererPropertyPageClassID = hx.ClassID

    Set pGroupLayer = New GroupLayer
    pGroupLayer.Add pNewFlayer
    pGroupLayer.Add pNewFLayer2
    pGroupLayer.Name = pDataset.BrowseName
    pGroupLayer.Expanded = False

    pMxDoc.FocusMap.AddLayer pGroupLayer
    pMxDoc.UpdateContents
    pMxDoc.ActiveView.Refresh
  End If

ClearMemory:
  Set pMxDoc = Nothing
  Set pGeom = Nothing
  Set pArray = Nothing
  Set pPolygon = Nothing
  Set pActiveView = Nothing
  Set pDisplay = Nothing
  Set pDisplayTransform = Nothing
  Set pPoint = Nothing
  Set pPtColl = Nothing
  Set pNewPtColl = Nothing
  Set pNewPolygon = Nothing
  Set pSpRef = Nothing

End Sub

Public Sub ExportTreatments()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument

  Dim strSavePath As String

  Dim pFullNameElement As ITextElement
  Dim pCommonNameElement As ITextElement
  Dim pLocalityElement As ITextElement
  Dim pAddedLocalityElement As ITextElement
  Dim pMapNumberElement As ITextElement
  Dim pArray As esriSystem.IArray

  Set pArray = MyGeneralOperations.ReturnGraphicsByNameFromLayout(pMxDoc, "Full_Name", True)
  Set pFullNameElement = pArray.Element(0)
  Set pArray = MyGeneralOperations.ReturnGraphicsByNameFromLayout(pMxDoc, "Common_Name", True)
  Set pCommonNameElement = pArray.Element(0)
  Set pArray = MyGeneralOperations.ReturnGraphicsByNameFromLayout(pMxDoc, "Locality_Info", True)
  Set pLocalityElement = pArray.Element(0)
  Set pArray = MyGeneralOperations.ReturnGraphicsByNameFromLayout(pMxDoc, "Added_Locality_Info", True)
  Set pAddedLocalityElement = pArray.Element(0)
  Set pArray = MyGeneralOperations.ReturnGraphicsByNameFromLayout(pMxDoc, "Map Number", True)
  Set pMapNumberElement = pArray.Element(0)

  Dim strFullName As String
  Dim strCommonName As String
  Dim strLocality As String
  Dim strAddedLocality As String

  Dim pFLayer As IFeatureLayer
  Dim pFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngFullNameIndex As Long
  Dim lngCommonNameIndex As Long
  Dim lngLocalBLOBIndex As Long
  Dim lngSQLBLOBIndex As Long

  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Range_Maps", pMxDoc.FocusMap)
  Set pFClass = pFLayer.FeatureClass
  lngFullNameIndex = pFClass.FindField("Full_Name")
  lngCommonNameIndex = pFClass.FindField("Common_Name")
  lngLocalBLOBIndex = pFClass.FindField("Localities_BLOB")
  lngSQLBLOBIndex = pFClass.FindField("SQL_Queries_BLOB")

  Set pFCursor = pFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature

  Dim lngCounter As Long
  Dim lngCount As Long
  lngCount = pFClass.FeatureCount(Nothing)
  lngCounter = 0

  Dim strAddedLocal As String

  Dim varFullLocal As Variant
  Dim varFullAdded As Variant
  Dim varFullSQL As Variant

  Dim pMemoryBlobStream As IMemoryBlobStream
  Dim pMemVariant As IMemoryBlobStreamVariant

  Dim strFullExcelLocal As String
  Dim strFullExcelSQL As String

  Do Until pFeature Is Nothing
    strFullName = pFeature.Value(lngFullNameIndex)
    strCommonName = pFeature.Value(lngCommonNameIndex)

    Set pMemoryBlobStream = pFeature.Value(lngLocalBLOBIndex)
    Set pMemVariant = pMemoryBlobStream
    pMemVariant.ExportToVariant varFullLocal
    strLocality = aml_func_mod.BasicTrimAllCasesMultipleCharacters(ThisDocument.ReturnTextFromByteArray(varFullLocal), _
        "", " " & vbCrLf)

    Set pMemoryBlobStream = pFeature.Value(lngSQLBLOBIndex)
    Set pMemVariant = pMemoryBlobStream
    pMemVariant.ExportToVariant varFullSQL
    strFullExcelSQL = aml_func_mod.BasicTrimAllCasesMultipleCharacters(ThisDocument.ReturnTextFromByteArray(varFullSQL), _
        "", " " & vbCrLf)

    lngCounter = lngCounter + 1

    pFullNameElement.Text = strFullName
    pCommonNameElement.Text = strCommonName
    pLocalityElement.Text = "<BOL>Localites: </BOL> " & strLocality
    pAddedLocalityElement.Text = "<BOL>SQL Queries to identify Locality regions: </BOL> " & vbCrLf & _
        Replace(strFullExcelSQL, "\\", vbCrLf)
    pMapNumberElement.Text = "<BOL>" & CStr(lngCounter) & " of " & CStr(lngCount) & "</BOL>"

    Dim pLyr As IGeoFeatureLayer
    Set pLyr = pFLayer

    Dim pRender As IUniqueValueRenderer
    Dim n As Long
    Set pRender = New UniqueValueRenderer

    Dim pLineSymbol As ISimpleLineSymbol
    Set pLineSymbol = New SimpleLineSymbol
    pLineSymbol.Width = 1
    pLineSymbol.Color = MyGeneralOperations.MakeColorRGB(38, 115, 0)

    pRender.FieldCount = 1
    pRender.Field(0) = "Full_Name"
    pRender.DefaultSymbol = Nothing
    pRender.UseDefaultSymbol = False

    Dim pMainSymbol As ISimpleFillSymbol
    Set pMainSymbol = New SimpleFillSymbol
    pMainSymbol.Style = esriSFSSolid
    pMainSymbol.Outline = pLineSymbol
    pMainSymbol.Color = MyGeneralOperations.MakeColorRGB(0, 200, 0)

    pRender.AddValue strFullName, "Full_Name", pMainSymbol

    pRender.ColorScheme = "Custom"
    pRender.fieldType(0) = True
    Set pLyr.Renderer = pRender
    pLyr.DisplayField = "Full_Name"

    Dim hx As IRendererPropertyPage
    Set hx = New UniqueValuePropertyPage
    pLyr.RendererPropertyPageClassID = hx.ClassID

    pMxDoc.ActiveView.ContentsChanged
    pMxDoc.UpdateContents

    pMxDoc.ActiveView.Refresh

    Debug.Print "Exporting PDFs..."
    strSavePath = MyGeneralOperations.MakeUniquedBASEName( _
        "D:\arcGIS_stuff\consultation\Springs_Stewardship_Institute\Range_Maps\Sample_Maps\" & _
        Replace(strFullName, " ", "_") & ".pdf")

    Dim booExport As Boolean
    booExport = ExportActiveView(strSavePath)
    Set pFeature = pFCursor.NextFeature
  Loop

ClearMemory:
  Set pMxDoc = Nothing
  Set pFullNameElement = Nothing
  Set pCommonNameElement = Nothing
  Set pLocalityElement = Nothing
  Set pAddedLocalityElement = Nothing
  Set pArray = Nothing
  Set pFLayer = Nothing
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pLyr = Nothing
  Set pRender = Nothing
  Set pLineSymbol = Nothing
  Set pMainSymbol = Nothing
  Set hx = Nothing

  Debug.Print "Done..."
End Sub

Public Function ReturnExcelLocalities(strName As String, strOrigLocalities As String) As String

  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New ExcelWorkspaceFactory
  Dim pWS As IWorkspace
  Set pWS = pWSFact.OpenFromFile( _
      "D:\arcGIS_stuff\consultation\Springs_Stewardship_Institute\Range_Maps\Species_Lists_and_Literature\Invertebrates\" & _
      "Odonata\ODO_Master_112214.xlsx", 0)

  Dim pFeatWS As IFeatureWorkspace
  Set pFeatWS = pWS
  Dim pTable As ITable
  Set pTable = pFeatWS.OpenTable("ODO Import$")

  Dim lngExcelNameIndex As Long
  Dim lngExcelLocalitiesIndex As Long
  Dim lngOrigLocalitiesIndex As Long
  lngExcelNameIndex = pTable.FindField("FullName")
  lngOrigLocalitiesIndex = pTable.FindField("Localities")
  lngExcelLocalitiesIndex = pTable.FindField("Added by Jeff")
  Dim strText As String

  Dim pMemoryBlobStream As IMemoryBlobStream
  Dim pMemVariant As IMemoryBlobStreamVariant
  Dim varTest As Variant
  Dim varTestOrigLocal As Variant

  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim strTestName As String
  Set pCursor = pTable.Search(Nothing, False)
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing Or strTestName = strName
    strTestName = pRow.Value(lngExcelNameIndex)
    If strTestName = strName Then
      Set pMemoryBlobStream = pRow.Value(lngExcelLocalitiesIndex)
      Set pMemVariant = pMemoryBlobStream
      pMemVariant.ExportToVariant varTest
      Set pMemoryBlobStream = pRow.Value(lngOrigLocalitiesIndex)
      Set pMemVariant = pMemoryBlobStream
      pMemVariant.ExportToVariant varTestOrigLocal
    End If
    Set pRow = pCursor.NextRow
  Loop

  strText = CStr(varTest)
  If strText <> "" Then
    strText = Replace(strText, "//", vbCrLf & "     ")
  End If

  strOrigLocalities = CStr(varTestOrigLocal)

  ReturnExcelLocalities = strText

ClearMemory:
  Set pWSFact = Nothing
  Set pWS = Nothing
  Set pFeatWS = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing

End Function

Public Function ExportActiveView(strFilename As String) As Boolean

  Dim pApp As IApplication
  Dim pMxDoc As IMxDocument

  Set pApp = Application
  Set pMxDoc = ThisDocument

  ExportActiveView = False

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
  Dim sNameRoot As String
  Dim sOutputDir As String
  Dim iOutputResolution As Long
  Dim iScreenResolution As Long
  Dim bContinue As Boolean
  Dim msg As String
  Dim pTrackCancel As ITrackCancel
  Dim pGraphicsExtentEnv As IEnvelope
  Dim bClipToGraphicsExtent As Boolean
  Dim pUnitConvertor As IUnitConverter

  Dim pExportPNG As IExportPNG

  Set pActiveView = pMxDoc.ActiveView
  Set pTrackCancel = New CancelTracker

  Set pExport = New ExportPNG

  iOutputResolution = 400
  bClipToGraphicsExtent = False

  Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
  iPrevOutputImageQuality = pOutputRasterSettings.ResampleRatio
  If TypeOf pExport Is IExportImage Then
    SetOutputQuality pActiveView, 1
  ElseIf TypeOf pExport Is IOutputRasterSettings Then
    Set pOutputRasterSettings = pExport
    pOutputRasterSettings.ResampleRatio = 1
    Set pOutputRasterSettings = Nothing
  End If

  pExport.ExportFileName = strFilename
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
  Set pExport.StepProgressor = pApp.StatusBar.ProgressBar
  pTrackCancel.Reset
  pTrackCancel.CancelOnClick = False
  pTrackCancel.CancelOnKeyPress = True
  bContinue = pTrackCancel.Continue()

  hdc = pExport.StartExporting

  pActiveView.Output hdc, pExport.Resolution, exportRECT, pMapExtEnv, pTrackCancel

  bContinue = pTrackCancel.Continue()
  If bContinue Then
    msg = "Writing export file..."
    pApp.StatusBar.Message(0) = msg
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
    pApp.StatusBar.Message(0) = msg
  End If

  SetOutputQuality pActiveView, iPrevOutputImageQuality
  Set pTrackCancel = Nothing
  Set pMapExtEnv = Nothing
  Set pPixelBoundsEnv = Nothing

  ExportActiveView = True

End Function

Private Sub SetOutputQuality(pActiveView As IActiveView, iResampleRatio As Long)

  Dim pMap As IMap
  Dim pGraphicsContainer As IGraphicsContainer
  Dim pElement As IElement
  Dim pOutputRasterSettings As IOutputRasterSettings
  Dim pMapFrame As IMapFrame
  Dim pTmpActiveView As IActiveView

  If TypeOf pActiveView Is IMap Then
    Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
    pOutputRasterSettings.ResampleRatio = iResampleRatio
  ElseIf TypeOf pActiveView Is IPageLayout Then

    Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
    pOutputRasterSettings.ResampleRatio = iResampleRatio

    Set pGraphicsContainer = pActiveView
    pGraphicsContainer.Reset
    Set pElement = pGraphicsContainer.Next
    Do While Not pElement Is Nothing
      If TypeOf pElement Is IMapFrame Then
        Set pMapFrame = pElement
        Set pTmpActiveView = pMapFrame.Map
        Set pOutputRasterSettings = pTmpActiveView.ScreenDisplay.DisplayTransformation
        pOutputRasterSettings.ResampleRatio = iResampleRatio
      End If
      DoEvents
      Set pElement = pGraphicsContainer.Next
    Loop
    Set pMap = Nothing
    Set pMapFrame = Nothing
    Set pGraphicsContainer = Nothing
    Set pTmpActiveView = Nothing
  End If
  Set pOutputRasterSettings = Nothing

End Sub

Function GetGraphicsExtent(pActiveView As IActiveView) As IEnvelope

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

  Set GetGraphicsExtent = pBounds

  Set pBounds = Nothing
  Set pEnv = Nothing
  Set pGraphicsContainer = Nothing
  Set pPageLayout = Nothing
  Set pDisplay = Nothing
  Set pElement = Nothing

End Function

Public Function GetLegend_Point(pElement As IUnknown, _
                                strPos As String) _
                                As IPoint

    Dim theOutput As IPoint
    Set theOutput = Nothing
    Dim pEnv As IEnvelope

    If TypeOf pElement Is IElement Then
      Dim pTemp As IElement
      Set pTemp = pElement
      Set pEnv = pTemp.Geometry.Envelope
    Else
      Set pEnv = pElement
    End If

    Select Case UCase(strPos)
        Case "UL"
            Set theOutput = pEnv.UpperLeft
        Case "UR"
            Set theOutput = pEnv.UpperRight
        Case "LL"
            Set theOutput = pEnv.LowerLeft
        Case "LR"
            Set theOutput = pEnv.LowerRight
        Case "UC"
            Set theOutput = New Point
            theOutput.PutCoords ((pEnv.XMax - pEnv.XMin) / 2) + pEnv.XMin, pEnv.YMax
        Case Else
            MsgBox "position not supported: " & strPos
    End Select

    Set GetLegend_Point = theOutput

End Function ' GetLegend_Point

Public Sub Move_Legend(pElement As IElement, _
                       pFromPoint As IPoint, pToPoint As IPoint)

    Dim pTrans2D As ITransform2D
    Set pTrans2D = pElement

    pTrans2D.Move (pToPoint.x - pFromPoint.x), _
                  (pToPoint.Y - pFromPoint.Y)

End Sub ' Move_Legend

Sub CreateAndApplyUVRenderer(pLayer As IFeatureLayer, strFieldName As String, strLabel As String)

  Dim pApp As Application
  Dim pDoc As IMxDocument
  Set pDoc = ThisDocument
  Dim pMap As IMap
  Set pMap = pDoc.FocusMap

  Dim pFLayer As IFeatureLayer
  Set pFLayer = pLayer
  Dim pLyr As IGeoFeatureLayer
  Set pLyr = pFLayer

  Dim pFeatCls As IFeatureClass
  Set pFeatCls = pFLayer.FeatureClass
  Dim pQueryFilter As IQueryFilter
  Set pQueryFilter = New QueryFilter 'empty supports: SELECT *
  Dim pFeatCursor As IFeatureCursor
  Dim pFeature As IFeature

  Dim pValArray As esriSystem.IVariantArray
  Set pValArray = New varArray
  Dim lngVal As Long
  Dim lngFieldIndex As Long
  lngFieldIndex = pFeatCls.FindField(strFieldName)
  Dim pColl As Collection
  Set pColl = New Collection

  Set pFeatCursor = pFeatCls.Search(Nothing, True)
  Set pFeature = pFeatCursor.NextFeature
  Do Until pFeature Is Nothing
    lngVal = pFeature.Value(lngFieldIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pColl, CStr(lngVal)) Then
      pColl.Add True, CStr(lngVal)
      pValArray.Add lngVal
    End If
    Set pFeature = pFeatCursor.NextFeature
  Loop

  Dim lngArray() As Long
  ReDim lngArray(pValArray.Count - 1)
  Dim lngIndex As Long
  For lngIndex = 0 To pValArray.Count - 1
    lngArray(lngIndex) = pValArray.Element(lngIndex)
  Next lngIndex
  QuickSort.LongAscending lngArray, LBound(lngArray), UBound(lngArray)

  Dim booOK As Boolean
  Dim pColorRamp As IAlgorithmicColorRamp
  Set pColorRamp = New AlgorithmicColorRamp
  pColorRamp.Algorithm = esriHSVAlgorithm
  Dim pLightRed As IRgbColor
  Set pLightRed = New RgbColor
  pLightRed.RGB = RGB(255, 235, 214)
  Dim pDarkRed As IRgbColor
  Set pDarkRed = New RgbColor
  pDarkRed.RGB = RGB(196, 10, 10)
  pColorRamp.size = UBound(lngArray) + 1

  If pValArray.Count > 1 Then
    pColorRamp.FromColor = pLightRed
    pColorRamp.ToColor = pDarkRed
    pColorRamp.CreateRamp booOK
  End If

  Dim pRender As IUniqueValueRenderer, n As Long
  Set pRender = New UniqueValueRenderer

  Dim symd As ISimpleFillSymbol
  Set symd = New SimpleFillSymbol
  symd.Style = esriSFSSolid
  symd.Outline.Width = 0

  pRender.FieldCount = 1
  pRender.Field(0) = strFieldName
  pRender.DefaultSymbol = symd
  pRender.UseDefaultSymbol = False

  Dim pFill As ISimpleFillSymbol
  Dim pOutline As ISimpleLineSymbol

  For lngIndex = 0 To pValArray.Count - 1
    Set pFill = New SimpleFillSymbol
    pFill.Style = esriSFSSolid
    If pValArray.Count = 1 Then
      pFill.Color = pLightRed
    Else
      pFill.Color = pColorRamp.Color(lngIndex)
    End If
    Set pOutline = New SimpleLineSymbol
    pOutline.Width = 0
    If pValArray.Count = 1 Then
      pOutline.Color = pLightRed
    Else
      pOutline.Color = pColorRamp.Color(lngIndex)
    End If
    pOutline.Style = esriSLSNull
    pFill.Outline = pOutline
    pRender.AddValue lngArray(lngIndex), strFieldName, pFill
    pRender.Label(lngArray(lngIndex)) = CStr(lngArray(lngIndex)) & " " & strLabel
  Next lngIndex

  pRender.ColorScheme = "Custom"
  pRender.fieldType(0) = False
  Set pLyr.Renderer = pRender
  pLyr.DisplayField = strFieldName

  Dim hx As IRendererPropertyPage
  Set hx = New UniqueValuePropertyPage
  pLyr.RendererPropertyPageClassID = hx.ClassID

  pDoc.ActiveView.ContentsChanged
  pDoc.UpdateContents

  pDoc.ActiveView.Refresh

End Sub

Sub CreateAndApplyElevationRenderer(pRLayer As IRasterLayer)

  Dim pApp As Application
  Dim pDoc As IMxDocument
  Set pDoc = ThisDocument
  Dim pMap As IMap
  Set pMap = pDoc.FocusMap

  Dim pRaster As IRaster
  Set pRaster = pRLayer.Raster

  Dim pClassRen As IRasterClassifyColorRampRenderer
  Set pClassRen = New RasterClassifyColorRampRenderer
  Dim pRasRen As IRasterRenderer
  Set pRasRen = pClassRen

  Set pRasRen.Raster = pRaster
  pClassRen.ClassCount = 4

  pRasRen.Update

  Dim pFill As IFillSymbol

  Dim pColor1 As IRgbColor
  Set pColor1 = New RgbColor
  pColor1.RGB = RGB(255, 255, 191)
  Set pFill = New SimpleFillSymbol
  pFill.Color = pColor1
  pClassRen.Symbol(0) = pFill
  pClassRen.Break(0) = 0
  pClassRen.Label(0) = "< 5200' (< 1585m)"

  Dim pColor2 As IRgbColor
  Set pColor2 = New RgbColor
  pColor2.RGB = RGB(162, 254, 162)
  Set pFill = New SimpleFillSymbol
  pFill.Color = pColor2
  pClassRen.Symbol(1) = pFill
  pClassRen.Break(1) = 1585
  pClassRen.Label(1) = "5200' - 6700' (1585m - 2042m)"

  Dim pColor3 As IRgbColor
  Set pColor3 = New RgbColor
  pColor3.RGB = RGB(91, 143, 91)
  Set pFill = New SimpleFillSymbol
  pFill.Color = pColor3
  pClassRen.Symbol(2) = pFill
  pClassRen.Break(2) = 2042
  pClassRen.Label(2) = "6700' - 8200' (2042m - 2499m)"

  Dim pColor4 As IRgbColor
  Set pColor4 = New RgbColor
  pColor4.RGB = RGB(240, 252, 255)
  Set pFill = New SimpleFillSymbol
  pFill.Color = pColor4
  pClassRen.Symbol(3) = pFill
  pClassRen.Break(3) = 2499
  pClassRen.Label(3) = "> 8200' (> 2499m)"

  pRasRen.Update

  Set pRLayer.Renderer = pClassRen
  pDoc.ActiveView.ContentsChanged
  pDoc.UpdateContents

  pDoc.ActiveView.Refresh

End Sub

Sub CreateAndApplySlopeRenderer(pRLayer As IRasterLayer)

  Dim pApp As Application
  Dim pDoc As IMxDocument
  Set pDoc = ThisDocument
  Dim pMap As IMap
  Set pMap = pDoc.FocusMap

  Dim pRaster As IRaster
  Set pRaster = pRLayer.Raster

  Dim pClassRen As IRasterClassifyColorRampRenderer
  Set pClassRen = New RasterClassifyColorRampRenderer
  Dim pRasRen As IRasterRenderer
  Set pRasRen = pClassRen

  Dim pClassUIProps As IRasterClassifyUIProperties
  Set pClassUIProps = pClassRen
  Dim pClassify As IClassifyGEN
  Set pClassify = New DefinedInterval
  Set pClassUIProps.ClassificationMethod = pClassify.ClassID

  pRasRen.Update
  Set pRasRen.Raster = pRaster
  pClassRen.ClassCount = 5

  pRasRen.Update

  Dim booOK As Boolean
  Dim pColorRamp As IAlgorithmicColorRamp
  Set pColorRamp = New AlgorithmicColorRamp
  pColorRamp.Algorithm = esriHSVAlgorithm
  Dim pLightBlue As IRgbColor
  Set pLightBlue = New RgbColor
  pLightBlue.RGB = RGB(232, 252, 255)
  Dim pDarkBlue As IRgbColor
  Set pDarkBlue = New RgbColor
  pDarkBlue.RGB = RGB(35, 73, 105)
  pColorRamp.size = 5

  pColorRamp.FromColor = pLightBlue
  pColorRamp.ToColor = pDarkBlue

  pColorRamp.CreateRamp booOK

  Dim pFill As IFillSymbol

  Set pFill = New SimpleFillSymbol
  pFill.Color = pColorRamp.Color(0)
  pClassRen.Symbol(0) = pFill
  pClassRen.Break(0) = 0
  pClassRen.Label(0) = "< 20% (< 11�)"

  Set pFill = New SimpleFillSymbol
  pFill.Color = pColorRamp.Color(1)
  pClassRen.Symbol(1) = pFill
  pClassRen.Break(1) = 11.302232
  pClassRen.Label(1) = "20% - 40% (11� - 22�)"

  Set pFill = New SimpleFillSymbol
  pFill.Color = pColorRamp.Color(2)
  pClassRen.Symbol(2) = pFill
  pClassRen.Break(2) = 21.801409
  pClassRen.Label(2) = "40% - 100% (22� - 45�)"

  Set pFill = New SimpleFillSymbol
  pFill.Color = pColorRamp.Color(3)
  pClassRen.Symbol(3) = pFill
  pClassRen.Break(3) = 45
  pClassRen.Label(3) = "100% - 200% (45� - 63�)"

  Set pFill = New SimpleFillSymbol
  pFill.Color = pColorRamp.Color(4)
  pClassRen.Symbol(4) = pFill
  pClassRen.Break(4) = 63.434949
  pClassRen.Label(4) = "> 200% (> 63�)"

  pRasRen.Update

  Set pRLayer.Renderer = pClassRen
  pDoc.ActiveView.ContentsChanged
  pDoc.UpdateContents

  pDoc.ActiveView.Refresh

End Sub

Public Sub TestSlopeRenderer()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pRLayer As IRasterLayer
  Set pRLayer = MyGeneralOperations.ReturnLayerByName("Slope", pMxDoc.FocusMap)
  CreateAndApplySlopeRenderer pRLayer

End Sub

Public Sub TestElevRenderer()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pRLayer As IRasterLayer
  Set pRLayer = MyGeneralOperations.ReturnLayerByName("Elevation", pMxDoc.FocusMap)
  CreateAndApplyElevationRenderer pRLayer

End Sub

Public Sub TestRenderer()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFLayer As IFeatureLayer
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Historical Fires", pMxDoc.FocusMap)
  Dim strFieldName As String
  strFieldName = "AllLT_2010"
  Dim strLabel As String
  strLabel = "Treatments"

  CreateAndApplyUVRenderer pFLayer, strFieldName, strLabel

End Sub


