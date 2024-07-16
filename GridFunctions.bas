Attribute VB_Name = "GridFunctions"
Option Explicit

Public Enum enumRasterType
   enum_Grid_Type
   enum_Imagine_Type
   enum_TIFF_Type
   enum_JPEG_Type
   enum_JP2000_Type
   enum_BMP_Type
   enum_PNG_Type
   enum_GIF_Type
   enum_PCI_Raster_Type
   enum_X11_Pixmap_Type
   enum_PCRaster_Type
   enum_Memory_Raster_Type
   enum_HDF4_Type
   enum_BIL_Type
   enum_BIP_Type
   enum_BSQ_Type
   enum_IDRISI_Type
   enum_Geodatabase_Type
End Enum

Public Function OpenRasterWorkspace(sPath As String) As IRasterWorkspace

  Dim pWKSF As IWorkspaceFactory
  Set pWKSF = New RasterWorkspaceFactory

  Dim pRasterWs As IRasterWorkspace
  Set pRasterWs = pWKSF.OpenFromFile(sPath, 0)
  Set OpenRasterWorkspace = pRasterWs

  Set pWKSF = Nothing
  Set pRasterWs = Nothing

End Function

Public Sub SaveRasterAs(pRasterBandCol As IRasterBandCollection, strPath As String, strName As String, aRasterType As enumRasterType)

  Dim pSaveAs As IRasterBandCollection
  Set pSaveAs = pRasterBandCol
  Dim pRasterWs As IRasterWorkspace
  Set pRasterWs = OpenRasterWorkspace(strPath)
  If aRasterType = enum_Grid_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "GRID"
  ElseIf aRasterType = enum_Imagine_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "IMAGINE Image"
  ElseIf aRasterType = enum_TIFF_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "TIFF"
  ElseIf aRasterType = enum_JPEG_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "JPG"
  ElseIf aRasterType = enum_JP2000_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "JP2"
  ElseIf aRasterType = enum_BMP_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "BMP"
  ElseIf aRasterType = enum_PNG_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "PNG"
  ElseIf aRasterType = enum_GIF_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "GIF"
  ElseIf aRasterType = enum_PCI_Raster_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "PIX"
  ElseIf aRasterType = enum_X11_Pixmap_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "XPM"
  ElseIf aRasterType = enum_PCRaster_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "MAP"
  ElseIf aRasterType = enum_Memory_Raster_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "MEM"
  ElseIf aRasterType = enum_HDF4_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "HDF4"
  ElseIf aRasterType = enum_BIL_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "BIL"
  ElseIf aRasterType = enum_BIP_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "BIP"
  ElseIf aRasterType = enum_BSQ_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "BSQ"
  ElseIf aRasterType = enum_IDRISI_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "GDB"
  ElseIf aRasterType = enum_Geodatabase_Type Then
    pSaveAs.SaveAs strName, pRasterWs, "GDB"
  End If

  Set pSaveAs = Nothing
  Set pRasterWs = Nothing

End Sub

Public Function ReturnCellSize(pRaster As IRaster) As Double
  On Error GoTo erh
  Dim pRasLayer As IRasterLayer
  Set pRasLayer = New RasterLayer
  pRasLayer.CreateFromRaster pRaster

  Dim pRasterProps As IRasterProps
  Set pRasterProps = pRaster

  Dim lngNumRows As Long
  lngNumRows = pRasLayer.RowCount

  Dim pEnvelope As IEnvelope
  Set pEnvelope = pRasterProps.Extent

  ReturnCellSize = pEnvelope.Height / lngNumRows

  Set pRasLayer = Nothing
  Set pRasterProps = Nothing
  Set pEnvelope = Nothing

  Exit Function
erh:
    MsgBox "Failed in ReturnCellSize: " & err.Description

End Function

Public Function ReturnPixelHeight(pRaster As IRaster) As Double
  Dim pRasLayer As IRasterLayer
  Set pRasLayer = New RasterLayer
  pRasLayer.CreateFromRaster pRaster

  Dim pRasterProps As IRasterProps
  Set pRasterProps = pRaster

  Dim lngNumRows As Long
  lngNumRows = pRasLayer.RowCount

  Dim pEnvelope As IEnvelope
  Set pEnvelope = pRasterProps.Extent

  ReturnPixelHeight = pEnvelope.Height / lngNumRows

  Set pRasLayer = Nothing
  Set pRasterProps = Nothing
  Set pEnvelope = Nothing

  Exit Function

End Function

Public Function CalcContinuousGridStats(pRaster As IRaster, pRasterStats As IRasterStatistics, _
      lngNumBins As Long) As esriSystem.IVariantArray

  Dim pRasterAnalysisProps As IRasterAnalysisProps
  Dim pRasterProps As IRasterProps
  Set pRasterAnalysisProps = pRaster
  Set pRasterProps = pRaster

  Dim dblMaximum As Double
  Dim dblMinimum As Double
  Dim dblMean As Double
  Dim dblMedian As Double
  Dim dblMode As Double
  Dim dblStDev As Double

  Set pRasterAnalysisProps = Nothing
  Set pRasterProps = Nothing

End Function

Public Function ReturnPixelWidth(pRaster As IRaster) As Double
  Dim pRasLayer As IRasterLayer
  Set pRasLayer = New RasterLayer
  pRasLayer.CreateFromRaster pRaster

  Dim pRasterProps As IRasterProps
  Set pRasterProps = pRaster

  Dim lngNumCols As Long
  lngNumCols = pRasLayer.ColumnCount

  Dim pEnvelope As IEnvelope
  Set pEnvelope = pRasterProps.Extent

  ReturnPixelWidth = pEnvelope.Width / lngNumCols

  Set pRasLayer = Nothing
  Set pRasterProps = Nothing
  Set pEnvelope = Nothing

  Exit Function

End Function

Public Function ReturnPointsByCellSize(pRaster As IRaster, ByVal pLine As IGeometry) As IPointCollection
  On Error GoTo erh

  Dim pCurve As ICurve
  Dim dblLength As Double

  If TypeOf pLine Is ICurve Then
    Set pCurve = pLine
    dblLength = pCurve.length
  Else
    MsgBox "Invalid geometry type!  Must implement 'ICurve'..."
    Set ReturnPointsByCellSize = Nothing
    Exit Function
  End If
  Dim pPointCollection As IPointCollection

  Dim pSrcSpRef As ISpatialReference
  Dim pRasProps As IRasterProps
  Set pRasProps = pRaster
  Set pSrcSpRef = pRasProps.SpatialReference

  Dim pTrgSpRef As ISpatialReference
  Set pTrgSpRef = pLine.SpatialReference

  If Not GridFunctions.CompareSpatialReferences(pSrcSpRef, pTrgSpRef) Then
    pLine.Project pSrcSpRef
  End If

  Dim dblCellSize As Double
  dblCellSize = GridFunctions.ReturnCellSize(pRaster)
  Dim NumPoints As Long
  NumPoints = Int(dblLength / dblCellSize) + 1

  Dim pMpt As IMultipoint
  Dim pSegCol As ISegmentCollection
  Set pSegCol = pLine
  Set pMpt = GridFunctions.EllipticArcToPolygon2(pSegCol, NumPoints)
  Set pPointCollection = pMpt

  Set ReturnPointsByCellSize = pPointCollection

  Exit Function

  Set pCurve = Nothing
  Set pPointCollection = Nothing
  Set pSrcSpRef = Nothing
  Set pRasProps = Nothing
  Set pTrgSpRef = Nothing
  Set pMpt = Nothing
  Set pSegCol = Nothing

erh:
  If (Erl <> 0) Then
    MsgBox "Failed in ReturnPointsByCellSize: " & err.Description & vbCrLf & "Error at line " & CStr(Erl)
  Else
    MsgBox "Failed in ReturnPointsByCellSize: " & err.Description & vbCrLf & "No Line Number Available..."
  End If

End Function

Public Function CompareSpatialReferences(ByVal pSourceSR As ISpatialReference, ByVal pTargetSR As ISpatialReference) As Boolean

  On Error GoTo erh

  If pSourceSR Is Nothing And pTargetSR Is Nothing Then
    CompareSpatialReferences = True
    Exit Function
  ElseIf pSourceSR Is Nothing Or pTargetSR Is Nothing Then
    CompareSpatialReferences = False
    Exit Function
  End If

  Dim pSourceClone As IClone
  Dim pTargetClone As IClone
  Dim bSREqual As Boolean

  Set pSourceClone = pSourceSR
  Set pTargetClone = pTargetSR

  bSREqual = pSourceClone.IsEqual(pTargetClone)

  If Not bSREqual Then
    CompareSpatialReferences = False
    Exit Function
  End If

  Dim pSourceSR2 As ISpatialReference2
  Dim bXYIsEqual As Boolean

  Set pSourceSR2 = pSourceSR
  bXYIsEqual = pSourceSR2.IsXYPrecisionEqual(pTargetSR)

  If Not bXYIsEqual Then
    CompareSpatialReferences = False
    Exit Function
  End If

  CompareSpatialReferences = True

  Set pSourceClone = Nothing
  Set pTargetClone = Nothing
  Set pSourceSR2 = Nothing

  Exit Function
erh:
    MsgBox "Failed in CompareSpatialReferences: " & err.Description
End Function

Sub SetSpatialAnalysisSettings(TargetEnv As IRasterAnalysisEnvironment, _
                               SourceEnv As IRasterAnalysisEnvironment)
    On Error GoTo erh
    If Not SourceEnv Is Nothing Then
        Set TargetEnv.OutWorkspace = SourceEnv.OutWorkspace
        If Not SourceEnv.OutSpatialReference Is Nothing Then
            Set TargetEnv.OutSpatialReference = SourceEnv.OutSpatialReference
        End If
        TargetEnv.DefaultOutputRasterPrefix = SourceEnv.DefaultOutputRasterPrefix
        TargetEnv.DefaultOutputVectorPrefix = SourceEnv.DefaultOutputVectorPrefix
        If Not SourceEnv.Mask Is Nothing Then
            Set TargetEnv.Mask = SourceEnv.Mask
        End If
        Dim nCellSize As Double
        SourceEnv.GetCellSize 3, nCellSize
        If nCellSize <> 0 Then
            TargetEnv.SetCellSize 3, nCellSize
        End If
        Dim pExtent As IEnvelope
        SourceEnv.GetExtent 3, pExtent
        If Not pExtent Is Nothing Then
            TargetEnv.SetExtent 3, pExtent
        End If
        TargetEnv.VerifyType = SourceEnv.VerifyType
    End If
    Exit Sub

    Set pExtent = Nothing

erh:
    MsgBox "Failed in SetSpatialAnalysisSettings: " & err.Description
End Sub

Public Function ClipRasterToPolygon(pRaster As IRaster, ByVal pPolygon As IPolygon, SaveInside As Boolean, _
      Optional ByVal pClipEnvelope As IEnvelope, Optional CellSize As Double, _
      Optional ByVal pEnv As IRasterAnalysisEnvironment, Optional booShowProgress As Boolean, _
      Optional pApp As IApplication) As IRaster

  If booShowProgress Then
    Dim pSBar As IStatusBar
    Set pSBar = pApp.StatusBar
    pSBar.ProgressBar.position = 1
    Dim pPro As IStepProgressor
    Set pPro = pSBar.ProgressBar
  End If

  Dim booWorkingWithCurves As Boolean
  Dim pSegmentCollectionCurves As ISegmentCollection
  Set pSegmentCollectionCurves = pPolygon
  Dim pSegmentCurve As ISegment
  Set pSegmentCurve = pSegmentCollectionCurves.Segment(0)
  Dim pGeometryTypeA As esriGeometryType
  pGeometryTypeA = pSegmentCurve.GeometryType

  booWorkingWithCurves = (pGeometryTypeA = esriGeometryBezier3Curve) Or _
    (pGeometryTypeA = esriGeometryCircularArc) Or _
    (pGeometryTypeA = esriGeometryEllipticArc)

  If booWorkingWithCurves Then
    Dim pNewPoints As IPointCollection
    Set pNewPoints = GridFunctions.EllipticArcToPolygon2(pSegmentCollectionCurves, 100)
    Dim pNewPolygon As IPointCollection
    Set pNewPolygon = New Polygon
    pNewPolygon.SetPointCollection pNewPoints
    Set pPolygon = pNewPolygon
  End If

  Dim pExtractionOp As IExtractionOp
  Set pExtractionOp = New RasterExtractionOp

  Dim pRastAnalysisEnv As IRasterAnalysisEnvironment
  Dim pSpatialReference As ISpatialReference
  Dim pPolySpatRef As ISpatialReference
  Set pPolySpatRef = pPolygon.SpatialReference
  Dim SpatRefSame As Boolean
  Dim pEnvelope As IEnvelope

  If Not pEnv Is Nothing Then
    Set pRastAnalysisEnv = pEnv
    Set pSpatialReference = pRastAnalysisEnv.OutSpatialReference

    SpatRefSame = CompareSpatialReferences(pSpatialReference, pPolySpatRef)
    If Not SpatRefSame Then pPolygon.Project pSpatialReference
    pEnv.GetExtent esriRasterEnvValue, pEnvelope

  Else

    Set pRastAnalysisEnv = pExtractionOp
    pRastAnalysisEnv.RestoreToPreviousDefaultEnvironment

    Dim theEnvType As esriRasterEnvSettingEnum
    Dim theExtEnvType As esriRasterEnvSettingEnum
    Dim pTempEnv As IEnvelope
    Dim theCellSize As Double
    pRastAnalysisEnv.GetCellSize theEnvType, theCellSize
    pRastAnalysisEnv.GetExtent theExtEnvType, pTempEnv
    Dim pRasterAnalysisProps As IRasterAnalysisProps
    Dim pRasterProps As IRasterProps
        Set pRasterAnalysisProps = pRaster
        Set pRasterProps = pRaster

    Set pSpatialReference = pRasterProps.SpatialReference

    SpatRefSame = CompareSpatialReferences(pSpatialReference, pPolySpatRef)
    If Not SpatRefSame Then pPolygon.Project pSpatialReference

    If CellSize <= 0 Then
      theCellSize = pRasterAnalysisProps.PixelHeight
      theEnvType = pRastAnalysisEnv.VerifyType
    Else
      theCellSize = CellSize
      theEnvType = esriRasterEnvValue
    End If

    Dim pTopoOp As ITopologicalOperator2
    If pClipEnvelope Is Nothing Then
      Dim pPolyEnvelope As IEnvelope
          Set pPolyEnvelope = pPolygon.Envelope
      Dim pRastEnvelope As IEnvelope
          Set pRastEnvelope = pRasterProps.Extent
          Set pTopoOp = pPolyEnvelope
          pTopoOp.IsKnownSimple = False
          pTopoOp.Simplify

      Set pEnvelope = pTopoOp.Intersect(pPolyEnvelope, esriGeometry2Dimension)
    Else
      SpatRefSame = CompareSpatialReferences(pSpatialReference, pClipEnvelope.SpatialReference)
      If Not SpatRefSame Then
        pClipEnvelope.Project pSpatialReference
      End If
      Set pEnvelope = pClipEnvelope.Envelope
    End If

    pRastAnalysisEnv.SetCellSize theEnvType, theCellSize
    pRastAnalysisEnv.SetExtent esriRasterEnvValue, pEnvelope
    Set pRastAnalysisEnv.OutSpatialReference = pSpatialReference
  End If

  DoEvents

  Dim pIntPolygon As IPolygon4
  Set pTopoOp = pPolygon
  pTopoOp.IsKnownSimple = False
  pTopoOp.Simplify
  Set pIntPolygon = pTopoOp.Intersect(pEnvelope, esriGeometry2Dimension)
  Set pIntPolygon = pPolygon
  Set pIntPolygon.SpatialReference = pPolygon.SpatialReference

  Set pTopoOp = pIntPolygon
  pTopoOp.IsKnownSimple = False
  pTopoOp.Simplify

  Dim pGeometryCollection As IGeometryCollection
  Dim pExtRing As IGeometryCollection
  Dim pIntRingBag As IGeometryCollection
  Set pGeometryCollection = pIntPolygon.ConnectedComponentBag
  Dim pIntGeoCol As IGeometryCollection
  Dim pIntPoly As IPolygon4
  Dim pOutGeoCol As IGeometryCollection
  Dim pOutPoly As IPolygon4

  Dim pSubPoly As IPolygon4
  Dim pSubRing As IRing
  Dim anIndex As Long
  Dim anIndex2 As Long

  Dim pClipRaster As IRaster
  Dim pOuterRaster As IRaster
  Dim pInnerRaster As IRaster

  Dim pNegativeGeometryCollection As IGeometryCollection
  Dim anIndex3 As Long
  Dim pNegPoly As IPolygon4

  Dim pRasMakerOp As IRasterMakerOp
  Set pRasMakerOp = New RasterMakerOp
  SetSpatialAnalysisSettings pRasMakerOp, pRastAnalysisEnv
  Dim pCondOp As IConditionalOp
  Set pCondOp = New RasterConditionalOp
  SetSpatialAnalysisSettings pCondOp, pRastAnalysisEnv
  Dim pLogicOp As ILogicalOp
  Set pLogicOp = New RasterMathOps
  SetSpatialAnalysisSettings pLogicOp, pRastAnalysisEnv
  Dim pMathOp As IMathOp
  Set pMathOp = New RasterMathOps
  SetSpatialAnalysisSettings pMathOp, pRastAnalysisEnv

  Dim pRasLayer As IRasterLayer
  Dim pTestGeometry As IGeometry
  Dim pTestGeoColl As IGeometryCollection
  Dim pSegmentCollection1 As ISegmentCollection
  Dim pSegment1 As ISegment
  Dim pGeometryType As esriGeometryType
  Dim pEllArcPolygon As IPolygon4

  Dim pFinalGrid As IRaster
  Set pFinalGrid = pRasMakerOp.MakeConstant(0, True)
  Set pClipRaster = pRasMakerOp.MakeConstant(1, True)

  If booShowProgress Then
    pPro.MaxRange = pGeometryCollection.GeometryCount + 2
    pPro.StepValue = 1
    pPro.Show
  End If

  DoEvents

  For anIndex = 0 To pGeometryCollection.GeometryCount - 1

    Set pSubPoly = pGeometryCollection.Geometry(anIndex)
    Set pExtRing = pSubPoly.ExteriorRingBag
    Set pSubRing = pExtRing.Geometry(0)
    Set pOutGeoCol = New Polygon
    pOutGeoCol.AddGeometry pSubRing
    Set pOutPoly = pOutGeoCol
    Set pTopoOp = pOutPoly
    pTopoOp.IsKnownSimple = False
    pTopoOp.Simplify

    Set pSegmentCollection1 = pOutPoly
    Set pSegment1 = pSegmentCollection1.Segment(0)
    pGeometryType = pSegment1.GeometryType

    If pGeometryType = esriGeometryCircularArc Then
      Set pOuterRaster = pExtractionOp.Circle(pClipRaster, pSegment1, Not SaveInside)
    ElseIf pGeometryType = esriGeometryEllipticArc Then
      Set pEllArcPolygon = EllipticArcToPolygon(pSegmentCollection1, 75)
      Set pOuterRaster = pExtractionOp.Polygon(pClipRaster, pEllArcPolygon, Not SaveInside)
    ElseIf pGeometryType = esriGeometryEnvelope Then
      Set pOuterRaster = pExtractionOp.Rectangle(pClipRaster, pOutPoly, Not SaveInside)
    Else
      Set pOuterRaster = pExtractionOp.Polygon(pClipRaster, pOutPoly, Not SaveInside)
    End If
    Set pOuterRaster = pLogicOp.IsNull(pOuterRaster)

    If pSubPoly.InteriorRingCount(pSubRing) > 0 Then
      Set pIntRingBag = pSubPoly.InteriorRingBag(pSubRing)

      For anIndex2 = 0 To pIntRingBag.GeometryCount - 1
        Set pIntGeoCol = New Polygon
        pIntGeoCol.AddGeometry pIntRingBag.Geometry(anIndex2)
        Set pIntPoly = pIntGeoCol
        Set pTopoOp = pIntPoly
        pTopoOp.IsKnownSimple = False
        pTopoOp.Simplify

        Set pSegmentCollection1 = pIntPoly
        Set pSegment1 = pSegmentCollection1.Segment(0)
        pGeometryType = pSegment1.GeometryType

        If pGeometryType = esriGeometryCircularArc Then
          Set pInnerRaster = pExtractionOp.Circle(pClipRaster, pSegment1, SaveInside)
          Set pInnerRaster = pLogicOp.IsNull(pInnerRaster)
          Set pOuterRaster = pMathOp.Times(pInnerRaster, pOuterRaster)
        ElseIf pGeometryType = esriGeometryEllipticArc Then
          Set pEllArcPolygon = EllipticArcToPolygon(pSegmentCollection1, 75)
          Set pInnerRaster = pExtractionOp.Polygon(pClipRaster, pEllArcPolygon, SaveInside)
          Set pInnerRaster = pLogicOp.IsNull(pInnerRaster)
          Set pOuterRaster = pMathOp.Times(pInnerRaster, pOuterRaster)
        ElseIf pGeometryType = esriGeometryEnvelope Then
          Set pInnerRaster = pExtractionOp.Rectangle(pClipRaster, pIntPoly, SaveInside)
          Set pInnerRaster = pLogicOp.IsNull(pInnerRaster)
          Set pOuterRaster = pMathOp.Times(pInnerRaster, pOuterRaster)
        Else
          Set pNegativeGeometryCollection = pIntPoly.ConnectedComponentBag
          For anIndex3 = 0 To pNegativeGeometryCollection.GeometryCount - 1
            Set pNegPoly = pNegativeGeometryCollection.Geometry(anIndex3)
            Set pTopoOp = pNegPoly
            pTopoOp.IsKnownSimple = False
            pTopoOp.Simplify
            Set pInnerRaster = pExtractionOp.Polygon(pClipRaster, pNegPoly, SaveInside)
            Set pInnerRaster = pLogicOp.IsNull(pInnerRaster)
            Set pOuterRaster = pMathOp.Times(pInnerRaster, pOuterRaster)
          Next anIndex3
        End If

        DoEvents
      Next anIndex2

    End If

    Set pFinalGrid = pMathOp.Plus(pFinalGrid, pOuterRaster)
    If booShowProgress Then
      pPro.Step
    End If
    DoEvents
  Next anIndex

  Set pFinalGrid = pCondOp.SetNull(pLogicOp.EqualTo(pFinalGrid, pRasMakerOp.MakeConstant(0, True)), pFinalGrid)
  If booShowProgress Then
    pPro.Step
  End If
  Set ClipRasterToPolygon = pMathOp.Times(pFinalGrid, pRaster)
  If booShowProgress Then
    pPro.Step
  End If

  DoEvents

  pRastAnalysisEnv.RestoreToPreviousDefaultEnvironment

  If booShowProgress Then
    pPro.Hide
  End If

  Set pSBar = Nothing
  Set pPro = Nothing
  Set pSegmentCollectionCurves = Nothing
  Set pSegmentCurve = Nothing
  Set pNewPoints = Nothing
  Set pNewPolygon = Nothing
  Set pExtractionOp = Nothing
  Set pRastAnalysisEnv = Nothing
  Set pSpatialReference = Nothing
  Set pPolySpatRef = Nothing
  Set pEnvelope = Nothing
  Set pTempEnv = Nothing
  Set pRasterAnalysisProps = Nothing
  Set pRasterProps = Nothing
  Set pTopoOp = Nothing
  Set pPolyEnvelope = Nothing
  Set pRastEnvelope = Nothing
  Set pIntPolygon = Nothing
  Set pGeometryCollection = Nothing
  Set pExtRing = Nothing
  Set pIntRingBag = Nothing
  Set pIntGeoCol = Nothing
  Set pIntPoly = Nothing
  Set pOutGeoCol = Nothing
  Set pOutPoly = Nothing
  Set pSubPoly = Nothing
  Set pSubRing = Nothing
  Set pClipRaster = Nothing
  Set pOuterRaster = Nothing
  Set pInnerRaster = Nothing
  Set pNegativeGeometryCollection = Nothing
  Set pNegPoly = Nothing
  Set pRasMakerOp = Nothing
  Set pCondOp = Nothing
  Set pLogicOp = Nothing
  Set pMathOp = Nothing
  Set pRasLayer = Nothing
  Set pTestGeometry = Nothing
  Set pTestGeoColl = Nothing
  Set pSegmentCollection1 = Nothing
  Set pSegment1 = Nothing
  Set pEllArcPolygon = Nothing
  Set pFinalGrid = Nothing

End Function

Public Function EllipticArcToPolygon2(SegCollection As ISegmentCollection, NumVertices As Long) As IMultipoint

On Error GoTo erh

  Dim pCurve As ICurve
  Dim pGeometry As IGeometry

  Dim anIndex As Long
  Dim lngSegCount As Long
  lngSegCount = SegCollection.SegmentCount - 1
  Dim theLength As Double
  theLength = 0
  Dim theTestLength As Double
  Dim lngLengths() As Long
  ReDim lngLengths(lngSegCount)
  For anIndex = 0 To lngSegCount
    theTestLength = SegCollection.Segment(anIndex).length
    theLength = theLength + theTestLength
    lngLengths(anIndex) = theTestLength
  Next anIndex

  Dim pProportion As Double
  Dim lngVertices() As Long
  Dim lngNumVertices As Long
  ReDim lngVertices(lngSegCount)
  For anIndex = 0 To lngSegCount
    lngNumVertices = Int((lngLengths(anIndex) / theLength) * NumVertices)
    If lngNumVertices < 8 Then lngNumVertices = 8
    lngVertices(anIndex) = lngNumVertices
  Next anIndex

  Dim pMpt As IPointCollection
  Set pMpt = New Multipoint
  Dim pPoint As IPoint
  Set pPoint = New Point
  Dim pClone As IClone

  Dim pRatio As Double
  Dim anIndex2 As Long

  For anIndex = 0 To lngSegCount
    lngNumVertices = lngVertices(anIndex)
    pRatio = 1 / lngNumVertices
    Set pCurve = SegCollection.Segment(anIndex)

    For anIndex2 = 0 To lngNumVertices
      pCurve.QueryPoint 0, (pRatio * anIndex2), True, pPoint
      Set pClone = pPoint

      pMpt.AddPoint pClone.Clone
    Next anIndex2
  Next anIndex

  Set EllipticArcToPolygon2 = pMpt

  Set pCurve = Nothing
  Set pGeometry = Nothing
  Set pMpt = Nothing
  Set pPoint = Nothing
  Set pClone = Nothing

    Exit Function

erh:
    MsgBox "Failed in EllipticArcToPolygon2: " & err.Description
End Function

Public Function DistributePointsAlongShape(pCurve As ICurve, SeparationDistance As Double) As IPointCollection

On Error GoTo erh

  Dim anIndex As Long
  Dim theLength As Double
  theLength = pCurve.length

  Dim pMpt As IPointCollection
  Set pMpt = New Multipoint
  Dim pPoint As IPoint
  Set pPoint = New Point
  Dim pClone As IClone

  Dim dblRatio As Double
  dblRatio = SeparationDistance / theLength

  Dim theCurrentDist As Double
  theCurrentDist = 0

  Do While theCurrentDist < 1
    pCurve.QueryPoint esriNoExtension, theCurrentDist, True, pPoint
    Set pClone = pPoint
    pMpt.AddPoint pClone.Clone
    theCurrentDist = theCurrentDist + dblRatio
  Loop
  pCurve.QueryPoint esriNoExtension, 1, True, pPoint
  Set pClone = pPoint
  pMpt.AddPoint pClone.Clone

  Dim pGeometry As IGeometry
  Set pGeometry = pMpt
  Set pGeometry.SpatialReference = pCurve.SpatialReference

  Set DistributePointsAlongShape = pMpt

  Set pMpt = Nothing
  Set pPoint = Nothing
  Set pClone = Nothing
  Set pGeometry = Nothing

    Exit Function

erh:
    MsgBox "Failed in DistributePointsAlongShape: " & err.Description
End Function

Public Function CellValues(pPoints As IPointCollection, pRaster As IRaster) As esriSystem.IVariantArray

    Dim pRP As IRasterProps
    Set pRP = pRaster

    Dim dblCellSize As Double
    dblCellSize = ReturnCellSize(pRaster)

    Dim pExtent As IEnvelope
    Set pExtent = pRP.Extent
    Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
    pExtent.QueryCoords X1, Y1, X2, Y2

    Dim pPB As IPixelBlock3
    Dim dWidth As Double, dHeight As Double
    dWidth = pRP.Width
    dHeight = pRP.Height

    Dim pPnt As IPnt
    Set pPnt = New Pnt
    pPnt.SetCoords dWidth, dHeight

    Dim pOrigin As IPnt
    Set pOrigin = New Pnt
    pOrigin.SetCoords 0, 0

    Set pPB = pRaster.CreatePixelBlock(pPnt)
    pRaster.Read pOrigin, pPB

    Dim lngIndex As Long
    Dim dblCellValue As Double
    Dim pPoint As IPoint
    Dim dx As Double, dy As Double
    Dim nX As Double, ny As Double
    Dim iX As Long, iY As Long

    Dim pOutArray As esriSystem.IVariantArray
    Set pOutArray = New esriSystem.varArray

    Dim vCellValue As Variant

    For lngIndex = 0 To pPoints.PointCount - 1
      Set pPoint = pPoints.Point(lngIndex)
      If pPoint.x < X1 Or pPoint.x > X2 Or pPoint.Y < Y1 Or pPoint.Y > Y2 Then
        pOutArray.Add Null
      Else

        dx = pPoint.x - X1
        dy = Y2 - pPoint.Y

        nX = dx / dblCellSize
        ny = dy / dblCellSize

        iX = Int(nX)
        iY = Int(ny)

        If (iX < 0) Then iX = 0
        If (iY < 0) Then iY = 0
        If (iX > pRP.Width - 1) Then
          iX = pRP.Width - 1
        End If
        If (iY > pRP.Height - 1) Then
          iY = pRP.Height - 1
        End If

        vCellValue = pPB.GetVal(0, iX, iY)
        Debug.Print "From CellValues function..." & vCellValue
        If IsEmpty(vCellValue) Then
          pOutArray.Add Null
        Else
          pOutArray.Add CDbl(vCellValue)
        End If
      End If
    Next lngIndex
    Set CellValues = pOutArray

  Set pRP = Nothing
  Set pExtent = Nothing
  Set pPB = Nothing
  Set pPnt = Nothing
  Set pOrigin = Nothing
  Set pPoint = Nothing
  Set pOutArray = Nothing

End Function

Public Function CellValue(pPoint As IPoint, pRaster As IRaster) As Variant

    Dim pRP As IRasterProps
    Set pRP = pRaster

    Dim dblCellSize As Double
    dblCellSize = ReturnCellSize(pRaster)

    Dim pExtent As IEnvelope
    Set pExtent = pRP.Extent
    Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
    pExtent.QueryCoords X1, Y1, X2, Y2

    If pPoint.x < X1 Or pPoint.x > X2 Or pPoint.Y < Y1 Or pPoint.Y > Y2 Then
      CellValue = Null
      Exit Function
    End If

    Dim pCellPoint As IPoint  'Get dx dy from left-top
    Dim dx As Double, dy As Double
    dx = pPoint.x - X1
    dy = Y2 - pPoint.Y

    Dim nX As Double, ny As Double
    nX = dx / dblCellSize
    ny = dy / dblCellSize

    Dim iX As Long, iY As Long
    iX = Int(nX)
    iY = Int(ny)

    If (iX < 0) Then iX = 0
    If (iY < 0) Then iY = 0
    If (iX > pRP.Width - 1) Then
      iX = pRP.Width - 1
    End If
    If (iY > pRP.Height - 1) Then
      iY = pRP.Height - 1
    End If

    Set pCellPoint = New Point
    pCellPoint.PutCoords CDbl(iX), CDbl(iY)

    Dim pPB As IPixelBlock3
    Dim dWidth As Double, dHeight As Double
    dWidth = pRP.Width
    dHeight = pRP.Height

    Dim pPnt As IPnt
    Set pPnt = New Pnt
    pPnt.SetCoords dWidth, dHeight

    Dim pOrigin As IPnt
    Set pOrigin = New Pnt
    pOrigin.SetCoords 0, 0

    Set pPB = pRaster.CreatePixelBlock(pPnt)

    pRaster.Read pOrigin, pPB
    Dim vCellValue As Variant
    vCellValue = pPB.GetVal(0, iX, iY)
    Debug.Print "From function..." & vCellValue
    If IsEmpty(vCellValue) Then
      CellValue = Null
    Else
      CellValue = CDbl(vCellValue)
    End If

  Set pRP = Nothing
  Set pExtent = Nothing
  Set pCellPoint = Nothing
  Set pPB = Nothing
  Set pPnt = Nothing
  Set pOrigin = Nothing

End Function

Public Function CleanPolyline(ByVal pPolyline As IPolyline, pBoundary As IPolyline) As IPolyline

  Dim pPointCol As IPointCollection
  Set pPointCol = pPolyline

  Dim pNewPointCol As IPointCollection
  Set pNewPointCol = New Polyline

  Dim pSegment As IPointCollection
  Dim pStartPoint As IPoint
  Dim pEndPoint As IPoint
  Dim pClone As IClone

  Dim pCount As Long
  pCount = pPointCol.PointCount

  Set pStartPoint = pPointCol.Point(0)
  Set pClone = pStartPoint
  pNewPointCol.AddPoint pClone.Clone

  Dim anIndex1 As Long
  Dim anIndex2 As Long
  Dim anIndexStep As Long

  Dim pRelOp As IRelationalOperator

  anIndex1 = 1
  anIndexStep = 0

  Do Until anIndex1 = pPointCol.PointCount - 1
    Set pEndPoint = pPointCol.Point(anIndex1)
    Set pSegment = New Polyline
    pSegment.AddPoint pStartPoint
    pSegment.AddPoint pEndPoint
    Set pRelOp = pSegment
    If pRelOp.Disjoint(pBoundary) Then
      anIndex1 = anIndex1 + 1
    Else
      anIndex2 = anIndex1 - 1
      Do Until anIndex2 <= anIndexStep
        Set pStartPoint = pPointCol.Point(anIndex2)
        Set pSegment = New Polyline
        pSegment.AddPoint pStartPoint
        pSegment.AddPoint pEndPoint
        Set pRelOp = pSegment
        If pRelOp.Disjoint(pBoundary) Then
          anIndex2 = anIndex2 - 1
        Else
          Set pClone = pStartPoint
          pNewPointCol.AddPoint pClone.Clone
          anIndexStep = anIndex2 + 1
          Set pStartPoint = pPointCol.Point(anIndexStep)
          Set pClone = pStartPoint
          pNewPointCol.AddPoint pClone.Clone
          anIndex1 = anIndexStep + 1
          Exit Do
        End If
      Loop
    End If
  Loop
  Set pClone = pPointCol.Point(pPointCol.PointCount - 1)
  pNewPointCol.AddPoint pClone.Clone
  Set CleanPolyline = pNewPointCol
  Set CleanPolyline.SpatialReference = pPolyline.SpatialReference

  Set pPointCol = Nothing
  Set pNewPointCol = Nothing
  Set pSegment = Nothing
  Set pStartPoint = Nothing
  Set pEndPoint = Nothing
  Set pClone = Nothing
  Set pRelOp = Nothing

End Function

Private Function IsCellNaN(expression As Variant) As Boolean

  On Error Resume Next
  If Not IsNumeric(expression) Then
    IsCellNaN = False
    Exit Function
  End If
  If (CStr(expression) = "-1.#QNAN") Or (CStr(expression) = "1,#QNAN") Then ' can vary by locale
    IsCellNaN = True
  Else
    IsCellNaN = False
  End If

End Function

Public Function CellValues4CellInterp(pPoints As IPointCollection, pRaster As IRaster, _
    Optional lngBandIndex As Long = 0) As esriSystem.IVariantArray

    Dim pRP As IRasterProps
    Set pRP = pRaster

    Dim dblCellSizeX As Double
    Dim dblCellSizeY As Double
    dblCellSizeX = ReturnPixelWidth(pRaster)
    dblCellSizeY = ReturnPixelHeight(pRaster)

    Dim dblHalfCellX As Double
    dblHalfCellX = dblCellSizeX / 2
    Dim dblHalfCellY As Double
    dblHalfCellY = dblCellSizeY / 2

    Dim pExtent As IEnvelope
    Set pExtent = pRP.Extent
    Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
    pExtent.QueryCoords X1, Y1, X2, Y2

    Dim pPB As IPixelBlock3

    Dim dWidth As Double, dHeight As Double
    dWidth = pRP.Width
    dHeight = pRP.Height

    Dim pPnt As IPnt
    Set pPnt = New Pnt
    pPnt.SetCoords 2, 2
    Set pPB = pRaster.CreatePixelBlock(pPnt)

    Dim pOrigin As IPnt

    Dim lngIndex As Long
    Dim dblCellValue As Double
    Dim pPoint As IPoint
    Dim dx As Double, dy As Double
    Dim nX As Double, ny As Double
    Dim dblXRemainder As Double, dblYRemainder As Double
    Dim iX As Long, iY As Long
    Dim lngMaxX As Long, lngMaxY As Long

    lngMaxX = pRP.Width - 1
    lngMaxY = pRP.Height - 1

    Dim bytQuadrant As Byte       ' 1 FOR NE, 2 FOR NW, 3 FOR SW, 4 FOR SE
    Dim varInterpVal As Variant

    Dim dblPropX As Double
    Dim dblPropY As Double

    Dim pOutArray As esriSystem.IVariantArray
    Set pOutArray = New esriSystem.varArray

    Dim vCellValueNE As Variant
    Dim vCellValueNW As Variant
    Dim vCellValueSE As Variant
    Dim vCellValueSW As Variant

    Dim booIsNull As Boolean
    Dim dblWestProp As Double
    Dim dblEastProp As Double

    For lngIndex = 0 To pPoints.PointCount - 1
      Set pPoint = pPoints.Point(lngIndex)

      If pPoint.x < X1 Or pPoint.x > X2 Or pPoint.Y < Y1 Or pPoint.Y > Y2 Then
        pOutArray.Add Null
      Else

        dx = pPoint.x - X1
        dy = Y2 - pPoint.Y

        nX = dx / dblCellSizeX
        ny = dy / dblCellSizeY

        iX = Int(nX)
        iY = Int(ny)

        If (iX < 0) Then iX = 0
        If (iY < 0) Then iY = 0
        If (iX > lngMaxX) Then
          iX = lngMaxX
        End If
        If (iY > lngMaxY - 1) Then
          iY = lngMaxY - 1
        End If

        dblXRemainder = (nX - iX) * dblCellSizeX
        dblYRemainder = (ny - iY) * dblCellSizeY

        If dblYRemainder < dblHalfCellY Then                  ' ON NORTH SIDE OF CELL, SOUTH HALF OF PIXEL BLOCK
          dblPropY = (dblYRemainder + dblHalfCellY) / dblCellSizeY
          If dblXRemainder > dblHalfCellX Then                ' ON EAST SIDE OF CELL, WEST HALF OF PIXEL BLOCK
            bytQuadrant = 1                                   ' ON NORTHEAST CORNER OF CELL, SOUTHWEST CORNER OF PIXEL BLOCK
            dblPropX = 1 - ((dblXRemainder - dblHalfCellX) / dblCellSizeX)
          Else                                                ' ON WEST SIDE OF CELL, EAST HALF OF PIXEL BLOCK
            bytQuadrant = 2                                   ' ON NORTHWEST CORNER OF CELL, SOUTHEAST CORNER OF PIXEL BLOCK
            dblPropX = (dblHalfCellX + dblXRemainder) / dblCellSizeX
          End If
        Else                                                  ' ON SOUTH SIDE OF CELL, NORTH HALF OF PIXEL BLOCK
          dblPropY = 1 - ((dblYRemainder - dblHalfCellY) / dblCellSizeY)
          If dblXRemainder > dblHalfCellX Then                ' ON EAST SIDE, WEST HALF OF PIXEL BLOCK
            bytQuadrant = 4                                   ' ON SOUTHEAST CORNER OF CELL, NORTHWEST CORNER OF PIXEL BLOCK
            dblPropX = 1 - ((dblXRemainder - dblHalfCellX) / dblCellSizeX)
          Else                                                ' ON WEST SIDE OF CELL, EAST HALF OF PIXEL BLOCK
            bytQuadrant = 3                                   ' ON SOUTHWEST CORNER OF CELL, NORTHEAST CORNER OF PIXEL BLOCK
            dblPropX = (dblHalfCellX + dblXRemainder) / dblCellSizeX
          End If
        End If

        Set pOrigin = New Pnt

        booIsNull = False
        Select Case bytQuadrant
          Case 1              ' NORTHEAST                =================
            If iX = lngMaxX Or iY = 0 Then
              booIsNull = True
            Else
              pOrigin.SetCoords iX, iY - 1
              pRaster.Read pOrigin, pPB
              vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
              vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
              vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
              vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
              If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
                IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                dblWestProp = (CDbl(vCellValueNW) * (1 - dblPropY)) + (CDbl(vCellValueSW) * dblPropY)
                dblEastProp = (CDbl(vCellValueNE) * (1 - dblPropY)) + (CDbl(vCellValueSE) * dblPropY)
                varInterpVal = CVar((dblWestProp * dblPropX) + (dblEastProp * (1 - dblPropX)))
              End If
            End If
          Case 2              ' NORTHWEST                =================
            If iX = 0 Or iY = 0 Then
              booIsNull = True
            Else
              pOrigin.SetCoords iX - 1, iY - 1
              pRaster.Read pOrigin, pPB
              vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
              vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
              vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
              vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
              If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
                IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                dblWestProp = (CDbl(vCellValueNW) * (1 - dblPropY)) + (CDbl(vCellValueSW) * dblPropY)
                dblEastProp = (CDbl(vCellValueNE) * (1 - dblPropY)) + (CDbl(vCellValueSE) * dblPropY)
                varInterpVal = CVar((dblWestProp * (1 - dblPropX)) + (dblEastProp * dblPropX))
              End If
            End If
          Case 3              ' SOUTHWEST                =================
            If iX = 0 Or iY = lngMaxY Then
              booIsNull = True
            Else
              pOrigin.SetCoords iX - 1, iY
              pRaster.Read pOrigin, pPB
              vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
              vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
              vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
              vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
              If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
                IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                dblWestProp = (CDbl(vCellValueNW) * dblPropY) + (CDbl(vCellValueSW) * (1 - dblPropY))
                dblEastProp = (CDbl(vCellValueNE) * dblPropY) + (CDbl(vCellValueSE) * (1 - dblPropY))
                varInterpVal = CVar((dblWestProp * (1 - dblPropX)) + (dblEastProp * dblPropX))
              End If
            End If
          Case 4              ' SOUTHEAST                =================
            If iX = lngMaxX Or iY = lngMaxY Then
              booIsNull = True
            Else
              pOrigin.SetCoords iX, iY
              pRaster.Read pOrigin, pPB
              vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
              vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
              vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
              vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
              If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
                IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                dblWestProp = (CDbl(vCellValueNW) * dblPropY) + (CDbl(vCellValueSW) * (1 - dblPropY))
                dblEastProp = (CDbl(vCellValueNE) * dblPropY) + (CDbl(vCellValueSE) * (1 - dblPropY))
                varInterpVal = CVar((dblWestProp * dblPropX) + (dblEastProp * (1 - dblPropX)))
              End If

            End If
        End Select

        If booIsNull Then
          pOutArray.Add Null
        Else
          pOutArray.Add CDbl(varInterpVal)
        End If

          "   dblPropX = " & CStr(dblPropX) & vbCrLf & _
          "   dblPropY = " & CStr(dblPropY) & vbCrLf & _
          "   Quadrant = " & CStr(bytQuadrant) & vbCrLf & _
          "   vCellValueNW = " & CStr(vCellValueNW) & vbCrLf & _
          "   vCellValueNE = " & CStr(vCellValueNE) & vbCrLf & _
          "   vCellValueSW = " & CStr(vCellValueSW) & vbCrLf & _
          "   vCellValueSE = " & CStr(vCellValueSE) & vbCrLf & _
          "   dblWestProp = " & CStr(dblWestProp) & vbCrLf & _
          "   dblEastProp = " & CStr(dblEastProp) & vbCrLf & _
          "   Interpolated Value = " & CStr(varInterpVal)
      End If
    Next lngIndex
    Set CellValues4CellInterp = pOutArray

End Function

Public Function CellValues4CellInterp_ByNumbers2(dblArray() As Double, pRaster As IRaster, _
    dblCellSizeX As Double, dblCellSizeY As Double, X1 As Double, X2 As Double, _
    Y1 As Double, Y2 As Double, lngRastWidth As Long, lngRastHeight As Long, varRasterArray As Variant, _
    varNullval As Variant, Optional lngBandIndex As Long = 0) As Variant()

    Dim dblHalfCellX As Double
    dblHalfCellX = dblCellSizeX / 2
    Dim dblHalfCellY As Double
    dblHalfCellY = dblCellSizeY / 2

    Dim lngIndex As Long
    Dim dblCellValue As Double
    Dim dx As Double, dy As Double
    Dim nX As Double, ny As Double
    Dim dblXRemainder As Double, dblYRemainder As Double
    Dim iX As Long, iY As Long
    Dim lngMaxX As Long, lngMaxY As Long

    lngMaxX = lngRastWidth - 1
    lngMaxY = lngRastHeight - 1

    Dim bytQuadrant As Byte       ' 1 FOR NE, 2 FOR NW, 3 FOR SW, 4 FOR SE
    Dim varInterpVal As Variant

    Dim dblPropX As Double
    Dim dblPropY As Double

    Dim varReturn() As Variant
    ReDim varReturn(UBound(dblArray, 2))

    Dim vCellValueNE As Variant
    Dim vCellValueNW As Variant
    Dim vCellValueSE As Variant
    Dim vCellValueSW As Variant

    Dim booIsNull As Boolean
    Dim dblWestProp As Double
    Dim dblEastProp As Double
    Dim dblCoordX As Double
    Dim dblCoordY As Double

    For lngIndex = 0 To UBound(dblArray, 2)
      dblCoordX = dblArray(0, lngIndex)
      dblCoordY = dblArray(1, lngIndex)
      If dblCoordX < X1 Or dblCoordX > X2 Or dblCoordY < Y1 Or dblCoordY > Y2 Then
        varReturn(lngIndex) = Null
      Else

        dx = dblCoordX - X1
        dy = Y2 - dblCoordY

        nX = dx / dblCellSizeX
        ny = dy / dblCellSizeY

        iX = Int(nX)
        iY = Int(ny)

        If (iX < 0) Then iX = 0
        If (iY < 0) Then iY = 0
        If (iX > lngMaxX) Then
          iX = lngMaxX
        End If
        If (iY > lngMaxY - 1) Then
          iY = lngMaxY - 1
        End If

        dblXRemainder = (nX - iX) * dblCellSizeX
        dblYRemainder = (ny - iY) * dblCellSizeY

        If dblYRemainder < dblHalfCellY Then                  ' ON NORTH SIDE OF CELL, SOUTH HALF OF PIXEL BLOCK
          dblPropY = (dblYRemainder + dblHalfCellY) / dblCellSizeY
          If dblXRemainder > dblHalfCellX Then                ' ON EAST SIDE OF CELL, WEST HALF OF PIXEL BLOCK
            bytQuadrant = 1                                   ' ON NORTHEAST CORNER OF CELL, SOUTHWEST CORNER OF PIXEL BLOCK
            dblPropX = 1 - ((dblXRemainder - dblHalfCellX) / dblCellSizeX)
          Else                                                ' ON WEST SIDE OF CELL, EAST HALF OF PIXEL BLOCK
            bytQuadrant = 2                                   ' ON NORTHWEST CORNER OF CELL, SOUTHEAST CORNER OF PIXEL BLOCK
            dblPropX = (dblHalfCellX + dblXRemainder) / dblCellSizeX
          End If
        Else                                                  ' ON SOUTH SIDE OF CELL, NORTH HALF OF PIXEL BLOCK
          dblPropY = 1 - ((dblYRemainder - dblHalfCellY) / dblCellSizeY)
          If dblXRemainder > dblHalfCellX Then                ' ON EAST SIDE, WEST HALF OF PIXEL BLOCK
            bytQuadrant = 4                                   ' ON SOUTHEAST CORNER OF CELL, NORTHWEST CORNER OF PIXEL BLOCK
            dblPropX = 1 - ((dblXRemainder - dblHalfCellX) / dblCellSizeX)
          Else                                                ' ON WEST SIDE OF CELL, EAST HALF OF PIXEL BLOCK
            bytQuadrant = 3                                   ' ON SOUTHWEST CORNER OF CELL, NORTHEAST CORNER OF PIXEL BLOCK
            dblPropX = (dblHalfCellX + dblXRemainder) / dblCellSizeX
          End If
        End If

        booIsNull = False
        Select Case bytQuadrant
          Case 1              ' NORTHEAST                =================
            If iX = lngMaxX Or iY = 0 Then
              booIsNull = True
            Else

              vCellValueNW = varRasterArray(iX, iY - 1)
              vCellValueSW = varRasterArray(iX, iY)
              vCellValueNE = varRasterArray(iX + 1, iY - 1)
              vCellValueSE = varRasterArray(iX + 1, iY)

              If (varNullval = vCellValueNW) Or (varNullval = vCellValueNE) Or (varNullval = vCellValueSW) Or _
                (varNullval = vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                If Abs(CDbl(vCellValueNW)) > 10 ^ 10 Or Abs(CDbl(vCellValueNE)) > 10 ^ 10 Or _
                    Abs(CDbl(vCellValueSE)) > 10 ^ 10 Or Abs(CDbl(vCellValueSW)) > 10 ^ 10 Then
                  booIsNull = True
                Else
                  dblWestProp = (CDbl(vCellValueNW) * (1 - dblPropY)) + (CDbl(vCellValueSW) * dblPropY)
                  dblEastProp = (CDbl(vCellValueNE) * (1 - dblPropY)) + (CDbl(vCellValueSE) * dblPropY)
                  varInterpVal = CVar((dblWestProp * dblPropX) + (dblEastProp * (1 - dblPropX)))
                End If
              End If
            End If
          Case 2              ' NORTHWEST                =================
            If iX = 0 Or iY = 0 Then
              booIsNull = True
            Else

              vCellValueNW = varRasterArray(iX - 1, iY - 1)
              vCellValueSW = varRasterArray(iX - 1, iY)
              vCellValueNE = varRasterArray(iX, iY - 1)
              vCellValueSE = varRasterArray(iX, iY)

              If (varNullval = vCellValueNW) Or (varNullval = vCellValueNE) Or (varNullval = vCellValueSW) Or _
                (varNullval = vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                If Abs(CDbl(vCellValueNW)) > 10 ^ 10 Or Abs(CDbl(vCellValueNE)) > 10 ^ 10 Or _
                    Abs(CDbl(vCellValueSE)) > 10 ^ 10 Or Abs(CDbl(vCellValueSW)) > 10 ^ 10 Then
                  booIsNull = True
                Else
                  dblWestProp = (CDbl(vCellValueNW) * (1 - dblPropY)) + (CDbl(vCellValueSW) * dblPropY)
                  dblEastProp = (CDbl(vCellValueNE) * (1 - dblPropY)) + (CDbl(vCellValueSE) * dblPropY)
                  varInterpVal = CVar((dblWestProp * (1 - dblPropX)) + (dblEastProp * dblPropX))
                End If
              End If
            End If
          Case 3              ' SOUTHWEST                =================
            If iX = 0 Or iY = lngMaxY Then
              booIsNull = True
            Else

              vCellValueNW = varRasterArray(iX - 1, iY)
              vCellValueSW = varRasterArray(iX - 1, iY + 1)
              vCellValueNE = varRasterArray(iX, iY)
              vCellValueSE = varRasterArray(iX, iY + 1)

              If (varNullval = vCellValueNW) Or (varNullval = vCellValueNE) Or (varNullval = vCellValueSW) Or _
                (varNullval = vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                If Abs(CDbl(vCellValueNW)) > 10 ^ 10 Or Abs(CDbl(vCellValueNE)) > 10 ^ 10 Or _
                    Abs(CDbl(vCellValueSE)) > 10 ^ 10 Or Abs(CDbl(vCellValueSW)) > 10 ^ 10 Then
                  booIsNull = True
                Else
                  dblWestProp = (CDbl(vCellValueNW) * dblPropY) + (CDbl(vCellValueSW) * (1 - dblPropY))
                  dblEastProp = (CDbl(vCellValueNE) * dblPropY) + (CDbl(vCellValueSE) * (1 - dblPropY))
                  varInterpVal = CVar((dblWestProp * (1 - dblPropX)) + (dblEastProp * dblPropX))
                End If
              End If
            End If
          Case 4              ' SOUTHEAST                =================
            If iX = lngMaxX Or iY = lngMaxY Then
              booIsNull = True
            Else

              vCellValueNW = varRasterArray(iX, iY)
              vCellValueSW = varRasterArray(iX, iY + 1)
              vCellValueNE = varRasterArray(iX + 1, iY)
              vCellValueSE = varRasterArray(iX + 1, iY + 1)

              If (varNullval = vCellValueNW) Or (varNullval = vCellValueNE) Or (varNullval = vCellValueSW) Or _
                (varNullval = vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                If Abs(CDbl(vCellValueNW)) > 10 ^ 10 Or Abs(CDbl(vCellValueNE)) > 10 ^ 10 Or _
                    Abs(CDbl(vCellValueSE)) > 10 ^ 10 Or Abs(CDbl(vCellValueSW)) > 10 ^ 10 Then
                  booIsNull = True
                Else
                  dblWestProp = (CDbl(vCellValueNW) * dblPropY) + (CDbl(vCellValueSW) * (1 - dblPropY))
                  dblEastProp = (CDbl(vCellValueNE) * dblPropY) + (CDbl(vCellValueSE) * (1 - dblPropY))
                  varInterpVal = CVar((dblWestProp * dblPropX) + (dblEastProp * (1 - dblPropX)))
                End If
              End If

            End If
        End Select

        If booIsNull Then
          varReturn(lngIndex) = Null
        Else
          varReturn(lngIndex) = varInterpVal
        End If

          "   dblPropX = " & CStr(dblPropX) & vbCrLf & _
          "   dblPropY = " & CStr(dblPropY) & vbCrLf & _
          "   Quadrant = " & CStr(bytQuadrant) & vbCrLf & _
          "   vCellValueNW = " & CStr(vCellValueNW) & vbCrLf & _
          "   vCellValueNE = " & CStr(vCellValueNE) & vbCrLf & _
          "   vCellValueSW = " & CStr(vCellValueSW) & vbCrLf & _
          "   vCellValueSE = " & CStr(vCellValueSE) & vbCrLf & _
          "   dblWestProp = " & CStr(dblWestProp) & vbCrLf & _
          "   dblEastProp = " & CStr(dblEastProp) & vbCrLf & _
          "   Interpolated Value = " & CStr(varInterpVal)
      End If
    Next lngIndex
    CellValues4CellInterp_ByNumbers2 = varReturn

ClearMemory:
  varInterpVal = Null
  Erase varReturn
  vCellValueNE = Null
  vCellValueNW = Null
  vCellValueSE = Null
  vCellValueSW = Null
End Function

Public Function CellValues4CellInterp_ByNumbers(dblArray() As Double, pRaster As IRaster, _
    dblCellSizeX As Double, dblCellSizeY As Double, X1 As Double, X2 As Double, _
    Y1 As Double, Y2 As Double, lngRastWidth As Long, lngRastHeight As Long, _
    Optional lngBandIndex As Long = 0) As Variant()

    Dim dblHalfCellX As Double
    dblHalfCellX = dblCellSizeX / 2
    Dim dblHalfCellY As Double
    dblHalfCellY = dblCellSizeY / 2

    Dim pPB As IPixelBlock3

    Dim pPnt As IPnt
    Set pPnt = New Pnt
    pPnt.SetCoords 2, 2
    Set pPB = pRaster.CreatePixelBlock(pPnt)

    Dim pOrigin As IPnt

    Dim lngIndex As Long
    Dim dblCellValue As Double
    Dim dx As Double, dy As Double
    Dim nX As Double, ny As Double
    Dim dblXRemainder As Double, dblYRemainder As Double
    Dim iX As Long, iY As Long
    Dim lngMaxX As Long, lngMaxY As Long

    lngMaxX = lngRastWidth - 1
    lngMaxY = lngRastHeight - 1

    Dim bytQuadrant As Byte       ' 1 FOR NE, 2 FOR NW, 3 FOR SW, 4 FOR SE
    Dim varInterpVal As Variant

    Dim dblPropX As Double
    Dim dblPropY As Double

    Dim varReturn() As Variant
    ReDim varReturn(UBound(dblArray, 2))

    Dim vCellValueNE As Variant
    Dim vCellValueNW As Variant
    Dim vCellValueSE As Variant
    Dim vCellValueSW As Variant

    Dim booIsNull As Boolean
    Dim dblWestProp As Double
    Dim dblEastProp As Double
    Dim dblCoordX As Double
    Dim dblCoordY As Double

    For lngIndex = 0 To UBound(dblArray, 2)
      dblCoordX = dblArray(0, lngIndex)
      dblCoordY = dblArray(1, lngIndex)
      If dblCoordX < X1 Or dblCoordX > X2 Or dblCoordY < Y1 Or dblCoordY > Y2 Then
        varReturn(lngIndex) = Null
      Else

        dx = dblCoordX - X1
        dy = Y2 - dblCoordY

        nX = dx / dblCellSizeX
        ny = dy / dblCellSizeY

        iX = Int(nX)
        iY = Int(ny)

        If (iX < 0) Then iX = 0
        If (iY < 0) Then iY = 0
        If (iX > lngMaxX) Then
          iX = lngMaxX
        End If
        If (iY > lngMaxY - 1) Then
          iY = lngMaxY - 1
        End If

        dblXRemainder = (nX - iX) * dblCellSizeX
        dblYRemainder = (ny - iY) * dblCellSizeY

        If dblYRemainder < dblHalfCellY Then                  ' ON NORTH SIDE OF CELL, SOUTH HALF OF PIXEL BLOCK
          dblPropY = (dblYRemainder + dblHalfCellY) / dblCellSizeY
          If dblXRemainder > dblHalfCellX Then                ' ON EAST SIDE OF CELL, WEST HALF OF PIXEL BLOCK
            bytQuadrant = 1                                   ' ON NORTHEAST CORNER OF CELL, SOUTHWEST CORNER OF PIXEL BLOCK
            dblPropX = 1 - ((dblXRemainder - dblHalfCellX) / dblCellSizeX)
          Else                                                ' ON WEST SIDE OF CELL, EAST HALF OF PIXEL BLOCK
            bytQuadrant = 2                                   ' ON NORTHWEST CORNER OF CELL, SOUTHEAST CORNER OF PIXEL BLOCK
            dblPropX = (dblHalfCellX + dblXRemainder) / dblCellSizeX
          End If
        Else                                                  ' ON SOUTH SIDE OF CELL, NORTH HALF OF PIXEL BLOCK
          dblPropY = 1 - ((dblYRemainder - dblHalfCellY) / dblCellSizeY)
          If dblXRemainder > dblHalfCellX Then                ' ON EAST SIDE, WEST HALF OF PIXEL BLOCK
            bytQuadrant = 4                                   ' ON SOUTHEAST CORNER OF CELL, NORTHWEST CORNER OF PIXEL BLOCK
            dblPropX = 1 - ((dblXRemainder - dblHalfCellX) / dblCellSizeX)
          Else                                                ' ON WEST SIDE OF CELL, EAST HALF OF PIXEL BLOCK
            bytQuadrant = 3                                   ' ON SOUTHWEST CORNER OF CELL, NORTHEAST CORNER OF PIXEL BLOCK
            dblPropX = (dblHalfCellX + dblXRemainder) / dblCellSizeX
          End If
        End If

        Set pOrigin = New Pnt

        booIsNull = False
        Select Case bytQuadrant
          Case 1              ' NORTHEAST                =================
            If iX = lngMaxX Or iY = 0 Then
              booIsNull = True
            Else
              pOrigin.SetCoords iX, iY - 1
              pRaster.Read pOrigin, pPB
              vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
              vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
              vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
              vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
              If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
                IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                dblWestProp = (CDbl(vCellValueNW) * (1 - dblPropY)) + (CDbl(vCellValueSW) * dblPropY)
                dblEastProp = (CDbl(vCellValueNE) * (1 - dblPropY)) + (CDbl(vCellValueSE) * dblPropY)
                varInterpVal = CVar((dblWestProp * dblPropX) + (dblEastProp * (1 - dblPropX)))
              End If
            End If
          Case 2              ' NORTHWEST                =================
            If iX = 0 Or iY = 0 Then
              booIsNull = True
            Else
              pOrigin.SetCoords iX - 1, iY - 1
              pRaster.Read pOrigin, pPB
              vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
              vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
              vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
              vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
              If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
                IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                dblWestProp = (CDbl(vCellValueNW) * (1 - dblPropY)) + (CDbl(vCellValueSW) * dblPropY)
                dblEastProp = (CDbl(vCellValueNE) * (1 - dblPropY)) + (CDbl(vCellValueSE) * dblPropY)
                varInterpVal = CVar((dblWestProp * (1 - dblPropX)) + (dblEastProp * dblPropX))
              End If
            End If
          Case 3              ' SOUTHWEST                =================
            If iX = 0 Or iY = lngMaxY Then
              booIsNull = True
            Else
              pOrigin.SetCoords iX - 1, iY
              pRaster.Read pOrigin, pPB
              vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
              vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
              vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
              vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
              If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
                IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                dblWestProp = (CDbl(vCellValueNW) * dblPropY) + (CDbl(vCellValueSW) * (1 - dblPropY))
                dblEastProp = (CDbl(vCellValueNE) * dblPropY) + (CDbl(vCellValueSE) * (1 - dblPropY))
                varInterpVal = CVar((dblWestProp * (1 - dblPropX)) + (dblEastProp * dblPropX))
              End If
            End If
          Case 4              ' SOUTHEAST                =================
            If iX = lngMaxX Or iY = lngMaxY Then
              booIsNull = True
            Else
              pOrigin.SetCoords iX, iY
              pRaster.Read pOrigin, pPB
              vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
              vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
              vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
              vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
              If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
                IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
                IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                    booIsNull = True
              Else
                dblWestProp = (CDbl(vCellValueNW) * dblPropY) + (CDbl(vCellValueSW) * (1 - dblPropY))
                dblEastProp = (CDbl(vCellValueNE) * dblPropY) + (CDbl(vCellValueSE) * (1 - dblPropY))
                varInterpVal = CVar((dblWestProp * dblPropX) + (dblEastProp * (1 - dblPropX)))
              End If

            End If
        End Select

        If booIsNull Then
          varReturn(lngIndex) = Null
        Else
          varReturn(lngIndex) = varInterpVal
        End If

          "   dblPropX = " & CStr(dblPropX) & vbCrLf & _
          "   dblPropY = " & CStr(dblPropY) & vbCrLf & _
          "   Quadrant = " & CStr(bytQuadrant) & vbCrLf & _
          "   vCellValueNW = " & CStr(vCellValueNW) & vbCrLf & _
          "   vCellValueNE = " & CStr(vCellValueNE) & vbCrLf & _
          "   vCellValueSW = " & CStr(vCellValueSW) & vbCrLf & _
          "   vCellValueSE = " & CStr(vCellValueSE) & vbCrLf & _
          "   dblWestProp = " & CStr(dblWestProp) & vbCrLf & _
          "   dblEastProp = " & CStr(dblEastProp) & vbCrLf & _
          "   Interpolated Value = " & CStr(varInterpVal)
      End If
    Next lngIndex
    CellValues4CellInterp_ByNumbers = varReturn

End Function

Public Function CellValues2(pPoints As IPointCollection, pRaster As IRaster, _
        Optional lngBandIndex As Long = 0) As esriSystem.IVariantArray

    Dim pRP As IRasterProps
    Set pRP = pRaster

    Dim dblCellSizeX As Double
    Dim dblCellSizeY As Double
    dblCellSizeX = ReturnPixelWidth(pRaster)
    dblCellSizeY = ReturnPixelHeight(pRaster)

    Dim dblHalfCellX As Double
    dblHalfCellX = dblCellSizeX / 2
    Dim dblHalfCellY As Double
    dblHalfCellY = dblCellSizeY / 2

    Dim pExtent As IEnvelope
    Set pExtent = pRP.Extent
    Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
    pExtent.QueryCoords X1, Y1, X2, Y2

    Dim pPB As IPixelBlock3

    Dim dWidth As Double, dHeight As Double
    dWidth = pRP.Width
    dHeight = pRP.Height

    Dim pPnt As IPnt
    Set pPnt = New Pnt
    pPnt.SetCoords 1, 1
    Set pPB = pRaster.CreatePixelBlock(pPnt)

    Dim pOrigin As IPnt

    Dim lngIndex As Long
    Dim dblCellValue As Double
    Dim pPoint As IPoint
    Dim dx As Double, dy As Double
    Dim nX As Double, ny As Double
    Dim dblXRemainder As Double, dblYRemainder As Double
    Dim iX As Long, iY As Long
    Dim lngMaxX As Long, lngMaxY As Long

    lngMaxX = pRP.Width - 1
    lngMaxY = pRP.Height - 1

    Dim dblPropX As Double
    Dim dblPropY As Double

    Dim pOutArray As esriSystem.IVariantArray
    Set pOutArray = New esriSystem.varArray

    Set pOrigin = New Pnt

    Dim vCellValue As Variant
    Dim booIsNull As Boolean

    For lngIndex = 0 To pPoints.PointCount - 1
      Set pPoint = pPoints.Point(lngIndex)

      If pPoint.x < X1 Or pPoint.x > X2 Or pPoint.Y < Y1 Or pPoint.Y > Y2 Then
        pOutArray.Add Null
      Else

        dx = pPoint.x - X1
        dy = Y2 - pPoint.Y

        nX = dx / dblCellSizeX
        ny = dy / dblCellSizeY

        iX = Int(nX)
        iY = Int(ny)

        If (iX < 0) Then iX = 0
        If (iY < 0) Then iY = 0
        If (iX > lngMaxX) Then
          iX = lngMaxX
        End If
        If (iY > lngMaxY - 1) Then
          iY = lngMaxY - 1
        End If

        booIsNull = False

        If iX = lngMaxX Or iY = 0 Then
          booIsNull = True
        Else
          pOrigin.SetCoords iX, iY
          pRaster.Read pOrigin, pPB
          vCellValue = pPB.GetVal(lngBandIndex, 0, 0)
          If IsCellNaN(vCellValue) Or IsEmpty(vCellValue) Then
            booIsNull = True
          End If
        End If

        If booIsNull Then
          pOutArray.Add Null
        Else
          pOutArray.Add CDbl(vCellValue)
        End If

          "   dblPropX = " & CStr(dblPropX) & vbCrLf & _
          "   dblPropY = " & CStr(dblPropY) & vbCrLf & _
          "   Quadrant = " & CStr(bytQuadrant) & vbCrLf & _
          "   vCellValueNW = " & CStr(vCellValueNW) & vbCrLf & _
          "   vCellValueNE = " & CStr(vCellValueNE) & vbCrLf & _
          "   vCellValueSW = " & CStr(vCellValueSW) & vbCrLf & _
          "   vCellValueSE = " & CStr(vCellValueSE) & vbCrLf & _
          "   dblWestProp = " & CStr(dblWestProp) & vbCrLf & _
          "   dblEastProp = " & CStr(dblEastProp) & vbCrLf & _
          "   Interpolated Value = " & CStr(varInterpVal)
      End If
    Next lngIndex
    Set CellValues2 = pOutArray

End Function

Public Function BuildRadiusMask(dblRadius As Double, dblCellSize As Double) As Boolean()

  Dim lngCells As Long
  lngCells = Int(dblRadius / dblCellSize)
  Dim lngEdge As Long
  lngEdge = (lngCells * 2)    ' USED FOR ARRAY INDEX, SO DON'T ADD 1 TO THIS:  ACTUAL NUMBER OF EDGE CELLS = (lngCells*2)+1
  Dim booReturn() As Boolean
  ReDim booReturn(lngEdge, lngEdge)
  Dim dblOrig As Double
  dblOrig = lngCells * dblCellSize
  Dim lngRow As Long
  Dim lngCol As Long
  For lngCol = 0 To lngEdge
    For lngRow = 0 To lngEdge
      booReturn(lngRow, lngCol) = Sqr((dblOrig - lngRow * dblCellSize) ^ 2 + (dblOrig - lngCol * dblCellSize) ^ 2) < dblRadius
    Next lngRow
  Next lngCol
  BuildRadiusMask = booReturn

End Function

Public Function CellValue2(pPoint As IPoint, pRaster As IRaster, _
        Optional lngBandIndex As Long = 0) As Variant

    Dim pRP As IRasterProps
    Set pRP = pRaster

    Dim dblCellSizeX As Double
    Dim dblCellSizeY As Double
    dblCellSizeX = ReturnPixelWidth(pRaster)
    dblCellSizeY = ReturnPixelHeight(pRaster)

    Dim dblHalfCellX As Double
    dblHalfCellX = dblCellSizeX / 2
    Dim dblHalfCellY As Double
    dblHalfCellY = dblCellSizeY / 2

    Dim pExtent As IEnvelope
    Set pExtent = pRP.Extent
    Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
    pExtent.QueryCoords X1, Y1, X2, Y2

    Dim pPB As IPixelBlock3

    Dim dWidth As Double, dHeight As Double
    dWidth = pRP.Width
    dHeight = pRP.Height

    Dim pPnt As IPnt
    Set pPnt = New Pnt
    pPnt.SetCoords 1, 1
    Set pPB = pRaster.CreatePixelBlock(pPnt)

    Dim pOrigin As IPnt

    Dim lngIndex As Long
    Dim dblCellValue As Double
    Dim dx As Double, dy As Double
    Dim nX As Double, ny As Double
    Dim dblXRemainder As Double, dblYRemainder As Double
    Dim iX As Long, iY As Long
    Dim lngMaxX As Long, lngMaxY As Long

    lngMaxX = pRP.Width - 1
    lngMaxY = pRP.Height - 1

    Dim dblPropX As Double
    Dim dblPropY As Double

    Dim pOutArray As esriSystem.IVariantArray
    Set pOutArray = New esriSystem.varArray

    Set pOrigin = New Pnt

    Dim vCellValue As Variant
    Dim booIsNull As Boolean

    If pPoint.x < X1 Or pPoint.x > X2 Or pPoint.Y < Y1 Or pPoint.Y > Y2 Then
      CellValue2 = Null
    Else

      dx = pPoint.x - X1
      dy = Y2 - pPoint.Y

      nX = dx / dblCellSizeX
      ny = dy / dblCellSizeY

      iX = Int(nX)
      iY = Int(ny)

      If (iX < 0) Then iX = 0
      If (iY < 0) Then iY = 0
      If (iX > lngMaxX) Then
        iX = lngMaxX
      End If
      If (iY > lngMaxY - 1) Then
        iY = lngMaxY - 1
      End If

      booIsNull = False

      If iX = lngMaxX Or iY = 0 Then
        booIsNull = True
      Else
        pOrigin.SetCoords iX, iY
        pRaster.Read pOrigin, pPB
        vCellValue = pPB.GetVal(lngBandIndex, 0, 0)
        If IsCellNaN(vCellValue) Or IsEmpty(vCellValue) Then
          booIsNull = True
        End If
      End If

      If booIsNull Then
        CellValue2 = Null
      Else
        CellValue2 = CDbl(vCellValue)
      End If
    End If

  Set pRP = Nothing
  Set pExtent = Nothing
  Set pPB = Nothing
  Set pPnt = Nothing
  Set pOrigin = Nothing
  Set pOutArray = Nothing

End Function

Public Sub TestDrawRectangles()
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pGeom As IGeometry
  Set pGeom = pMxDoc.ActiveView.Extent
  Dim pRLayer As IRasterLayer
  Set pRLayer = MyGeneralOperations.ReturnLayerByName("gpw_UN_NullIsZero", pMxDoc.FocusMap)
  Set pRLayer = MyGeneralOperations.ReturnLayerByName("Flat_Area_Shifted3", pMxDoc.FocusMap)
  Dim pFLayer As IFeatureLayer
  Dim pFClass As IFeatureClass
  Dim pFeature As IFeature
  Dim pPoly As IPolygon
  Dim pSelSet As ISelectionSet
  Dim pFeatSel As IFeatureSelection
  Dim pFCursor As IFeatureCursor

  Set pFLayer = MyGeneralOperations.ReturnLayerByName("hydrobasins_world", pMxDoc.FocusMap)
  Set pFeatSel = pFLayer
  Set pSelSet = pFeatSel.SelectionSet
  Dim pEnv As IEnvelope

  pSelSet.Search Nothing, False, pFCursor
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    Set pPoly = pFeature.ShapeCopy
    If pEnv Is Nothing Then
      Set pEnv = pPoly.Envelope
    Else
      pEnv.Union pPoly.Envelope
    End If
    Set pFeature = pFCursor.NextFeature
  Loop

  Call DrawRectanglesAroundCellsInView(pEnv, pRLayer, True, True, pMxDoc)
End Sub

Public Function CellValue4CellInterp(pPoint As IPoint, pRaster As IRaster, _
    Optional lngBandIndex As Long = 0) As Variant

    Dim pRP As IRasterProps
    Set pRP = pRaster

    Dim dblCellSizeX As Double
    Dim dblCellSizeY As Double
    dblCellSizeX = ReturnPixelWidth(pRaster)
    dblCellSizeY = ReturnPixelHeight(pRaster)

    Dim dblHalfCellX As Double
    dblHalfCellX = dblCellSizeX / 2
    Dim dblHalfCellY As Double
    dblHalfCellY = dblCellSizeY / 2

    Dim pExtent As IEnvelope
    Set pExtent = pRP.Extent
    Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
    pExtent.QueryCoords X1, Y1, X2, Y2

    Dim pPB As IPixelBlock3

    Dim dWidth As Double, dHeight As Double
    dWidth = pRP.Width
    dHeight = pRP.Height

    Dim pPnt As IPnt
    Set pPnt = New Pnt
    pPnt.SetCoords 2, 2
    Set pPB = pRaster.CreatePixelBlock(pPnt)

    Dim pOrigin As IPnt

    Dim lngIndex As Long
    Dim dblCellValue As Double
    Dim dx As Double, dy As Double
    Dim nX As Double, ny As Double
    Dim dblXRemainder As Double, dblYRemainder As Double
    Dim iX As Long, iY As Long
    Dim lngMaxX As Long, lngMaxY As Long

    lngMaxX = pRP.Width - 1
    lngMaxY = pRP.Height - 1

    Dim bytQuadrant As Byte       ' 1 FOR NE, 2 FOR NW, 3 FOR SW, 4 FOR SE
    Dim varInterpVal As Variant

    Dim dblPropX As Double
    Dim dblPropY As Double

    Dim pOutArray As esriSystem.IVariantArray
    Set pOutArray = New esriSystem.varArray

    Dim vCellValueNE As Variant
    Dim vCellValueNW As Variant
    Dim vCellValueSE As Variant
    Dim vCellValueSW As Variant

    Dim booIsNull As Boolean
    Dim dblWestProp As Double
    Dim dblEastProp As Double

    If pPoint.x < X1 Or pPoint.x > X2 Or pPoint.Y < Y1 Or pPoint.Y > Y2 Then
      pOutArray.Add Null
    Else

      dx = pPoint.x - X1
      dy = Y2 - pPoint.Y

      nX = dx / dblCellSizeX
      ny = dy / dblCellSizeY

      iX = Int(nX)
      iY = Int(ny)

      If (iX < 0) Then iX = 0
      If (iY < 0) Then iY = 0
      If (iX > lngMaxX) Then
        iX = lngMaxX
      End If
      If (iY > lngMaxY - 1) Then
        iY = lngMaxY - 1
      End If

      dblXRemainder = (nX - iX) * dblCellSizeX
      dblYRemainder = (ny - iY) * dblCellSizeY

      If dblYRemainder < dblHalfCellY Then                  ' ON NORTH SIDE OF CELL, SOUTH HALF OF PIXEL BLOCK
        dblPropY = (dblYRemainder + dblHalfCellY) / dblCellSizeY
        If dblXRemainder > dblHalfCellX Then                ' ON EAST SIDE OF CELL, WEST HALF OF PIXEL BLOCK
          bytQuadrant = 1                                   ' ON NORTHEAST CORNER OF CELL, SOUTHWEST CORNER OF PIXEL BLOCK
          dblPropX = 1 - ((dblXRemainder - dblHalfCellX) / dblCellSizeX)
        Else                                                ' ON WEST SIDE OF CELL, EAST HALF OF PIXEL BLOCK
          bytQuadrant = 2                                   ' ON NORTHWEST CORNER OF CELL, SOUTHEAST CORNER OF PIXEL BLOCK
          dblPropX = (dblHalfCellX + dblXRemainder) / dblCellSizeX
        End If
      Else                                                  ' ON SOUTH SIDE OF CELL, NORTH HALF OF PIXEL BLOCK
        dblPropY = 1 - ((dblYRemainder - dblHalfCellY) / dblCellSizeY)
        If dblXRemainder > dblHalfCellX Then                ' ON EAST SIDE, WEST HALF OF PIXEL BLOCK
          bytQuadrant = 4                                   ' ON SOUTHEAST CORNER OF CELL, NORTHWEST CORNER OF PIXEL BLOCK
          dblPropX = 1 - ((dblXRemainder - dblHalfCellX) / dblCellSizeX)
        Else                                                ' ON WEST SIDE OF CELL, EAST HALF OF PIXEL BLOCK
          bytQuadrant = 3                                   ' ON SOUTHWEST CORNER OF CELL, NORTHEAST CORNER OF PIXEL BLOCK
          dblPropX = (dblHalfCellX + dblXRemainder) / dblCellSizeX
        End If
      End If

      Set pOrigin = New Pnt

      booIsNull = False
      Select Case bytQuadrant
        Case 1              ' NORTHEAST                =================
          If iX = lngMaxX Or iY = 0 Then
            booIsNull = True
          Else
            pOrigin.SetCoords iX, iY - 1
            pRaster.Read pOrigin, pPB
            vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
            vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
            vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
            vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
            If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
              IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
              IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                  booIsNull = True
            Else
              dblWestProp = (CDbl(vCellValueNW) * (1 - dblPropY)) + (CDbl(vCellValueSW) * dblPropY)
              dblEastProp = (CDbl(vCellValueNE) * (1 - dblPropY)) + (CDbl(vCellValueSE) * dblPropY)
              varInterpVal = CVar((dblWestProp * dblPropX) + (dblEastProp * (1 - dblPropX)))
            End If
          End If
        Case 2              ' NORTHWEST                =================
          If iX = 0 Or iY = 0 Then
            booIsNull = True
          Else
            pOrigin.SetCoords iX - 1, iY - 1
            pRaster.Read pOrigin, pPB
            vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
            vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
            vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
            vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
            If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
              IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
              IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                  booIsNull = True
            Else
              dblWestProp = (CDbl(vCellValueNW) * (1 - dblPropY)) + (CDbl(vCellValueSW) * dblPropY)
              dblEastProp = (CDbl(vCellValueNE) * (1 - dblPropY)) + (CDbl(vCellValueSE) * dblPropY)
              varInterpVal = CVar((dblWestProp * (1 - dblPropX)) + (dblEastProp * dblPropX))
            End If
          End If
        Case 3              ' SOUTHWEST                =================
          If iX = 0 Or iY = lngMaxY Then
            booIsNull = True
          Else
            pOrigin.SetCoords iX - 1, iY
            pRaster.Read pOrigin, pPB
            vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
            vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
            vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
            vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
            If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
              IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
              IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                  booIsNull = True
            Else
              dblWestProp = (CDbl(vCellValueNW) * dblPropY) + (CDbl(vCellValueSW) * (1 - dblPropY))
              dblEastProp = (CDbl(vCellValueNE) * dblPropY) + (CDbl(vCellValueSE) * (1 - dblPropY))
              varInterpVal = CVar((dblWestProp * (1 - dblPropX)) + (dblEastProp * dblPropX))
            End If
          End If
        Case 4              ' SOUTHEAST                =================
          If iX = lngMaxX Or iY = lngMaxY Then
            booIsNull = True
          Else
            pOrigin.SetCoords iX, iY
            pRaster.Read pOrigin, pPB
            vCellValueNW = pPB.GetVal(lngBandIndex, 0, 0)
            vCellValueSW = pPB.GetVal(lngBandIndex, 0, 1)
            vCellValueNE = pPB.GetVal(lngBandIndex, 1, 0)
            vCellValueSE = pPB.GetVal(lngBandIndex, 1, 1)
            If IsCellNaN(vCellValueNW) Or IsCellNaN(vCellValueNE) Or IsCellNaN(vCellValueSW) Or _
              IsCellNaN(vCellValueSE) Or IsEmpty(vCellValueNW) Or IsEmpty(vCellValueNE) Or _
              IsEmpty(vCellValueSW) Or IsEmpty(vCellValueSE) Then
                  booIsNull = True
            Else
              dblWestProp = (CDbl(vCellValueNW) * dblPropY) + (CDbl(vCellValueSW) * (1 - dblPropY))
              dblEastProp = (CDbl(vCellValueNE) * dblPropY) + (CDbl(vCellValueSE) * (1 - dblPropY))
              varInterpVal = CVar((dblWestProp * dblPropX) + (dblEastProp * (1 - dblPropX)))
            End If

          End If
      End Select

      If booIsNull Then
        CellValue4CellInterp = Null
      Else
        CellValue4CellInterp = CDbl(varInterpVal)
      End If

        "   dblPropX = " & CStr(dblPropX) & vbCrLf & _
        "   dblPropY = " & CStr(dblPropY) & vbCrLf & _
        "   Quadrant = " & CStr(bytQuadrant) & vbCrLf & _
        "   vCellValueNW = " & CStr(vCellValueNW) & vbCrLf & _
        "   vCellValueNE = " & CStr(vCellValueNE) & vbCrLf & _
        "   vCellValueSW = " & CStr(vCellValueSW) & vbCrLf & _
        "   vCellValueSE = " & CStr(vCellValueSE) & vbCrLf & _
        "   dblWestProp = " & CStr(dblWestProp) & vbCrLf & _
        "   dblEastProp = " & CStr(dblEastProp) & vbCrLf & _
        "   Interpolated Value = " & CStr(varInterpVal)
    End If

ClearMemory:
  Set pRP = Nothing
  Set pExtent = Nothing
  Set pPB = Nothing
  Set pPnt = Nothing
  Set pOrigin = Nothing
  varInterpVal = Null
  Set pOutArray = Nothing
  vCellValueNE = Null
  vCellValueNW = Null
  vCellValueSE = Null
  vCellValueSW = Null

End Function

Public Function ReturnBooleanArrayCellsInPolygon(pPolygon As IPolygon, _
    pRaster As IRaster, lngCellOriginX As Long, lngCellOriginY As Long) As Boolean()

  Dim pSpRef As ISpatialReference
  Dim pGeoDataset As IGeoDataset
  Set pGeoDataset = pRaster
  Set pSpRef = pPolygon.SpatialReference
  Dim pClone As IClone
  Dim pTestPolygon As IPolygon

  If Not MyGeneralOperations.CompareSpatialReferences(pSpRef, pGeoDataset.SpatialReference) Then
    pTestPolygon.Project pGeoDataset.SpatialReference
    Set pClone = pPolygon
    Set pTestPolygon = pClone.Clone
  Else
    Set pTestPolygon = pPolygon
  End If

  Dim dblPolyXMin As Double
  Dim dblPolyYMin As Double
  Dim dblPolyXMax As Double
  Dim dblPolyYMax As Double
  Dim dblCellX As Double
  Dim dblCellY As Double
  Dim pPolyEnv As IEnvelope
  Dim pRastEnv As IEnvelope

  Set pPolyEnv = pTestPolygon.Envelope
  Set pRastEnv = pGeoDataset.Extent

  dblPolyXMin = pPolyEnv.XMin
  dblPolyYMin = pPolyEnv.YMin
  dblPolyXMax = pPolyEnv.XMax
  dblPolyYMax = pPolyEnv.YMax

  dblCellX = GridFunctions.ReturnPixelWidth(pRaster)
  dblCellY = GridFunctions.ReturnPixelHeight(pRaster)

  Dim pPtColl As IPointCollection
  Dim pGeom As IGeometry
  Dim pRelOp As IRelationalOperator
  Dim pClipPoints As IPointCollection
  Dim pPoint As IPoint
  Dim dblRasterXOrigin As Double
  Dim dblRasterYOrigin As Double
  dblRasterXOrigin = pRastEnv.XMin + (dblCellX / 2)
  dblRasterYOrigin = pRastEnv.YMax - (dblCellY / 2)

  Dim lngX As Long
  Dim lngY As Long
  Dim dblStartX As Double
  Dim dblStartY As Double
  Dim dblEndX As Double
  Dim dblEndY As Double
  Dim lngWidthCells As Long
  Dim lngHeightCells As Long

  Dim dblXCoord As Double
  Dim dblYCoord As Double

  dblStartX = (dblPolyXMin - MyGeometricOperations.ModDouble(dblPolyXMin - pRastEnv.XMin, dblCellX)) + dblCellX / 2
  dblStartY = (dblPolyYMin - MyGeometricOperations.ModDouble(dblPolyYMin - pRastEnv.YMin, dblCellY)) + dblCellY / 2
  dblEndX = (dblPolyXMax - MyGeometricOperations.ModDouble(dblPolyXMax - pRastEnv.XMin, dblCellX)) + dblCellX / 2
  dblEndY = (dblPolyYMax - MyGeometricOperations.ModDouble(dblPolyYMax - pRastEnv.YMin, dblCellY)) + dblCellY / 2
  lngWidthCells = Round((dblEndX - dblStartX) / dblCellX)  ' THIS IS ACTUALLY WIDTH - 1
  lngHeightCells = Round((dblEndY - dblStartY) / dblCellY)  ' THIS IS ACTUALLY HEIGHT - 1

  Dim pRaster2 As IRaster2
  Set pRaster2 = pRaster
  lngCellOriginY = Round((dblRasterYOrigin - dblEndY) / dblCellY)
  lngCellOriginX = Round((dblStartX - dblRasterXOrigin) / dblCellX)
  Dim booReturn() As Boolean
  ReDim booReturn(lngWidthCells, lngHeightCells)

  Set pRelOp = pTestPolygon
  Set pPoint = New Point
  Set pPoint.SpatialReference = pTestPolygon.SpatialReference

  For lngX = 0 To lngWidthCells
    For lngY = 0 To lngHeightCells

      pPoint.PutCoords dblStartX + (CDbl(lngX) * dblCellX), dblStartY + (CDbl(lngY) * dblCellY)

      booReturn(lngX, lngHeightCells - lngY) = Not pRelOp.Disjoint(pPoint)

    Next lngY
  Next lngX

  TrimExtraneousEdges booReturn, lngCellOriginX, lngCellOriginY, lngWidthCells, lngHeightCells

  ReturnBooleanArrayCellsInPolygon = booReturn

  GoTo ClearMemory

ClearMemory:
  Set pSpRef = Nothing
  Set pGeoDataset = Nothing
  Set pClone = Nothing
  Set pTestPolygon = Nothing
  Set pPolyEnv = Nothing
  Set pRastEnv = Nothing
  Set pPtColl = Nothing
  Set pGeom = Nothing
  Set pRelOp = Nothing
  Set pClipPoints = Nothing
  Set pPoint = Nothing
  Set pRaster2 = Nothing
  Erase booReturn

End Function

Public Sub TrimExtraneousEdges(booReturn() As Boolean, lngCellOriginX As Long, _
    lngCellOriginY As Long, lngWidthCells As Long, lngHeightCells As Long)

  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim booTrimLeft As Boolean
  Dim booTrimRight As Boolean
  Dim booTrimTop As Boolean
  Dim booTrimBottom As Boolean

  If UBound(booReturn, 2) > 0 Then
    booTrimLeft = True
    booTrimRight = True
    For lngIndex2 = 0 To UBound(booReturn, 2)
      If booReturn(0, lngIndex2) Then booTrimLeft = False
      If booReturn(UBound(booReturn), lngIndex2) Then booTrimRight = False
    Next lngIndex2
  Else
    booTrimLeft = False
    booTrimRight = False
  End If

  If UBound(booReturn, 1) > 0 Then
    booTrimTop = True
    booTrimBottom = True
    For lngIndex1 = 0 To UBound(booReturn, 1)
      If booReturn(lngIndex1, 0) Then booTrimTop = False
      If booReturn(lngIndex1, UBound(booReturn, 2)) Then booTrimBottom = False
    Next lngIndex1
  Else
    booTrimTop = False
    booTrimBottom = False
  End If

  Dim booNew() As Boolean

  If booTrimBottom And UBound(booReturn, 2) > 0 Then
    lngHeightCells = lngHeightCells - 1
    ReDim booNew(UBound(booReturn, 1), UBound(booReturn, 2) - 1)
    For lngIndex1 = 0 To UBound(booReturn, 1)
      For lngIndex2 = 0 To UBound(booReturn, 2) - 1
        booNew(lngIndex1, lngIndex2) = booReturn(lngIndex1, lngIndex2)
      Next lngIndex2
    Next lngIndex1
    booReturn = booNew
  End If

  If booTrimRight And UBound(booReturn, 1) > 0 Then
    lngWidthCells = lngWidthCells - 1
    ReDim booNew(UBound(booReturn, 1) - 1, UBound(booReturn, 2))
    For lngIndex1 = 0 To UBound(booReturn, 1) - 1
      For lngIndex2 = 0 To UBound(booReturn, 2)
        booNew(lngIndex1, lngIndex2) = booReturn(lngIndex1, lngIndex2)
      Next lngIndex2
    Next lngIndex1
    booReturn = booNew
  End If

  If booTrimTop And UBound(booReturn, 2) > 0 Then
    lngHeightCells = lngHeightCells - 1
    lngCellOriginY = lngCellOriginY + 1
    ReDim booNew(UBound(booReturn, 1), UBound(booReturn, 2) - 1)
    For lngIndex1 = 0 To UBound(booReturn, 1)
      For lngIndex2 = 1 To UBound(booReturn, 2)
        booNew(lngIndex1, lngIndex2 - 1) = booReturn(lngIndex1, lngIndex2)
      Next lngIndex2
    Next lngIndex1
    booReturn = booNew
  End If

  If booTrimLeft And UBound(booReturn, 1) > 0 Then
    lngWidthCells = lngWidthCells - 1
    lngCellOriginX = lngCellOriginX + 1
     ReDim booNew(UBound(booReturn, 1) - 1, UBound(booReturn, 2))
    For lngIndex1 = 1 To UBound(booReturn, 1)
      For lngIndex2 = 0 To UBound(booReturn, 2)
        booNew(lngIndex1 - 1, lngIndex2) = booReturn(lngIndex1, lngIndex2)
      Next lngIndex2
    Next lngIndex1
    booReturn = booNew
  End If

  GoTo ClearMemory

ClearMemory:
  Erase booNew

End Sub

Public Function ConvertVarArrayToArrayOfValues(varValues() As Variant, dblNumbers() As Double) As Boolean
  On Error GoTo ErrHandler

  ReDim dblNumbers((UBound(varValues, 1) + 1) * (UBound(varValues, 2) + 1) - 1)
  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim lngCounter As Long
  lngCounter = -1
  Dim varVal As Variant

  For lngIndex1 = 0 To UBound(varValues, 1)
    For lngIndex2 = 0 To UBound(varValues, 2)
      varVal = varValues(lngIndex1, lngIndex2)
      If Not IsNull(varVal) Then
        lngCounter = lngCounter + 1
        dblNumbers(lngCounter) = CDbl(varVal)
      End If
    Next lngIndex2
  Next lngIndex1

  If lngCounter = -1 Then
    ConvertVarArrayToArrayOfValues = False
  Else
    If lngCounter < UBound(dblNumbers) Then ReDim Preserve dblNumbers(lngCounter)
    ConvertVarArrayToArrayOfValues = True
  End If

  Exit Function

ErrHandler:
  ConvertVarArrayToArrayOfValues = False

End Function

Public Sub FillRasterParams(pRaster As IRaster, dblCellSizeX As Double, dblCellSizeY As Double, _
        X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, lngMaxX As Long, lngMaxY As Long)

    Dim pRP As IRasterProps
    Set pRP = pRaster
    dblCellSizeX = ReturnPixelWidth(pRaster)
    dblCellSizeY = ReturnPixelHeight(pRaster)

    Dim pExtent As IEnvelope
    Set pExtent = pRP.Extent
    pExtent.QueryCoords X1, Y1, X2, Y2

    lngMaxX = pRP.Width - 1
    lngMaxY = pRP.Height - 1

End Sub

Public Function CellValues2_Fast_byArray(booArray() As Boolean, pRaster As IRaster, _
        dblCellSizeX As Double, dblCellSizeY As Double, X1 As Double, Y1 As Double, _
        X2 As Double, Y2 As Double, pPB As IPixelBlock3, _
        lngMaxX As Long, lngMaxY As Long, pPnt As IPnt, pOrigin As IPnt, lngMaxIndex As Long, _
        Optional lngBandIndex As Long = 0) As Double()

    Dim lngIndex As Long
    Dim lngIndex2 As Long
    Dim vCellValue As Variant

    Dim dblReturn() As Double
    lngMaxIndex = -1

    pRaster.Read pOrigin, pPB
    Dim varPixels As Variant
    pPB.PixelType(0) = PT_FLOAT
    varPixels = pPB.PixelData(0)

    For lngIndex = 0 To UBound(booArray, 1)
      For lngIndex2 = 0 To UBound(booArray, 2)
        If booArray(lngIndex, lngIndex2) Then
          vCellValue = varPixels(lngIndex, lngIndex2)
          If Not IsNull(vCellValue) Then
            lngMaxIndex = lngMaxIndex + 1
            ReDim Preserve dblReturn(lngMaxIndex)
            dblReturn(lngMaxIndex) = CDbl(vCellValue)
          End If
        End If
      Next lngIndex2
    Next lngIndex

    CellValues2_Fast_byArray = dblReturn

End Function


