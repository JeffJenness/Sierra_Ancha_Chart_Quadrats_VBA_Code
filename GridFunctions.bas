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

Public Function ReturnCellCount(pRaster As IRaster) As Long

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

Public Function CalcGridLine(ByVal pStartPolygon As IPolygon, ByVal pEndPolygon As IPolygon, _
      ByVal pCorPolygon As IPolygon, pCorRaster As IRaster, pEnv As IRasterAnalysisEnvironment, _
      Optional ShouldClean As Boolean) As IPolyline

  Dim pClone As IClone

  Dim pEnvelope As IEnvelope
  pEnv.GetExtent esriRasterEnvValue, pEnvelope

  Dim pRasMakerOp As IRasterMakerOp
  Set pRasMakerOp = New RasterMakerOp
  GridFunctions.SetSpatialAnalysisSettings pRasMakerOp, pEnv

  Dim pSpRef As ISpatialReference
  Set pSpRef = pEnv.OutSpatialReference

  If Not (GridFunctions.CompareSpatialReferences(pStartPolygon.SpatialReference, pSpRef)) Then
    pStartPolygon.Project pSpRef
  End If
  If Not (GridFunctions.CompareSpatialReferences(pEndPolygon.SpatialReference, pSpRef)) Then
    pEndPolygon.Project pSpRef
  End If

  Dim pPoints As IPointCollection
  Set pPoints = GridFunctions.ReturnPointsByCellSize(pCorRaster, pEndPolygon)
  Dim anIndex As Long

  Dim pDistanceOp As IDistanceOp
  Set pDistanceOp = New RasterDistanceOp

  SetSpatialAnalysisSettings pDistanceOp, pEnv

  Dim pBaseRaster As IRaster
  Set pBaseRaster = pRasMakerOp.MakeConstant(1, True)
  Dim pSourceDataset As IRaster
  Set pSourceDataset = ClipRasterToPolygon(pBaseRaster, pStartPolygon, True, , , pEnv)

  Dim pCostDataset As IGeoDataset

  Dim pOutputRaster As IGeoDataset

  Set pOutputRaster = pDistanceOp.CostDistanceFull(pSourceDataset, pCorRaster, True, True, False)

  Dim pRasterBandCollection As IRasterBandCollection
  Set pRasterBandCollection = pOutputRaster
  Dim pDistBand As IRasterBand        ' DISTANCE BAND
  Set pDistBand = pRasterBandCollection.Item(0)
  Dim pDistRaster As IRasterBandCollection
  Set pDistRaster = New Raster
  pDistRaster.Add pDistBand, 0

  Dim pDistAsRaster As IRaster
  Set pDistAsRaster = pDistRaster

  Dim pBacklinkBand As IRasterBand    ' BACKLINK BAND
  Set pBacklinkBand = pRasterBandCollection.Item(1)
  Dim pBacklinkRaster As IRasterBandCollection
  Set pBacklinkRaster = New Raster
  pBacklinkRaster.Add pBacklinkBand, 0

  Dim pPathCollection As IGeometryCollection
  Set pPathCollection = pDistanceOp.CostPathAsPolyline(pPoints, pDistAsRaster, pBacklinkRaster)
  If pPathCollection.GeometryCount = 0 Then
    Dim pEmptyLine As IPolyline
    Set pEmptyLine = New Polyline
    pEmptyLine.SetEmpty
    Set CalcGridLine = pEmptyLine
    Exit Function
  End If

  Dim pPath As IPolyline
  Dim pMinPath As IPolyline
  Dim dblShortDist As Double
  Set pPath = pPathCollection.Geometry(0)
  Set pMinPath = pPath
  dblShortDist = pPath.length
  If pPathCollection.GeometryCount > 1 Then
    For anIndex = 1 To pPathCollection.GeometryCount - 1
      Set pPath = pPathCollection.Geometry(anIndex)
      If pPath.length < dblShortDist Then
        Set pMinPath = pPath
        dblShortDist = pPath.length
      End If
    Next anIndex
  End If

  Dim pBoundary As IPolyline
  Dim pTopoOp As ITopologicalOperator
  Set pTopoOp = pCorPolygon
  Set pBoundary = pTopoOp.Boundary

  Dim pCleanLine As IPolyline
  If (ShouldClean) Then
    Set pCleanLine = CleanPolyline(pMinPath, pBoundary)
  Else
    Set pCleanLine = pMinPath
  End If

  Set pDistAsRaster = Nothing
  Set pBacklinkRaster = Nothing
  pCleanLine.ReverseOrientation
  Set pCleanLine.SpatialReference = pEnv.OutSpatialReference
  Set CalcGridLine = pCleanLine

  Set pClone = Nothing
  Set pEnvelope = Nothing
  Set pRasMakerOp = Nothing
  Set pSpRef = Nothing
  Set pPoints = Nothing
  Set pDistanceOp = Nothing
  Set pBaseRaster = Nothing
  Set pSourceDataset = Nothing
  Set pCostDataset = Nothing
  Set pOutputRaster = Nothing
  Set pRasterBandCollection = Nothing
  Set pDistBand = Nothing
  Set pDistRaster = Nothing
  Set pDistAsRaster = Nothing
  Set pBacklinkBand = Nothing
  Set pBacklinkRaster = Nothing
  Set pPathCollection = Nothing
  Set pEmptyLine = Nothing
  Set pPath = Nothing
  Set pMinPath = Nothing
  Set pBoundary = Nothing
  Set pTopoOp = Nothing
  Set pCleanLine = Nothing

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

Public Sub DrawRectanglesAroundCellsInView(pGeom As IGeometry, pRasterLayer As IRasterLayer, _
    booDrawCenterpoints As Boolean, booDrawBoxes As Boolean, pMxDoc As IMxDocument)

  Dim strRasterLayerName As String
  strRasterLayerName = pRasterLayer.Name

  Dim pEnv As IEnvelope
  Set pEnv = pGeom.Envelope
  Dim pOrigin As IPoint
  Dim dblXMin As Double
  Dim dblXMax As Double
  Dim dblYMin As Double
  Dim dblYMax As Double
  dblXMin = pEnv.XMin
  dblXMax = pEnv.XMax
  dblYMin = pEnv.YMin
  dblYMax = pEnv.YMax

  Dim dblCellWidth As Double
  Dim dblCellHeight As Double
  Dim pPnt As IPnt
  Dim pRaster As IRaster
  Dim pRastBand As IRasterBand
  Dim pRastBandColl As IRasterBandCollection
  Dim pRastProps As IRasterProps

  Set pRaster = pRasterLayer.Raster
  Set pRastBandColl = pRaster
  Set pRastBand = pRastBandColl.Item(0)
  Set pRastProps = pRastBand
  Set pPnt = pRastProps.MeanCellSize
  dblCellWidth = pPnt.x
  dblCellHeight = pPnt.Y

  Dim dblRastXMin As Double
  Dim dblRastYMin As Double
  dblRastXMin = pRastProps.Extent.XMin
  dblRastYMin = pRastProps.Extent.YMin

  Dim dblShiftX As Double
  Dim dblShiftY  As Double
  dblShiftX = MyGeometricOperations.ModDouble(dblXMin - dblRastXMin, dblCellWidth)
  dblShiftY = MyGeometricOperations.ModDouble(dblYMin - dblRastYMin, dblCellHeight)

  Dim dblXIndex As Double
  Dim dblYIndex As Double
  dblXIndex = dblXMin - dblShiftX
  dblYIndex = dblYMin - dblShiftY

  Dim pPoly As IPolygon
  Dim pSpRef As ISpatialReference
  Set pSpRef = pEnv.SpatialReference

  Dim dblValue As Double
  Dim pValArray As esriSystem.IVariantArray
  Set pValArray = New esriSystem.varArray
  Dim pValSubArray As esriSystem.IVariantArray
  Dim pFieldsArray As esriSystem.IVariantArray
  Set pFieldsArray = New esriSystem.varArray
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Value"
    .Type = esriFieldTypeDouble
  End With
  pFieldsArray.Add pField

  Dim pPolyArray As esriSystem.IArray
  Set pPolyArray = New esriSystem.Array

  Dim pClone As IClone
  Dim pPtValArray As esriSystem.IVariantArray
  Set pPtValArray = New esriSystem.varArray
  Dim pPtValSubArray As esriSystem.IVariantArray
  Dim pPtArray As esriSystem.IArray
  Set pPtArray = New esriSystem.Array
  Dim pPtFieldsArray As esriSystem.IVariantArray
  Set pPtFieldsArray = New esriSystem.varArray
  Dim varVal As Variant

  Set pClone = pField
  pPtFieldsArray.Add pClone.Clone

  Dim pTempPoint As IPoint

  Do While dblXIndex < dblXMax
    Do While dblYIndex < dblYMax

      Dim pPtColl As IPointCollection
      Set pPoly = New Polygon
      Set pPoly.SpatialReference = pSpRef
      Set pPtColl = pPoly

      Set pTempPoint = New Point
      Set pTempPoint.SpatialReference = pSpRef
      pTempPoint.PutCoords dblXIndex, dblYIndex
      pPtColl.AddPoint pTempPoint

      Set pTempPoint = New Point
      Set pTempPoint.SpatialReference = pSpRef
      pTempPoint.PutCoords dblXIndex, dblYIndex + dblCellHeight
      pPtColl.AddPoint pTempPoint

      Set pTempPoint = New Point
      Set pTempPoint.SpatialReference = pSpRef
      pTempPoint.PutCoords dblXIndex + dblCellWidth, dblYIndex + dblCellHeight
      pPtColl.AddPoint pTempPoint

      Set pTempPoint = New Point
      Set pTempPoint.SpatialReference = pSpRef
      pTempPoint.PutCoords dblXIndex + dblCellWidth, dblYIndex
      pPtColl.AddPoint pTempPoint

      pPoly.Close
      pPolyArray.Add pPoly

      Set pTempPoint = New Point
      Set pTempPoint.SpatialReference = pSpRef
      pTempPoint.PutCoords dblXIndex + (dblCellWidth / 2), dblYIndex + (dblCellHeight / 2)
      pPtArray.Add pTempPoint

      varVal = GridFunctions.CellValue2(pTempPoint, pRaster)

      Set pPtValSubArray = New esriSystem.varArray
      Set pValSubArray = New esriSystem.varArray
      If Not IsNull(varVal) Then
        dblValue = CDbl(varVal)
        pValSubArray.Add dblValue
        pPtValSubArray.Add dblValue
      Else
        pValSubArray.Add Null
        pPtValSubArray.Add Null
      End If

      pValArray.Add pValSubArray
      pPtValArray.Add pPtValSubArray

      dblYIndex = dblYIndex + dblCellHeight

    Loop
    dblYIndex = dblYMin - dblShiftY
    dblXIndex = dblXIndex + dblCellWidth
  Loop

  Dim pNewFlayer As IFeatureLayer
  Dim pNewFClass As IFeatureClass
  If booDrawBoxes Then
    Set pNewFClass = MyGeneralOperations.CreateInMemoryFeatureClass2(pPolyArray, pValArray, pFieldsArray)
    Set pNewFlayer = New FeatureLayer
    Set pNewFlayer.FeatureClass = pNewFClass

    pNewFlayer.Name = strRasterLayerName & "_Cells"
    pMxDoc.FocusMap.AddLayer pNewFlayer
  End If

  Dim pNewPtFLayer As IFeatureLayer
  Dim pNewPtFClass As IFeatureClass
  If booDrawCenterpoints Then
    Set pNewPtFClass = MyGeneralOperations.CreateInMemoryFeatureClass2(pPtArray, pPtValArray, pPtFieldsArray)
    Set pNewPtFLayer = New FeatureLayer
    Set pNewPtFLayer.FeatureClass = pNewPtFClass

    pNewPtFLayer.Name = strRasterLayerName & "_CenterPoints"
    pMxDoc.FocusMap.AddLayer pNewPtFLayer
  End If

  pMxDoc.UpdateContents
  pMxDoc.ActiveView.Refresh

  GoTo ClearMemory
ClearMemory:
  Set pEnv = Nothing
  Set pOrigin = Nothing
  Set pPnt = Nothing
  Set pRaster = Nothing
  Set pRasterLayer = Nothing
  Set pRastBand = Nothing
  Set pRastBandColl = Nothing
  Set pRastProps = Nothing
  Set pPoly = Nothing
  Set pSpRef = Nothing
  Set pValArray = Nothing
  Set pValSubArray = Nothing
  Set pFieldsArray = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pPolyArray = Nothing
  Set pClone = Nothing
  Set pPtValArray = Nothing
  Set pPtValSubArray = Nothing
  Set pPtArray = Nothing
  Set pPtFieldsArray = Nothing
  Set pTempPoint = Nothing
  Set pPtColl = Nothing
  Set pNewFlayer = Nothing
  Set pNewFClass = Nothing
  Set pNewPtFLayer = Nothing
  Set pNewPtFClass = Nothing

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

Public Function ReturnPointsDistributedInPolygon(pPolygon As IPolygon, pRaster As IRaster) As IPointCollection

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
  Dim pTopoOp As ITopologicalOperator
  Dim pClipPoints As IPointCollection
  Dim pPoint As IPoint

  Dim lngX As Long
  Dim lngY As Long
  Dim dblStartX As Double
  Dim dblStartY As Double
  Dim dblEndX As Double
  Dim dblEndY As Double
  Dim lngWidthCells As Double
  Dim lngHeightCells As Long

  Dim dblXCoord As Double
  Dim dblYCoord As Double

  dblStartX = (dblPolyXMin - MyGeometricOperations.ModDouble(dblPolyXMin - pRastEnv.XMin, dblCellX)) + dblCellX / 2
  dblStartY = (dblPolyYMin - MyGeometricOperations.ModDouble(dblPolyYMin - pRastEnv.YMin, dblCellY)) + dblCellY / 2
  dblEndX = (dblPolyXMax - MyGeometricOperations.ModDouble(dblPolyXMax - pRastEnv.XMin, dblCellX)) + dblCellX / 2
  dblEndY = (dblPolyYMax - MyGeometricOperations.ModDouble(dblPolyYMax - pRastEnv.YMin, dblCellY)) + dblCellY / 2
  lngWidthCells = Round((dblEndX - dblStartX) / dblCellX)
  lngHeightCells = Round((dblEndY - dblStartY) / dblCellY)

  Set pPtColl = New Multipoint
  Set pGeom = pPtColl
  Set pGeom.SpatialReference = pSpRef

  For lngX = 0 To lngWidthCells
    For lngY = 0 To lngHeightCells
      Set pPoint = New Point
      pPoint.PutCoords dblStartX + (CDbl(lngX) * dblCellX), dblStartY + (CDbl(lngY) * dblCellY)
      pPtColl.AddPoint pPoint
    Next lngY
  Next lngX

  Set pTopoOp = pTestPolygon
  Set pClipPoints = pTopoOp.Intersect(pPtColl, esriGeometry0Dimension)

  Set ReturnPointsDistributedInPolygon = pClipPoints

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
  Set pTopoOp = Nothing
  Set pClipPoints = Nothing
  Set pPoint = Nothing

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

Public Function CellValues4(booShouldReturn() As Boolean, lngCellOriginX As Long, _
    lngCellOriginY As Long, pRaster As IRaster) As Variant()

    Dim varReturn() As Variant
    ReDim varReturn(UBound(booShouldReturn, 1), UBound(booShouldReturn, 2))

    Dim lngIndexX As Long
    Dim lngIndexY As Long

    Dim pRastProps As IRasterProps
    Set pRastProps = pRaster
    Dim varNoValue As Variant
    varNoValue = pRastProps.NoDataValue

    Dim pPB As IPixelBlock3
    Dim lngWidth As Long
    Dim lngHeight As Long

    lngWidth = UBound(booShouldReturn, 1)
    lngHeight = UBound(booShouldReturn, 2)

    Dim pPnt As IPnt
    Set pPnt = New Pnt
    pPnt.SetCoords lngWidth + 1, lngHeight + 1

    Dim pOrigin As IPnt
    Set pOrigin = New Pnt
    pOrigin.SetCoords lngCellOriginX, lngCellOriginY

    Set pPB = pRaster.CreatePixelBlock(pPnt)
    pPB.PixelType(0) = PT_DOUBLE
    pRaster.Read pOrigin, pPB

    Dim dblCellValue As Double
    Dim varData As Variant
    Dim varCellVal As Variant

    varData = pPB.PixelData(0)

    If IsEmpty(pRastProps.NoDataValue) Then
      For lngIndexY = 0 To lngHeight
        For lngIndexX = 0 To lngWidth
          If booShouldReturn(lngIndexX, lngIndexY) Then
            varCellVal = CVar(varData(lngIndexX, lngIndexY))
            varReturn(lngIndexX, lngIndexY) = varCellVal
          Else
            varReturn(lngIndexX, lngIndexY) = Null
          End If
        Next lngIndexX
      Next lngIndexY
    Else
      For lngIndexY = 0 To lngHeight
        For lngIndexX = 0 To lngWidth
          If booShouldReturn(lngIndexX, lngIndexY) Then
            varCellVal = CVar(varData(lngIndexX, lngIndexY))
            If varCellVal = varNoValue(0) Then
              varReturn(lngIndexX, lngIndexY) = Null
            Else
              varReturn(lngIndexX, lngIndexY) = varCellVal
            End If
          Else
            varReturn(lngIndexX, lngIndexY) = Null
          End If
        Next lngIndexX
      Next lngIndexY

    End If

    CellValues4 = varReturn

  GoTo ClearMemory
ClearMemory:
  Erase varReturn
  Set pRastProps = Nothing
  varNoValue = Null
  Set pPB = Nothing
  Set pPnt = Nothing
  Set pOrigin = Nothing
  varData = Null
  varCellVal = Null

End Function

Public Function CellValues2_Fast_byArray_VectorAdjust(booArray() As Boolean, pRaster As IRaster, _
        dblCellSizeX As Double, dblCellSizeY As Double, X1 As Double, Y1 As Double, _
        X2 As Double, Y2 As Double, pPB As IPixelBlock3, _
        lngMaxX As Long, lngMaxY As Long, pPnt As IPnt, pOrigin As IPnt, lngMaxIndex As Long, _
        pAOIPolygon As IPolygon, varEnvelopes() As Variant, pFClass As IFeatureClass, _
        Optional lngBandIndex As Long = 0, Optional pMxDoc As IMxDocument, Optional booPause As Boolean, _
        Optional booForceFullArea As Boolean, Optional booUseDifferentArea As Boolean, _
        Optional pTileFClass As IFeatureClass) As Double()

  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim vCellValue As Variant

  Dim dblFullArea As Double
  Dim dblClipArea As Double
  Dim dblPolygonSumArea As Double
  Dim dblProportion As Double
  Dim pEnv As IEnvelope
  Dim pClipPoly As IPolygon
  Dim varReturn As Variant

  Dim dblReturn() As Double
  lngMaxIndex = -1

  pRaster.Read pOrigin, pPB
  Dim varPixels As Variant
  pPB.PixelType(0) = PT_FLOAT
  varPixels = pPB.PixelData(0)

  For lngIndex = 0 To UBound(booArray, 1)
    For lngIndex2 = 0 To UBound(booArray, 2)
      If booArray(lngIndex, lngIndex2) Then

        If lngIndex = 2 And lngIndex2 = 5 Then
          DoEvents
        End If

        Set pEnv = varEnvelopes(lngIndex, lngIndex2)

        varReturn = ReturnClipAndProportion(pAOIPolygon, MyGeometricOperations.EnvelopeToPolygon(pEnv), _
            pFClass, pMxDoc, lngIndex = 8 And lngIndex2 = 22, booForceFullArea, booUseDifferentArea, _
            pTileFClass)
        Set pClipPoly = varReturn(0)
        dblFullArea = varReturn(1)
        dblPolygonSumArea = varReturn(2)
        dblClipArea = varReturn(3)
        dblProportion = varReturn(4)

        vCellValue = varPixels(lngIndex, lngIndex2)
        If Not IsNull(vCellValue) Then
          lngMaxIndex = lngMaxIndex + 1
          ReDim Preserve dblReturn(4, lngMaxIndex)
          dblReturn(0, lngMaxIndex) = CDbl(vCellValue)
          dblReturn(1, lngMaxIndex) = dblFullArea
          dblReturn(2, lngMaxIndex) = dblPolygonSumArea
          dblReturn(3, lngMaxIndex) = dblClipArea
          dblReturn(4, lngMaxIndex) = dblProportion
        End If

      End If
    Next lngIndex2
  Next lngIndex

  CellValues2_Fast_byArray_VectorAdjust = dblReturn

ClearMemory:
  vCellValue = Null
  Set pEnv = Nothing
  Set pClipPoly = Nothing
  varReturn = Null
  Erase dblReturn
  varPixels = Null

End Function


