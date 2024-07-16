Attribute VB_Name = "Margaret_4"
Option Explicit

Public Sub ExtractCodeForDataPaper()

  Debug.Print "--------------------------------"

  Dim strReport As String
  Dim pColl As New Collection
  Dim pFunctions As New Collection
  Dim strArray() As String

  Dim varPaths() As Variant

    varPaths = Array( _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\aml_func_mod.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\CompareCodeModules.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\GridFunctions.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Map_Module.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Map_Module_Export_GeoTIFFs.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Margaret_4.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Margaret_Functions_3.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Margaret_maps.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Metadata_Functions.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\More_Margaret_Functions.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\MyGeneralOperations.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\MyGeometricOperations.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\QuickSort.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Repair_VM.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\SA_PRISM_Analysis.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\SierraAncha.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\SierraAnchaAnalysis.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\SierraAncha_Compare.bas", _
    "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\TestFunctions.bas")
  Dim strPath As String
  Dim strModuleName As String

  strReport = "  dim varPaths() as variant" & vbCrLf & "  varPaths = array( _" & vbCrLf
  Dim lngIndex As Long
  For lngIndex = 0 To UBound(varPaths)
    DoEvents
    strPath = varPaths(lngIndex)
    strModuleName = aml_func_mod.ReturnFilename2(strPath)
    strModuleName = aml_func_mod.ClipExtension2(strModuleName)

    Debug.Print "Reading '" & strModuleName & "'..."

    Call FillFunctionArrayAndCollection3(MyGeneralOperations.ReadTextFile(strPath), strArray, pColl, pFunctions, strModuleName)
  Next lngIndex

  Dim strPrimaryFunctions() As String
  MyGeneralOperations.AddValueToStringArray "OrganizeData_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "ReviseShapefiles_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "ConvertPointShapefiles_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "RepairOverlappingPolygons_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "AddEmptyFeaturesAndFeatureClasses_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "RecreateSubsetsOfConvertedDatasets_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "AddEmptyFeaturesAndFeatureClassesToCleaned_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "ShiftFinishedShapefilesToCoordinateSystem_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "ExportFinalDataset_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "SummarizeSpeciesBySite_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "SummarizeSpeciesByCorrectQuadrat_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "SummarizeYearByCorrectQuadratByYear_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "GenerateRData", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "ExportSubsetsOfSpeciesShapefiles_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "CreateFinalTables_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "ExportImages_SA", strPrimaryFunctions, False
  MyGeneralOperations.AddValueToStringArray "MakePageNumbers", strPrimaryFunctions, False

  Dim strFunction As String
  Dim strFunctionText As String
  Dim strAltModules() As String
  Dim strRunningFunctions() As String
  Dim lngIndex2 As Long
  Dim strCheckFunction As String
  Dim strModuleText As String
  Dim booFoundFunction As Boolean
  Dim strSpecifiedModuleName As String

  For lngIndex = 0 To UBound(strPrimaryFunctions)
    strFunction = strPrimaryFunctions(lngIndex)
    strModuleName = ReturnModuleName(strFunction, strArray, strAltModules)
    strModuleText = pFunctions.Item(strModuleName & ":" & strFunction)

    For lngIndex2 = 0 To UBound(strArray, 2)
      If strArray(2, lngIndex2) = "False" Then
        strCheckFunction = strArray(0, lngIndex2)
        If strCheckFunction = "ReturnReplacementColl_SA" Or lngIndex2 Mod 68 = 0 Or strCheckFunction = "FillQuadratNameColl_Rev_SA" Then
          DoEvents
        End If
        If strCheckFunction = "FillQuadratNameColl_Rev_SA" Then
          DoEvents
        End If
        booFoundFunction = FunctionFound(strModuleText, strCheckFunction, strSpecifiedModuleName)
        If booFoundFunction Then
          If strSpecifiedModuleName = "" Then
            strSpecifiedModuleName = ReturnModuleName(strCheckFunction, strArray, strAltModules)
          End If
          strArray(2, lngIndex) = "True"

          MyGeneralOperations.AddValueToStringArray strSpecifiedModuleName & ":" & strCheckFunction, strRunningFunctions, True

          If Not MyGeneralOperations.CheckCollectionForKey(pFunctions, strSpecifiedModuleName & ":" & strCheckFunction) Then
            Debug.Print "Failed to find '"; strSpecifiedModuleName & ":" & strCheckFunction & "'..."
          End If
        End If
      End If
    Next lngIndex2

  Next lngIndex

  Debug.Print "Initial Round:  Found the following " & Format(UBound(strRunningFunctions) + 1, "0") & " functions"
  For lngIndex = 0 To UBound(strRunningFunctions)
    Debug.Print CStr(lngIndex + 1) & "] " & strRunningFunctions(lngIndex)
  Next lngIndex

  Dim lngIncrement As Long
  lngIncrement = 0
  Dim strFunctionsToCheck() As String
  Dim strModuleAndFunction() As String

  strFunctionsToCheck = strRunningFunctions
  Do Until Not IsDimmed(strRunningFunctions)
    Erase strRunningFunctions
    lngIncrement = lngIncrement + 1
    For lngIndex = 0 To UBound(strFunctionsToCheck)
      strModuleAndFunction = Split(strFunctionsToCheck(lngIndex), ":")
      strFunction = strModuleAndFunction(1) ' strFunctionsToCheck(lngIndex)
      strModuleName = strModuleAndFunction(0) 'ReturnModuleName(strFunction, strArray, strAltModules)
      strModuleText = pFunctions.Item(strModuleName & ":" & strFunction)

      If strFunction = "ParseString" Then
        DoEvents
      End If

      For lngIndex2 = 0 To UBound(strArray, 2)
        If strArray(2, lngIndex2) = "False" Then
          strCheckFunction = strArray(0, lngIndex2)
          If strFunction = "ParseString" Then
            DoEvents
          End If
          If strCheckFunction = "ReturnReplacementColl_SA" Or lngIndex2 Mod 68 = 0 Or strCheckFunction = "FillQuadratNameColl_Rev_SA" Then
            DoEvents
          End If
          If strCheckFunction = "FillQuadratNameColl_Rev_SA" Then
            DoEvents
          End If
          booFoundFunction = FunctionFound(strModuleText, strCheckFunction, strSpecifiedModuleName)
          If booFoundFunction Then
            If strSpecifiedModuleName = "" Then
              strSpecifiedModuleName = ReturnModuleName(strCheckFunction, strArray, strAltModules)
            End If
            strArray(2, lngIndex) = "True"

            MyGeneralOperations.AddValueToStringArray strSpecifiedModuleName & ":" & strCheckFunction, strRunningFunctions, True

            If Not MyGeneralOperations.CheckCollectionForKey(pFunctions, strSpecifiedModuleName & ":" & strCheckFunction) Then
              Debug.Print "Failed to find '"; strSpecifiedModuleName & ":" & strCheckFunction & "'..."
            End If
          End If
        End If
      Next lngIndex2

    Next lngIndex

    If IsDimmed(strRunningFunctions) Then
      Debug.Print "Round " & Format(lngIncrement, "0") & ":  Found the following " & Format(UBound(strRunningFunctions) + 1, "0") & " functions"
      For lngIndex = 0 To UBound(strRunningFunctions)
        Debug.Print CStr(lngIndex + 1) & "] " & strRunningFunctions(lngIndex)
      Next lngIndex
    Else
      Debug.Print "Found no new functions..."
    End If
    strFunctionsToCheck = strRunningFunctions
  Loop

  Dim strBaseString As String
  strBaseString = strBaseString & "Attribute VB_Name = ""zzzModName""" & vbNewLine
  strBaseString = strBaseString & "Option Explicit" & vbNewLine
  strBaseString = strBaseString & " " & vbNewLine

  Dim strExportFile As String
  Dim strFileNames() As String
  Dim pExportModules As New Collection
  For lngIndex = 0 To UBound(strArray, 2)
    If strArray(2, lngIndex) = "True" Then
      strFunction = strArray(0, lngIndex)
      strModuleName = strArray(1, lngIndex)

      If Not MyGeneralOperations.CheckCollectionForKey(pExportModules, strModuleName) Then
        strExportFile = Replace(strBaseString, "zzzModName", strModuleName, , , vbTextCompare)
        MyGeneralOperations.AddValueToStringArray _
            "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Code_Files\" & strModuleName & ".bas", strFileNames, True
      Else
        strExportFile = pExportModules.Item(strModuleName)
        pExportModules.Remove strModuleName
      End If
      strModuleText = pFunctions.Item(strModuleName & ":" & strFunction)
      strExportFile = strExportFile & vbCrLf & strModuleText
      pExportModules.Add strExportFile, strModuleName
    End If
  Next lngIndex

  Dim strNewModule As String
  Dim strSplit() As String
  Dim strLine As String
  For lngIndex = 0 To UBound(strFileNames)
    strModuleName = aml_func_mod.ReturnFilename2(aml_func_mod.ClipExtension2(strFileNames(lngIndex)))
    strNewModule = ""
    strExportFile = pExportModules.Item(strModuleName)
    strSplit = Split(strExportFile, vbCrLf)
    For lngIndex2 = 0 To UBound(strSplit)
      If lngIndex2 Mod 200 = 0 Then
        DoEvents
      End If
      DoEvents
      strLine = strSplit(lngIndex2)
      Do Until Right(strLine, 1) <> " "
        strLine = Left(strLine, Len(strLine) - 1)
      Loop
      If Left(Trim(strLine), 1) <> "'" Then
        strNewModule = strNewModule & strLine & vbCrLf
      End If
    Next lngIndex2

    Do Until InStr(1, strNewModule, vbCrLf & vbCrLf & vbCrLf, vbTextCompare) = 0
      strNewModule = Replace(strNewModule, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    Loop

    pExportModules.Remove strModuleName
    pExportModules.Add strNewModule, strModuleName
    strSplit = Split(strNewModule, vbCrLf)
  Next lngIndex

  For lngIndex = 0 To UBound(strFileNames)
    strModuleName = aml_func_mod.ReturnFilename2(aml_func_mod.ClipExtension2(strFileNames(lngIndex)))
    strExportFile = pExportModules.Item(strModuleName)
    Debug.Print "Writing '" & strModuleName & "'..."
    MyGeneralOperations.WriteTextFile strFileNames(lngIndex), strExportFile, True, False
  Next lngIndex

  Debug.Print "Done..."

End Sub

Public Function FunctionFound(strFunctionText As String, strSearchFunctionName As String, strSpecifiedModuleName As String) As Boolean

  Dim lngPos As Long
  Dim strChar As String
  Dim strValid As String
  strSpecifiedModuleName = ""

  strValid = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_0123456789"
  lngPos = InStr(1, strFunctionText, strSearchFunctionName, vbTextCompare)
  If lngPos = 0 Then
    FunctionFound = False
  Else
    strChar = Mid(strFunctionText, lngPos + Len(strSearchFunctionName), 1)
    If InStr(1, strValid, strChar) = 0 Then
      FunctionFound = True
      lngPos = lngPos - 1
      strChar = Mid(strFunctionText, lngPos, 1)
      If strChar = "." Then
        lngPos = lngPos - 1
        strChar = Mid(strFunctionText, lngPos, 1)
        Do Until InStr(1, strValid, strChar) = 0
          lngPos = lngPos - 1
          strChar = Mid(strFunctionText, lngPos, 1)
        Loop
        lngPos = lngPos + 1
        strChar = Mid(strFunctionText, lngPos, 1)
        Do Until strChar = "."
          strSpecifiedModuleName = strSpecifiedModuleName & strChar
          lngPos = lngPos + 1
          strChar = Mid(strFunctionText, lngPos, 1)
        Loop
      End If
    End If
  End If

End Function

Public Function ReturnModuleName(strFunction As String, strArray() As String, strAltModules() As String, _
       Optional strSpecifiedModule As String = "") As String

  Dim lngIndex As Long
  Dim booFoundModule As Boolean
  Dim strModuleName As String
  booFoundModule = False

  Erase strAltModules

  For lngIndex = 0 To UBound(strArray, 2)
    If StrComp(strFunction, strArray(0, lngIndex), vbTextCompare) = 0 Then
      If strSpecifiedModule = "" Then
        If booFoundModule Then
          Select Case strFunction
            Case "DeclareWorkspaces"
              strModuleName = "SierraAnchaAnalysis"
            Case "DeclareWorkspaces"
              strModuleName = "Margaret_Functions"
            Case "Graphic_MakeFromGeometry"
              strModuleName = "MyGeneralOperations"
            Case Else
              Debug.Print "Unexpected event with Function '" & strFunction & "'..."
              MyGeneralOperations.AddValueToStringArray strArray(1, lngIndex), strAltModules, True
          End Select
        Else
          strModuleName = strArray(1, lngIndex)
          strArray(2, lngIndex) = "True"
          booFoundModule = True
        End If
      Else
        If strArray(1, lngIndex) = strSpecifiedModule Then
          strModuleName = strSpecifiedModule
          strArray(2, lngIndex) = "True"
          booFoundModule = True
        Else
          MyGeneralOperations.AddValueToStringArray strArray(1, lngIndex), strAltModules, True
        End If
      End If
    End If
  Next lngIndex

  If booFoundModule = False Then
    Debug.Print "Failed to find the module for function '" & strFunction & "'; using specified module '" & strSpecifiedModule & "'..."
  Else
    ReturnModuleName = strModuleName
  End If

End Function

Public Sub FillFunctionArrayAndCollection3(strModule As String, strFunctionNames() As String, _
    pFunctionItems As Collection, pFunctions As Collection, strModuleName As String)

  Dim strLines() As String
  strLines = Split(strModule, vbCrLf)
  Dim lngIndex As Long
  Dim strLine As String
  Dim strLineSplit() As String
  Dim strName As String
  Dim lngCounter As Long
  Dim strFunction As String
  Dim lngIndex2 As Long
  Dim booFoundEnd As Boolean
  Dim strType As String

  If IsDimmed(strFunctionNames) Then
    lngCounter = UBound(strFunctionNames, 2)
  Else
    lngCounter = -1
  End If

  For lngIndex = 0 To UBound(strLines)
    strLine = strLines(lngIndex)
    strLine = Replace(strLine, "Public ", "", , , vbTextCompare)
    strLine = Replace(strLine, "Private ", "", , , vbTextCompare)
    strLine = Replace(strLine, "Friend ", "", , , vbTextCompare)
    strLine = Trim(strLine)
    If Left(strLine, 4) = "Sub " Or Left(strLine, 9) = "Function " Or Left(strLine, 5) = "Enum " Then

      strType = ""
      If Left(strLine, 4) = "Sub " Then
        strType = "Sub"
      End If
      If Left(strLine, 9) = "Function " Then
        strType = "Function"
      End If
      If Left(strLine, 5) = "Enum " Then
        strType = "Enum"
      End If

      If strName = "ExportDataBySpecies" Then
        DoEvents
      End If

      strLineSplit = Split(strLine, " ")
      strLineSplit = Split(strLineSplit(1), "(", , vbTextCompare)
      strName = strLineSplit(0)

      If strModuleName = "pFClass" Or strName = "Index" Then
        DoEvents
      End If

      If strModuleName & ":" & strName = "Margaret_Functions:FillQuadratNameColl_Rev_SA" Then
        DoEvents
      End If

      If strName = "FillQuadratNameColl_Rev_SA" Then
        DoEvents
      End If

      lngCounter = lngCounter + 1
      If lngCounter = 38 Then
        DoEvents
      End If
      ReDim Preserve strFunctionNames(2, lngCounter)
      strFunctionNames(0, lngCounter) = strName
      strFunctionNames(1, lngCounter) = strModuleName
      strFunctionNames(2, lngCounter) = "False"

      pFunctionItems.Add True, strModuleName & ":" & strName

      strFunction = ""
      booFoundEnd = False
      lngIndex2 = lngIndex - 1
      Do Until booFoundEnd
        lngIndex2 = lngIndex2 + 1
        strLine = strLines(lngIndex2)
        booFoundEnd = Left(Trim(strLine), 4 + Len(strType)) = "End " & strType
        strFunction = strFunction & strLine & vbCrLf
      Loop

      pFunctions.Add strFunction, strModuleName & ":" & strName

    End If
  Next lngIndex

  GoTo ClearMemory
ClearMemory:
  Erase strLines
  Erase strLineSplit

End Sub

Public Sub CheckBoutelouaSpelling()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFiles As esriSystem.IStringArray
  Set pFiles = MyGeneralOperations.ReturnFilesFromNestedFolders2( _
      "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Special_Exports\CJ_Feb_18_2024", ".dbf")

  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Dim pProg As IStepProgressor
  Set pProg = pSBar.ProgressBar

  Dim pSpeciesColl As New Collection
  Dim strSpeciesNames() As String
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New ShapefileWorkspaceFactory
  Dim lngIndex As Long

  Dim strFullPath As String
  Dim strDir As String
  Dim strName As String
  Dim pFClass As IFeatureClass
  Dim lngSpeciesIndex As Long
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strSpecies As String

  pProg.position = 0
  pSBar.ShowProgressBar "Working on Shapefile #", 0, pFiles.Count, 1, True

  For lngIndex = 0 To pFiles.Count - 1
    pProg.Step
    pProg.Message = "Working on Shapefile #" & Format(lngIndex, "0") & " of " & Format(pFiles.Count, "#,##0")
    DoEvents
    strFullPath = pFiles.Element(lngIndex)
    strDir = aml_func_mod.ReturnDir3(strFullPath, False)
    strName = aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strFullPath))
    Set pWS = pWSFact.OpenFromFile(strDir, 0)
    Set pFClass = pWS.OpenFeatureClass(strName)
    lngSpeciesIndex = pFClass.FindField("Species")
    Set pFCursor = pFClass.Search(Nothing, True)
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing
      strSpecies = pFeature.Value(lngSpeciesIndex)
      AddAndIncrement pSpeciesColl, strSpeciesNames, strSpecies
      Set pFeature = pFCursor.NextFeature
    Loop

  Next lngIndex

  QuickSort.StringsAscending strSpeciesNames, 0, UBound(strSpeciesNames)
  For lngIndex = 0 To UBound(strSpeciesNames)
    strSpecies = strSpeciesNames(lngIndex)
    Debug.Print Format(lngIndex + 1, "0") & "] " & strSpecies & ": n = " & Format(pSpeciesColl.Item(strSpecies), "#,##0") & " rows"
  Next lngIndex

  pProg.position = 0
  pSBar.HideProgressBar

  Debug.Print "Found " & Format(pFiles.Count, "0") & " shapefiles..."
  Debug.Print "Done..."

End Sub

Public Sub AddAndIncrement(pColl As Collection, strArray() As String, strSpecies As String)

  Dim lngCount As Long
  If MyGeneralOperations.CheckCollectionForKey(pColl, strSpecies) Then
    lngCount = pColl.Item(strSpecies)
    pColl.Remove strSpecies
  Else
    lngCount = 0
    MyGeneralOperations.AddValueToStringArray strSpecies, strArray, True
  End If
  lngCount = lngCount + 1
  pColl.Add lngCount, strSpecies
End Sub


