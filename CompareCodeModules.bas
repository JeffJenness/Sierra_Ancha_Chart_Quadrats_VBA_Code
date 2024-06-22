Attribute VB_Name = "CompareCodeModules"
Option Explicit

Public Sub TestCompare()
  Debug.Print "-----------------------------"
  Dim strPath1 As String
  Dim strPath2 As String

  strPath1 = "D:\arcGIS_stuff\Teaching\FOR_525\2023\00_Homework\MyGeneralOperations.bas"
  strPath2 = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\MyGeneralOperations.bas"

  Dim strNameArray1() As String
  Dim pColl1 As Collection
  Dim strNameArray2() As String
  Dim pColl2 As Collection
  Dim lngIndex As Long
  Dim strRemainingText As String
  Dim strMissingIn1 As String
  Dim strMissingIn2 As String
  Dim strLargerIn1 As String
  Dim strLargerIn2 As String
  Dim strOutput As String

  Set pColl1 = ReturnCollectionOfFunctionsAndCode(strPath1, strNameArray1, strRemainingText)
  Set pColl2 = ReturnCollectionOfFunctionsAndCode(strPath2, strNameArray2, strRemainingText)
  strOutput = CompareModules(strNameArray1, pColl1, strPath1, strNameArray2, pColl2, strPath2, strMissingIn1, strMissingIn2, _
      strLargerIn1, strLargerIn2)

  strOutput = "Found " & Format(UBound(strNameArray1) + 1, "0") & " functions, subs, consts and enums in " & strPath1 & vbCrLf & _
       strMissingIn1 & vbCrLf & _
       strLargerIn1 & vbCrLf & vbCrLf & _
       "Found " & Format(UBound(strNameArray2) + 1, "0") & " functions, subs, consts and enums in " & strPath2 & vbCrLf & _
       strMissingIn2 & vbCrLf & _
       strLargerIn2

  Dim pDataObj As New MSForms.DataObject
  pDataObj.SetText strOutput
  pDataObj.PutInClipboard
  Set pDataObj = Nothing

  Debug.Print "Found " & Format(UBound(strNameArray1) + 1, "0") & " functions, subs, consts and enums in " & strPath1
  Debug.Print strMissingIn1
  Debug.Print strLargerIn1

  Debug.Print "Found " & Format(UBound(strNameArray2) + 1, "0") & " functions, subs, consts and enums in " & strPath2
  Debug.Print strMissingIn2
  Debug.Print strLargerIn2

End Sub

Public Function CompareModules(strNameArray1() As String, pColl1 As Collection, strPath1 As String, strNameArray2() As String, _
      pColl2 As Collection, strPath2 As String, strMissingIn1 As String, strMissingIn2 As String, _
      strLargerIn1 As String, strLargerIn2 As String) As String

  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim strName As String
  Dim lngMissingCount As Long
  Dim lngLargerCount1 As Long
  Dim lngLargerCount2 As Long
  Dim strCode1 As String
  Dim strCode2 As String
  Dim strPair() As String

  strLargerIn1 = ""
  strLargerIn2 = ""
  For lngIndex = 0 To UBound(strNameArray1)
    strName = strNameArray1(lngIndex)
    If MyGeneralOperations.CheckCollectionForKey(pColl1, strName) And MyGeneralOperations.CheckCollectionForKey(pColl2, strName) Then
      strPair = pColl1.Item(strName)
      strCode1 = strPair(1)
      strPair = pColl2.Item(strName)
      strCode2 = strPair(1)
      If Len(strCode1) > Len(strCode2) Then
        lngLargerCount1 = lngLargerCount1 + 1
        strLargerIn1 = strLargerIn1 & "    " & Format(lngLargerCount1, "0") & "] " & strName & vbCrLf
      ElseIf Len(strCode2) > Len(strCode1) Then
        lngLargerCount2 = lngLargerCount2 + 1
        strLargerIn2 = strLargerIn2 & "    " & Format(lngLargerCount2, "0") & "] " & strName & vbCrLf
      End If
    End If
  Next lngIndex
  If lngLargerCount1 = 0 Then
    strLargerIn1 = "  No items from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & _
        " were larger than the corresponding items in " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & "..." & vbCrLf & strLargerIn1
  ElseIf lngMissingCount = 1 Then
    strLargerIn1 = "  1 item from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & _
        " were larger than the corresponding items in " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & "..." & vbCrLf & strLargerIn1
  Else
    strLargerIn1 = "  " & Format(lngLargerCount1, "0") & " items from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & _
        " were larger than the corresponding items in " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & "..." & vbCrLf & strLargerIn1
  End If
  If lngLargerCount2 = 0 Then
    strLargerIn2 = "  No items from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & _
        " were larger than the corresponding items in " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & "..." & vbCrLf & strLargerIn2
  ElseIf lngMissingCount = 2 Then
    strLargerIn2 = "  1 item from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & _
        " were larger than the corresponding items in " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & "..." & vbCrLf & strLargerIn2
  Else
    strLargerIn2 = "  " & Format(lngLargerCount2, "0") & " items from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & _
        " were larger than the corresponding items in " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & "..." & vbCrLf & strLargerIn2
  End If

  strMissingIn1 = ""
  lngMissingCount = 0
  For lngIndex = 0 To UBound(strNameArray1)
    strName = strNameArray1(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pColl2, strName) Then
      lngMissingCount = lngMissingCount + 1
      strMissingIn1 = strMissingIn1 & "    " & Format(lngMissingCount, "0") & "] " & strName & vbCrLf
    End If
  Next lngIndex
  If lngMissingCount = 0 Then
    strMissingIn1 = "  No items from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & " missing from " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & "..." & vbCrLf & strMissingIn1
  ElseIf lngMissingCount = 1 Then
    strMissingIn1 = "  1 item from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & " missing from " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & "..." & vbCrLf & strMissingIn1
  Else
    strMissingIn1 = "  " & Format(lngMissingCount, "0") & " items from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & " missing from " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & "..." & vbCrLf & strMissingIn1
  End If

  strMissingIn2 = ""

  lngMissingCount = 0
  For lngIndex = 0 To UBound(strNameArray2)
    strName = strNameArray2(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pColl1, strName) Then
      lngMissingCount = lngMissingCount + 1
      strMissingIn2 = strMissingIn2 & "    " & Format(lngMissingCount, "0") & "] " & strName & vbCrLf
    End If
  Next lngIndex
  If lngMissingCount = 0 Then
    strMissingIn2 = "  No items from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & " missing from " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & "..." & vbCrLf & strMissingIn2

  ElseIf lngMissingCount = 1 Then
    strMissingIn2 = "  1 item from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & " missing from " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & "..." & vbCrLf & strMissingIn2

  Else
    strMissingIn2 = "  " & Format(lngMissingCount, "0") & " items from " & aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath2)) & " missing from " & _
        aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath1)) & "..." & vbCrLf & strMissingIn2
  End If

End Function

Public Sub AddLineToArray(strLine As String, strArray() As String)

  If Not IsDimmed(strArray) Then
    ReDim strArray(0)
    strArray(0) = strLine
  Else
    ReDim Preserve strArray(UBound(strArray) + 1)
    strArray(UBound(strArray)) = strLine
  End If

End Sub


