Attribute VB_Name = "SierraAncha"
Option Explicit

Public Function ReturnYearFromName(strName As String) As Long

  Dim varYears() As Variant
  varYears = Array(1395, 1935, 1936, 1940, 1941, 1942, 1952, 1955, 2017, 2018, 2019, 2020)
  Dim lngYear As Long
  Dim lngYearIndex As Long
  For lngYearIndex = 0 To UBound(varYears)
    lngYear = varYears(lngYearIndex)
    If InStr(1, strName, "_" & Format(lngYear, "0") & "_", vbTextCompare) > 0 Then
      ReturnYearFromName = lngYear
      If lngYear = 1395 Then lngYear = 1935
      Exit Function
    End If
  Next lngYearIndex

  MsgBox "Failed to find year in '" & strName & "'!"

End Function


