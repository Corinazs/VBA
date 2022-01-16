Sub PivotKontenAusblenden(gesellschaft As Worksheet)

  Dim bolVisible As Boolean
  Dim pvTab As PivotTable
  Dim pvField As PivotField, pvItem As PivotItem
  Dim index As Integer

  
  Set pvTab = gesellschaft.PivotTables(1)
  Set pvField = pvTab.PivotFields("Loc.Acc.")
  
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Sheets("Kontendefinition")
  letzteZeileKD = ws.Cells(Rows.Count, 2).End(xlUp).Row
  Dim arrValues As Variant
  
  Application.ScreenUpdating = False
  pvTab.RefreshTable
  pvField.ClearAllFilters

  For Each pvItem In pvField.VisibleItems
    bolVisible = False
    
    For index = 2 To letzteZeileKD
        If pvItem.name = ws.Cells(index, 2).Value Then
            bolVisible = True

            Exit For
        End If
    Next index
  
    If bolVisible = False Then pvItem.Visible = False
  Next pvItem
  Application.ScreenUpdating = True

End Sub
