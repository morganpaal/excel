Sub URLPictureInsert()
Dim cell, shp As Shape, target As Range
    Set Rng = ActiveSheet.Range("<<INSERT CELL RANGE HERE>>")
    For Each cell In Rng
       filenam = cell
       ActiveSheet.Pictures.Insert(filenam).Select
       

  Set shp = Selection.ShapeRange.Item(1)
   With shp
      .LockAspectRatio = msoTrue
      .Width = 100
      .Height = 100
      .Cut
   End With
   Cells(cell.Row, cell.Column + 1).PasteSpecial
Next

End Sub
