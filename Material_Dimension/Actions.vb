Imports Material_Dimension.CATIA_Property
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Actions
    Private CatiaFactory As CATIA_Property
    Dim ExcelApp As New Excel.Application
    Dim WB As Excel.Workbook
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CatiaFactory.PartDocument IsNot Nothing Then
            Call BoundingBox(CatiaFactory.PartDocument)
        Else
            Dim i As Integer = CatiaFactory.ProductDocument.Product.Products.Count
            Dim mProductDocument As ProductStructureTypeLib.ProductDocument
            mProductDocument = CatiaFactory.ProductDocument
            For j = 1 To i
                mProductDocument.Product.Products.Item(j).ApplyWorkMode(2)      'Setting Deside Mode
                mProductDocument.Product.Products.Item(j).ActivateDefaultShape()        ' Load Default Shape
                Dim mPart As MECMOD.PartDocument = mProductDocument.Product.Products.Item(j).ReferenceProduct.Parent
                Call MutipleBoundingBox(mPart, CatiaFactory.ProductDocument)
            Next
        End If
        Dim Assembly As ProductStructureTypeLib.ProductDocument = CatiaFactory.ProductDocument
        Call HideObjectFromProduct(Assembly, "definition_points")
        Call HideObjectFromProduct(Assembly, "Refer_Axis")
        Assembly.Selection.Clear()
        Assembly.Product.Update()
    End Sub

    Private Sub Actions_Load(sender As Object, e As EventArgs) Handles Me.Load
        CatiaFactory = SetInitialCATIA()
        Me.TopLevel = True
        Me.TopMost = True
        Dimenstion.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Assembly As ProductStructureTypeLib.ProductDocument = CatiaFactory.ProductDocument
        Call HideObjectFromProduct(Assembly, "BoundingBox*")
        Assembly.Selection.Clear()
        Assembly.Product.Update()

    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim Assembly As ProductStructureTypeLib.ProductDocument = CatiaFactory.ProductDocument
        Call ShowObjectFromProduct(Assembly, "BoundingBox*")
        Assembly.Selection.Clear()
        Assembly.Product.Update()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dialog As New FolderBrowserDialog
        Dim FileLocation As String
        If dialog.ShowDialog = DialogResult.OK Then
            FileLocation = dialog.SelectedPath
            'MsgBox(FileLocation)
        Else
            Exit Sub
        End If
        WB = ExcelApp.Workbooks.Add()
        ExcelApp.Visible = True
        Dim SheetXL As Excel.Worksheet = WB.ActiveSheet
        SheetXL.Cells(1, 1).Value = "組立件名稱"
        SheetXL.Cells(3, 1).Value = "子件素材尺寸"
        SheetXL.Cells(1, 3).Value = "長"
        SheetXL.Cells(1, 4).Value = "寬"
        SheetXL.Cells(1, 5).Value = "高"

        Dim ProDoc As ProductStructureTypeLib.ProductDocument = CatiaFactory.ProductDocument
        SheetXL.Cells(1, 2).Value = ProDoc.Name
        Dim CountPart As Integer = ProDoc.Product.Products.Count
        Dim BoundingBoxData(4) As String
        Dim Pic_Cells As Excel.Range
        For i = 1 To CountPart
            BoundingBoxData = GetBoundingBoxData(ProDoc.Product.Products.Item(i))
            Dim ItemName As String = ProDoc.Product.Products.Item(i).Name
            SheetXL.Cells(3 + i, 2).Value = BoundingBoxData(1)
            SheetXL.Cells(3 + i, 3).Value = BoundingBoxData(2)
            SheetXL.Cells(3 + i, 4).Value = BoundingBoxData(3)
            SheetXL.Cells(3 + i, 5).Value = BoundingBoxData(4)
            Call ExceptSelectedItemHide(ProDoc, i)
            Dim PartImage As String = GetISO_ViewPoint(ProDoc.Product.Products.Item(i), CatiaFactory.myCATIA, FileLocation)
            Pic_Cells = SheetXL.Cells(3 + i, 1)
            SheetXL.Shapes.AddPicture(PartImage, False, True, Pic_Cells.Left, Pic_Cells.Top, 150, 100)
            SheetXL.Cells(3 + i, 1).RowHeight = 100
            Kill(FileLocation & "\" & ItemName & ".JPG")
        Next
        SheetXL.Columns("A:F").AutoFit
        SheetXL.Cells(1, 1).ColumnWidth = 25
        If Len(Dir(FileLocation & "\" & Replace(ProDoc.Name, ".CATProduct", "") & "_BoundingBox_List.xls")) > 0 Then
            Kill(FileLocation & "\" & Replace(ProDoc.Name, ".CATProduct", "") & "_BoundingBox_List.xls")
        End If
        WB.SaveAs(FileLocation & "\" & Replace(ProDoc.Name, ".CATProduct", "") & "_BoundingBox_List", FileFormat:=56)
        WB.Close()
        ExcelApp.Quit()
        For i = 1 To CountPart
            Call ShowObjectFromProduct(ProDoc, ProDoc.Product.Products.Item(i).Name)
        Next
        CatiaFactory.myCATIA.StartCommand("Fit All In")
    End Sub
    Private Shared Function GetBoundingBoxData(ByRef mProduct As ProductStructureTypeLib.Product) As Array
        Dim Data(4) As String
        Data(1) = mProduct.Name
        Dim mPart As MECMOD.PartDocument = mProduct.ReferenceProduct.Parent         '抓取PartDocument
        Data(2) = mPart.Part.Parameters.Item("Bbox_dx.1").ValueAsString
        Data(3) = mPart.Part.Parameters.Item("Bbox_dy.1").ValueAsString
        Data(4) = mPart.Part.Parameters.Item("Bbox_dz.1").ValueAsString
        Return Data
    End Function
    Private Sub Actions_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Dimenstion.Hide()
    End Sub

End Class