Imports ProductStructureTypeLib
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Comparator
    ''' <summary>
    ''' WalksDown the 3D Tree in CATIA
    ''' </summary>
    Public Children3D As New Collection
    Public Children2D As New Collection


    Sub WalkDownTree(oInProduct As Object)  'As Product)

        Dim Validation As New Validation

        Dim oInstances As Products
        oInstances = oInProduct.Products

        '-----No instances found then this is CATPart

        If oInstances.Count = 0 Then
            'MsgBox "This is a CATPart with part number " & oInProduct.PartNumber
            Exit Sub
        End If

        '-----Found an instance therefore it is a CATProduct
        'MsgBox "This is a CATProduct with part number " & oInProduct.ReferenceProduct.PartNumber

        Dim k As Integer
        For k = 1 To oInstances.Count


            Dim oInst 'As Object
            oInst = oInstances.Item(k)

            Children3D.Add(oInst)

            'oInstances.Item(k).ApplyWorkMode(DESIGN_MODE)  'apply design mode

            'If oInstances.Item(k).Parent.Parent.PartNumber = "B4818GAED-101" Then
            '  If Validation.IsComponent(oInst) = True Then
            Call WalkDownTree(oInst)
            'End If
        Next


    End Sub

    Sub Write3DToExcel()

        'Dim XL As Object
        'Try
        '    XL = GetObject(, "Excel.Application")
        'Catch ex As Exception
        '    XL = CreateObject("Excel.Application")
        'End Try


        Dim oXL As Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        Try
            oXL = GetObject(, "Excel.Application")
            ' oXL.Sheets(1).Cells.Clear()
        Catch ex As Exception
            oXL = New Excel.Application


        End Try

        oXL.DisplayAlerts = False
        oXL.Visible = True
        'oXL.Workbooks.Add()


        ' Dim i As Integer
        Dim j As Integer


        'j = 1
        'For i = 1 To Children.Count

        '    If Children(i).PartNumber = "79A5552" Then
        '        Children(i).Parent.Parent.PartNumber = "HKLHJKHJKHJKHJKH"
        '    End If
        '    'If Children(i).Parent.Name = "B472289-527" Then
        '    XL.sheets(1).Cells(j + 3, 1).Value = 1
        '    XL.sheets(1).Cells(j + 3, 2).Value = Children(i).PartNumber
        '    XL.sheets(1).Cells(j + 3, 3).Value = Children(i).ReferenceProduct.Parent.Name
        '    'Cells(i + 13, 3).Value = Children(i).Name
        '    'Cells(i + 13, 3).Value = IsComponent(PartNumbers(i))
        '    XL.sheets(1).Cells(j + 13, 4).Value = Children(i).Parent.Parent.partnumber
        '    ' End If
        '    j = j + 1
        '    'End If
        'Next i


        Dim Realchildren = From child As Object In Children3D _
                            Select child.partnumber
        ' Where child.parent.parent.partnumber = "Product85"


        'oXL.Sheets(1).range("a1").CopyFromRecordset(Realchildren)
        ',child.nomenclature, child.ReferenceProduct.Parent.Name ', child.partnumber

        j = 1
        For i = 0 To Realchildren.Count

            'If Realchildren(i).PartNumber = "79A5552" Then
            '    Realchildren(i).Parent.Parent.PartNumber = "HKLHJKHJKHJKHJKH"
            'End If
            'If Realchildren(i).Parent.Name = "B472289-527" Then
            oXL.Sheets(1).Cells(i + 3, 1).Value = 1
            oXL.Sheets(1).Cells(i + 3, 2).Value = Realchildren(i)
            'oXL.Sheets(1).Cells(j + 3, 3).Value = Realchildren(i).Name
            'oXL.Sheets(1).Cells(j + 3, 3).Value = Realchildren(i).ReferenceProduct.Parent.Name
            ''Cells(i + 13, 3).Value = Realchildren(i).Name
            ''Cells(i + 13, 3).Value = IsComponent(PartNumbers(i))
            'XL.sheets(1).Cells(j + 13, 4).Value = Realchildren(i).Parent.Parent.partnumber
            ' End If
            'j = j + 1
            'End If
        Next i



    End Sub


    Sub Select3D()

        Dim CATIA As Object
        CATIA = GetObject(, "CATIA.Application")

        Dim ActiveProductDocument As ProductDocument

        Try
            ActiveProductDocument = CATIA.ActiveDocument
        Catch ex As Exception
            MsgBox("Rather than a beep" & vbCrLf & "Or a rude error message:" & vbCrLf & "Open a CATProduct in the active session")

            Exit Sub
        End Try


        Dim ActProd As Products
        ActProd = ActiveProductDocument.Product

        Dim what(1)
        what(0) = "Product"
        what(1) = "Part"

        Dim UserSel As Object
        UserSel = CATIA.ActiveDocument.Selection
        UserSel.Clear()




        Dim e As String
        e = UserSel.selectelement3(what, "Select a Product or a Component", 0, 2, 0)


        Dim SelectedElement As Integer

        Dim SelectedCollection As New Collection

        For SelectedElement = 1 To UserSel.Count

            SelectedCollection.Add(UserSel.Item(SelectedElement).Value)
        Next SelectedElement

        UserSel.Clear()



        Dim SelectedProductItem As Integer

        For SelectedProductItem = 1 To SelectedCollection.Count

            Dim oRootProd As Products
            oRootProd = SelectedCollection(SelectedProductItem)
            MsgBox("This is a CATPart with part number " & oRootProd.PartNumber)

            Call WalkDownTree(oRootProd)

        Next SelectedProductItem
        Dim count As Integer = 1
        Call Write3DToExcel()

        'Call Write3DToExcel2(QTYS, PartNumbers)

        '***************************************************************

        'Get the current CATIA assembly

        'Dim oProdDoc As ProductDocument
        'oProdDoc = CATIA.ActiveDocument

        'Dim oRootProd As Products
        'oRootProd = oProdDoc.Product

        'MsgBox("This is a CATPart with part number " & oRootProd.PartNumber)

        'Call WalkDownTree(oRootProd)
        'Call WriteToExcel()

        'MsgBox("Done " & Children(Children.Count).partnumber)
    End Sub
    Sub Select2D()

        Dim CATIA As Object
        CATIA = GetObject(, "CATIA.Application")

        Dim oXL As Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        Try
            oXL = GetObject(, "Excel.Application")
            ' oXL.Sheets(1).Cells.Clear()
        Catch ex As Exception
            oXL = New Excel.Application


        End Try

        oXL.DisplayAlerts = False
        oXL.Visible = True

        'Dim ActiveDrawingDocument 'As DrawingDocument
        'ActiveDwgDocument = CATIA.ActiveDocument

        'Dim ActProd As Product
        'Set ActDrw = ActiveProductDocument.Drawing

        'Dim ProductChildren As Products
        'Set ProductChildren = ActProd.Products


        Dim what(0)
        what(0) = "DrawingTable"
        Dim UserSel2D
        'Dim UserSel As SELECTEDELEMENT
        UserSel2D = CATIA.ActiveDocument.Selection
        UserSel2D.Clear()

        'Dim e As catbstr
        Dim e 'As String
        'e = UserSel2D.selectelement2(what, "Select a Product or a Component", False)
        e = UserSel2D.selectelement3(what, "Select a Product or a Component", True, 2, False)
        'UserForm1.Show
        'UserForm1.TextBox1 = (UserSel.Item(1).Value.partnumber)

        'MsgBox (UserSel2D.Item(1).Value.Name)
        'UserForm1.Show
        'UserForm1.TextBox1 = (UserSel.Item(1).Value.partnumber)

        Dim actTable
        actTable = UserSel2D.Item(1).Value

        Dim Dwg As Drawing = New Drawing
        Dim NumberOfUsefulCol As Integer
        Dim NumberOfUsefulRows As Integer
        Dim PartNumberCol As Integer

        Dim QtyCol
        QtyCol = 1


        Dim ItemNo = New Drawing.ItemNo
        Dim MatSpec = New Drawing.MatlSpec
        Dim Nomenclature = New Drawing.Nomenclature
        Dim PartNo = New Drawing.PartNo

        ItemNo.Column = actTable.NumberOfColumns
        MatSpec.Column = actTable.NumberOfColumns - 1
        Nomenclature.Column = actTable.NumberOfColumns - 2
        PartNo.Column = actTable.NumberOfColumns - 3

        ' The assemblies are between Column 1 and NUmberofCoumns-5

        'For i = 1 To actTable.NumberOfrows ' This would only work on the first table selected
        '    If Left(actTable.getcellstring(i + 1, PartNo.Column), 1) = "-" Then
        '        If CInt(Mid(actTable.getcellstring(i + 1, PartNo.Column), 2, 3)) Mod 100 >= 1 Then
        '            If IsNumeric(actTable.getcellstring(i + 1, ItemNo.Column)) = False Then
        '                Dwg.Cols = i
        '                Exit For
        '            End If
        '        End If
        '    End If
        'Next i

        'For i = 1 To actTable.NumberOfrows
        '    oXL.Cells(i + 13, 5) = actTable.getcellstring(i, QtyCol)
        '    oXL.Cells(i + 13, 6) = actTable.getcellstring(i, PartNo.Column)
        '    oXL.Cells(i + 13, 7) = actTable.getcellstring(i, Nomenclature.Column)
        '    oXL.Cells(i + 13, 8) = actTable.getcellstring(i, MatSpec.Column)
        '    oXL.Cells(i + 13, 9) = actTable.getcellstring(i, ItemNo.Column)
        'Next i
        For j = 1 To actTable.NumberOfColumns

            For i = 1 To actTable.NumberOfrows - 1
                oXL.Cells(i + 13, j) = actTable.getcellstring(i, j)
            Next i

        Next


        UserSel2D.Clear()
    End Sub
    Sub Write2DToExcel()
        
    End Sub
    ''' <summary>
    ''' Returns the real parent of a component
    ''' </summary>
    Function RealParent() As String

    End Function

    Sub Is3DPartIn2D()

    End Sub
    Sub Is2DPartIn3D()

    End Sub
    Sub Is3DQtyEquals2DQty()

    End Sub

End Class
