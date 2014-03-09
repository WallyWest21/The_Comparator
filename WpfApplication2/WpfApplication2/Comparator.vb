﻿Imports ProductStructureTypeLib
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading.Tasks
Imports ProductStructureTypeLib.CatWorkModeType
Imports INFITF.CatFileSelectionMode
Imports DRAFTINGITF
Imports MECMOD
Imports PARTITF
Imports System.Collections.ObjectModel
Imports INFITF.CATMultiSelectionMode
Imports System.Windows.Controls
Imports WpfApplication2
Imports System.Threading
Imports System.IO

'Imports INFITF
Public Class Comparator
    Dim Validation As New Validation
    Dim comp As New ChildrenList

    ''' <summary>
    ''' WalksDown the 3D Tree in CATIA
    ''' </summary>
    Public ReadOnly Children3D As New Collection
    ''' <summary>
    ''' It is a Collection of all the items in 2D
    ''' </summary>
    Public Children2D As Dictionary(Of oDrawing.PartsList, String)

    Public Property Realchildren3D As New ObservableCollection(Of String)
    Public Property LIstboxChildren2D As New ObservableCollection(Of String)
    Public Property Selected3DElements As New ObservableCollection(Of String)
    Public Property Available2DElements As New ObservableCollection(Of String)
    Public Property Selected2DElements As New ObservableCollection(Of String)
    Public ActiveDocuments As New Collection

    ''' <summary>
    ''' Returns the real parent of a component
    ''' </summary>
    Function RealParent(ByVal oInst) As String
        Dim oParent As Object
        oParent = oInst.parent.parent

        If Validation.IsComponent(oParent) = True Then
            Return RealParent(oParent.parent.parent)
        Else
            Return oParent.partnumber
        End If

    End Function
    Public Shared MaximumOfColumnsInBigTable As Integer = 0
    Public Shared MaximumOfRowsInBigTable As Integer = 0
    Public Shared Big2DTable(MaximumOfRowsInBigTable, MaximumOfColumnsInBigTable) As String
    Public StartingXLRow As Integer
    Public StartingXLColumn As Integer
    Public EndXLRow As Integer
    Public EndXLColumn As Integer


    Sub WalkDownTree(ByVal oInProduct As Object)
        'As Product)

        Dim Validation As New Validation
        Dim test As String
        Dim testobject As String
        Dim oInstances As Products
        oInstances = oInProduct.Products

        '-----No instances found then this is CATPart

        If oInstances.Count = 0 Then

            Exit Sub
        End If


        Try
            Parallel.For(1, oInstances.Count + 1, Sub(k)

                                                      Dim oInst As INFITF.AnyObject
                                                      oInst = oInstances.Item(k)
                                                      oInstances.Item(k).ApplyWorkMode(DESIGN_MODE)   'apply design mode
                                                      testobject = oInst.partnumber
                                                      'If Validation.IsComponent(oInst) = False And oInstances.Item(k).Parent.Parent.PartNumber = oInProduct.partnumber Then
                                                      If Validation.IsComponent(oInst) = False And RealParent(oInst) = oInProduct.partnumber Then

                                                          Children3D.Add(oInst)
                                                          Realchildren3D.Add(oInst.partnumber)
                                                          comp.Add(oInst.partnumber)

                                                      End If

                                                      If Validation.IsComponent(oInst) = True And RealParent(oInst) = oInProduct.partnumber Then
                                                          Call WalkDownTree(oInst)
                                                          test = RealParent(oInst)

                                                      End If

                                                  End Sub)

        Catch ex As Exception
            MsgBox("You need a multicore computer")
        End Try
        'Realchildren3D.Add("klhkjhklhjkhkjlhkljl")


        '    lst1.Add("New Item")

        'ListBox1.ItemsSource = ChildrenList.
        comp.Add("comparator")


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
            '  oXL = New Excel.Application


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

        '******************************************************************************************
        Dim Realchildren = From child In Children3D.AsParallel() _
        Group child By child.partnumber, child.nomenclature Into Group _
        Select qty = Group.Count, partnumber = partnumber, nomenclature = nomenclature


        '***************************************************************************
        '***************************************************************************


        Dim myTitle = "Hello this the comparator report"
        Dim Link As String = "D:\RHDSetup.log" '"http://www.w3schools.com"

        Dim D3HTML As XElement

        'For Each child In BOM3D
        'D3HTML = "<td>" & "1" & "</td><td><a href= " & Link & ">" & "CouldBeAVariable" & "</a></td>"
        ' For i = 1 To 5
        D3HTML =
<html>
    <table border="1" align="center">
        <tr><td>1</td><td><a href=<%= Link %>>CouldBeAVariable4</a></td></tr>
    </table>
</html>
        'D3HTML.AddAfterSelf(D3HTML)

        'Next
        ' Next



        Dim myHTML As XElement =
                    <html>
                        <head>
                            <title><%= myTitle %></title>
                        </head>
                        <body>
                            <h1>Welcome 3D to HTML report2312! Where the hood at?</h1>
                        </body>
                        <table border="1" align="center">
                            <tr><th>QTY</th><th>Part Number</th><th>Nomenclature</th></tr>
                            <tr><td>1</td><td><a href=<%= Link %>>CouldBeAVariable1</a></td></tr>
                        </table>
                    </html>



        ' myHTML.Element("html").Element("body").Element("table").Element("tr").AddAfterSelf(D3HTML)

        Dim myHTMLafter As XElement = myHTML.<table>(0)

        Dim Realchildrens = From childs In Children3D.AsParallel() _
        Group childs By childs.partnumber, childs.nomenclature Into Group _
        Select qty = Group.Count, partnumber = partnumber, nomenclature = nomenclature

        For Each results In Realchildrens
            MsgBox("yeah")
            D3HTML =
<html>
    <table border="1" align="center">
        <tr><td><%= results.qty %></td><td><a href=<%= Link %>><%= results.partnumber %></a></td></tr>
    </table>
</html>

            myHTMLafter.AddAfterSelf(D3HTML)
        Next



        '***************************************************************************************************
        '***************************************************************************************************

        '        For i = 1 To 5

        '            D3HTML =
        '<html>
        '    <table border="1" align="center">
        '        <tr><td><%= i + 12 %></td><td><a href=<%= Link %>><%= i %></a></td></tr>
        '    </table>
        '</html>

        '            myHTMLafter.AddAfterSelf(D3HTML)
        '        Next



        ' myHTMLafter.AddAfterSelf(From d3 In D3HTML.Elements() Where CInt(d3) = 4 Select d3)

        ' http://msdn.microsoft.com/en-us/library/system.xml.linq.xelement.addafterself(v=vs.110).aspx
        ' <%= D3HTML %>
        '<td>1</td><td><a href=<%= Link %>>CouldBeAVariable</a></td>
      















        '***********************************************************************************
        '***********************************************************************************





        'Dim Realchildren = From child In Children3D.AsParallel()
        ' Select child.partnumber, child.nomenclature
        '' Where child.parent.parent.partnumber = "Product85"


        'oXL.Sheets(1).range("a1").CopyFromRecordset(Realchildren)
        ',child.nomenclature, child.ReferenceProduct.Parent.Name ', child.partnumber

        j = 1
        Dim i As Integer = 1
        '  For i = 0 To Realchildren.Count


        For Each result In Realchildren

            'If Realchildren(i).PartNumber = "79A5552" Then
            '    Realchildren(i).Parent.Parent.PartNumber = "HKLHJKHJKHJKHJKH"
            'End If
            'If Realchildren(i).Parent.Name = "B472289-527" Then
            oXL.ActiveSheet.Cells(i + 13, 1).Value = result.qty
            oXL.ActiveSheet.Cells(i + 13, 2).Value = result.partnumber
            oXL.ActiveSheet.Cells(i + 13, 3).Value = result.nomenclature








            'oXL.Sheets(1).Cells(j + 3, 3).Value = Realchildren(i).Name
            'oXL.Sheets(1).Cells(j + 3, 3).Value = Realchildren(i).ReferenceProduct.Parent.Name
            ''Cells(i + 13, 3).Value = Realchildren(i).Name
            ''Cells(i + 13, 3).Value = IsComponent(PartNumbers(i))
            'XL.sheets(1).Cells(j + 13, 4).Value = Realchildren(i).Parent.Parent.partnumber
            ' End If
            'j = j + 1
            'End If
            i += 1
        Next





        'Dim Realchildren = From child As Object In Children3D.AsParallel().AsParallel _
        '                   Group By child.partnumber Into Group
        '                    Select partnumber

        ''  For Each kid In Realchildren3D
        'Console.WriteLine(1)
        ''Next


        '*************************************************
        '**************************************************


        Using writer As StreamWriter = New StreamWriter("TheComparator.html")
            writer.Write(myHTML)
        End Using


    End Sub


    Sub Select3D(Optional ByRef TheRealChildren = "boogie")


        'Dim MainWindow As New MainWindow

        Dim CATIA As INFITF.Application

        Try
            CATIA = GetObject(, "CATIA.Application")

        Catch ex As Exception
            MsgBox("The Application you seek" & vbCrLf & "Cannot be located." & vbCrLf & "Open a CATIA session.")
            Exit Sub
        End Try

        Dim ActiveProductDocument As ProductDocument

        Try
            ActiveProductDocument = CATIA.ActiveDocument
        Catch ex As Exception
            'MainWindow.Is3DSelected = False
            MsgBox("Rather than a beep" & vbCrLf & "Or a rude error message:" & vbCrLf & "Open a CATProduct in the active session")


            Exit Sub
        End Try

        Dim ActProd As Products
        ActProd = ActiveProductDocument.Product

        Dim what(0)
        what(0) = "Product"
        'what(1) = "Part"

        Dim UserSel As INFITF.Selection
        UserSel = CATIA.ActiveDocument.Selection
        UserSel.Clear()



        Dim e As String
        'e = UserSel.selectelement3(what, "Select a Product or a Component", False, 2, False)
        e = UserSel.SelectElement3(what, "Select a Product or a Component", False, CATMultiSelTriggWhenUserValidatesSelection, True)

        Dim SelectedElement As Integer
        Dim SelectedCollection As New Collection

        For SelectedElement = 1 To UserSel.Count

            SelectedCollection.Add(UserSel.Item(SelectedElement).Value)
            Selected3DElements.Add(UserSel.Item(SelectedElement).Value.partnumber)
        Next SelectedElement

        UserSel.Clear()

        Dim SelectedProductItem As Integer

        For SelectedProductItem = 1 To SelectedCollection.Count

            Dim oRootProd As Products
            oRootProd = SelectedCollection(SelectedProductItem)
            'MsgBox("This is a CATPart with part number " & oRootProd.PartNumber)

            Call WalkDownTree(oRootProd)

        Next SelectedProductItem
        Dim count As Integer = 1
        'Call Write3DToExcel()

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

        Dim CATIA As INFITF.Application
        Try
            CATIA = GetObject(, "CATIA.Application")

        Catch ex As Exception
            MsgBox("The Application you seek" & vbCrLf & "Cannot be located." & vbCrLf & "Open a CATIA session.")
            Exit Sub
        End Try

        Try
            Dim ActiveDrawingDocument As DrawingDocument = CATIA.ActiveDocument
        Catch ex As Exception
            MsgBox("Rather than a beep" & vbCrLf & "Or a rude error message:" & vbCrLf & "Open a CATDrawing in the active session")
            Exit Sub
        End Try

        '*************************************************************
        '*************************************************************

        Dim Dwg As oDrawing = New oDrawing

        Dim what(0)
        what(0) = "DrawingTable"
        Dim UserSel2D As INFITF.Selection
        UserSel2D = CATIA.ActiveDocument.Selection
        UserSel2D.Clear()

        'Dim e As catbstr
        Dim e As String

        e = UserSel2D.SelectElement3(what, "Select a Product or a Component", False, CATMultiSelTriggWhenUserValidatesSelection, True)

        Dim SelectedTable As Integer
        Dim SelectedTableCollection As New Collection '(Of DrawingTable)
        Dim ActiveTable As DrawingTable

        Dim ListBox2D As New ListBox

        Try
            For SelectedTable = 1 To UserSel2D.Count
                SelectedTableCollection.Add(UserSel2D.Item(SelectedTable).Value)
                MaximumOfRowsInBigTable += SelectedTableCollection(SelectedTable).NumberOfrows
            Next

        Catch ex As Exception
            MsgBox("Make sure you select a proper Drawing Table")
        End Try

        MaximumOfRowsInBigTable += -1
        MaximumOfColumnsInBigTable = SelectedTableCollection(1).NumberOfColumns - 1


        ReDim Big2DTable(MaximumOfRowsInBigTable, MaximumOfColumnsInBigTable)



        Dim RowIndexOfBigTable As Integer = 0
        Dim ColumnIndexOfTable As Integer = 1


        Dim ItemNo = New oDrawing.Item
        Dim MatSpec = New oDrawing.MatlSpec
        Dim Nomenclature = New oDrawing.Nomenclature
        Dim PartNo = New oDrawing.PartNo
        Dim CageCode = New oDrawing.CageCode


        For SelectedTable = SelectedTableCollection.Count To 1 Step -1

            For RowIndexOfTable As Integer = 1 To SelectedTableCollection(SelectedTable).NumberOfRows
                ActiveTable = SelectedTableCollection(SelectedTable)

                For ColumnIndexOfTable = 1 To MaximumOfColumnsInBigTable + 1
                    Big2DTable(RowIndexOfBigTable, ColumnIndexOfTable - 1) = ActiveTable.GetCellString(RowIndexOfTable, ColumnIndexOfTable)
                Next

                RowIndexOfBigTable += 1
            Next
        Next

        ItemNo.Column = MaximumOfColumnsInBigTable
        MatSpec.Column = MaximumOfColumnsInBigTable - 1
        Nomenclature.Column = MaximumOfColumnsInBigTable - 2
        PartNo.Column = MaximumOfColumnsInBigTable - 3
        CageCode.Column = MaximumOfColumnsInBigTable - 4

        Dim PartsList As oDrawing.PartsList
        Dim ParentColumn As Integer
        Dim ParentDescriptionRow As Integer
        Dim QTY As String

       

        Children2D = New Dictionary(Of oDrawing.PartsList, String)

        For i = 0 To MaximumOfRowsInBigTable

            If IsNumeric(Big2DTable(i, ItemNo.Column)) = True Then
                PartsList = New oDrawing.PartsList

                PartsList.DrawingNumber = "B478004"
                PartsList.DrawingTitle = "BONDED STRUCTURE"
                PartsList.MatSpec = New Collection
                PartsList.ItemNo = CInt(Big2DTable(i, ItemNo.Column))
                PartsList.MatSpec.Add(Big2DTable(i, MatSpec.Column))
                PartsList.Nomenclature = Big2DTable(i, Nomenclature.Column)
                PartsList.PartNumber = Big2DTable(i, PartNo.Column)

                For ParentColumn = 0 To CageCode.Column - 1
                    If String.IsNullOrEmpty(Big2DTable(i, ParentColumn)) = False Then ' If the qty is not zero than add the item with the associated parent
                        PartsList.Parent = Big2DTable(MaximumOfRowsInBigTable - 1, ParentColumn)
                        For ParentDescriptionRow = 0 To MaximumOfRowsInBigTable
                            If PartsList.Parent = Big2DTable(ParentDescriptionRow, PartNo.Column) Then
                                PartsList.ParentDescription = Big2DTable(ParentDescriptionRow, Nomenclature.Column)
                            End If
                        Next
                        QTY = Big2DTable(i, ParentColumn)
                        Children2D.Add(PartsList, QTY)
                    End If
                Next
            End If
        Next

        Dim Real2Dchildrens = From Child2D In Children2D.AsParallel
        Select Child2D.Key.DrawingNumber + Child2D.Key.Parent
        Distinct

        For Each Real2DChild In Real2Dchildrens
            Available2DElements.Add(Real2DChild)
        Next



    End Sub
    Sub Write2DToExcel(Selected2DAssy As String, Available2DAssy As Integer)


        Dim oXL As Excel.Application

        'Dim oWB As Excel.Workbook
        'oWB = oXL.ActiveWorkbook

        'Dim oSheet As Excel.Worksheet
        ' oSheet = oWB.ActiveSheet


        Try
            oXL = GetObject(, "Excel.Application")
            ' oXL.Sheets(1).Cells.Clear()
        Catch ex As Exception
            oXL = CreateObject("Excel.Application")


        End Try

        oXL.DisplayAlerts = False
        oXL.Visible = True

        'Dim XLColumn As Integer = 5

        Dim XLRow As Integer = 14

        Dim Real2Dchildren = From Child2D In Children2D.AsParallel() _
        Where Child2D.Key.DrawingNumber + Child2D.Key.Parent = Selected2DAssy _
        Select Child2D.Value, Child2D.Key.PartNumber, Child2D.Key.Nomenclature, Child2D.Key.ParentDescription, Child2D.Key.ItemNo, Child2D.Key.MatSpec

        For Each Result2D In Real2Dchildren
            oXL.ActiveSheet.Cells(XLRow, 5).Value = Result2D.Value
            oXL.ActiveSheet.Cells(XLRow, 6).Value = Result2D.PartNumber
            oXL.ActiveSheet.Cells(XLRow, 7).Value = Result2D.Nomenclature
            oXL.ActiveSheet.Cells(XLRow, 8).Value = Result2D.ItemNo
            XLRow += 1
        Next



        oXL.DisplayAlerts = True
    End Sub
    Sub Write3Dto2D()

        Dim CATIA As INFITF.Application
        Try
            CATIA = GetObject(, "CATIA.Application")

        Catch ex As Exception
            MsgBox("The Application you seek" & vbCrLf & "Cannot be located." & vbCrLf & "Open a CATIA session.")
            Exit Sub
        End Try

        Dim ActiveDwgDocument As DrawingDocument

        Try
            ActiveDwgDocument = CATIA.ActiveDocument

        Catch ex As Exception
            MsgBox("Rather than a beep" & vbCrLf & "Or a rude error message:" & vbCrLf & "Open a CATDrawing in the active session")
            Exit Sub
        End Try

        Dim oDrwSheets As DrawingSheets
        oDrwSheets = ActiveDwgDocument.Sheets

        Dim oDrwSheet As DrawingSheet
        oDrwSheet = oDrwSheets.ActiveSheet

        Dim oDrwView As DrawingView
        oDrwView = oDrwSheet.Views.ActiveView

        'Retrieve the view's tables collection
        Dim oDrwTables As DrawingTables
        oDrwTables = oDrwView.Tables

        ' Create a new drawing table

        Dim oDrwTable As DrawingTable
        oDrwTable = oDrwTables.Add(658.5, 89, 20, 6, 5, 20)

        ' Set the drawing table's name
        odrwtable.Name = "Part List"

        odrwtable.SetColumnSize(odrwtable.NumberOfColumns, 17.8171)
        odrwtable.SetColumnSize(odrwtable.NumberOfColumns - 1, 43.18)
        odrwtable.SetColumnSize(odrwtable.NumberOfColumns - 2, 101.6)
        odrwtable.SetColumnSize(odrwtable.NumberOfColumns - 3, 63.5)
        odrwtable.SetColumnSize(odrwtable.NumberOfColumns - 4, 22.28)
        oDrwTable.SetColumnSize(oDrwTable.NumberOfColumns - 4, 22.28)


        Dim Realchildren = From child In Children3D.AsParallel() _
        Group child By child.partnumber, child.nomenclature Into Group _
        Select qty = Group.Count, partnumber = partnumber, nomenclature = nomenclature



        Dim i As Integer
        For i = 4 To oDrwTable.NumberOfRows - 1
            Dim ItemNo As String
            ItemNo = i - 3
            oDrwTable.SetCellString(oDrwTable.NumberOfRows - i, oDrwTable.NumberOfColumns, ItemNo)
            oDrwTable.SetCellAlignment(oDrwTable.NumberOfRows - i, oDrwTable.NumberOfColumns, DRAFTINGITF.CatTablePosition.CatTableMiddleCenter)
            oDrwTable.SetCellAlignment(oDrwTable.NumberOfRows - i, oDrwTable.NumberOfColumns - 5, DRAFTINGITF.CatTablePosition.CatTableMiddleCenter)
        Next

        i = oDrwTable.NumberOfRows - 4
        For Each result In Realchildren
            oDrwTable.SetCellString(i, oDrwTable.NumberOfColumns - 5, result.qty)
            oDrwTable.SetCellString(i, oDrwTable.NumberOfColumns - 3, result.partnumber)
            oDrwTable.SetCellString(i, oDrwTable.NumberOfColumns - 1, result.nomenclature)
            i -= 1
        Next

        'Title Block
        oDrwTable.SetCellString(oDrwTable.NumberOfRows - 1, oDrwTable.NumberOfColumns, "ITEM" & vbCrLf & "NO")
        oDrwTable.SetCellAlignment(oDrwTable.NumberOfRows - 1, oDrwTable.NumberOfColumns, DRAFTINGITF.CatTablePosition.CatTableMiddleCenter)
        oDrwTable.SetCellString(oDrwTable.NumberOfRows - 1, oDrwTable.NumberOfColumns - 1, "MATERIAL" & vbCrLf & "SPECIFICATION")
        oDrwTable.SetCellAlignment(oDrwTable.NumberOfRows - 1, oDrwTable.NumberOfColumns - 1, DRAFTINGITF.CatTablePosition.CatTableMiddleCenter)

        oDrwTable.SetCellString(oDrwTable.NumberOfRows - 1, oDrwTable.NumberOfColumns - 2, "NOMENCLATURE" & vbCrLf & "OR DESCRIPTION")
        oDrwTable.SetCellAlignment(oDrwTable.NumberOfRows - 1, oDrwTable.NumberOfColumns - 2, DRAFTINGITF.CatTablePosition.CatTableMiddleCenter)

        oDrwTable.SetCellString(oDrwTable.NumberOfRows - 1, oDrwTable.NumberOfColumns - 3, "PART OR" & vbCrLf & "IDENTIFYING NO.")
        oDrwTable.SetCellAlignment(oDrwTable.NumberOfRows - 1, oDrwTable.NumberOfColumns - 3, DRAFTINGITF.CatTablePosition.CatTableMiddleCenter)


        oDrwTable.SetCellString(oDrwTable.NumberOfRows - 1, oDrwTable.NumberOfColumns - 4, "CAGE" & vbCrLf & "CODE")
        oDrwTable.SetCellAlignment(oDrwTable.NumberOfRows - 1, oDrwTable.NumberOfColumns - 4, DRAFTINGITF.CatTablePosition.CatTableMiddleCenter)

        oDrwTable.SetCellString(oDrwTable.NumberOfRows, oDrwTable.NumberOfColumns - 5, "QTY" & vbCrLf & "REQD")
        oDrwTable.SetCellAlignment(oDrwTable.NumberOfRows, oDrwTable.NumberOfColumns - 5, DRAFTINGITF.CatTablePosition.CatTableMiddleCenter)

        oDrwTable.MergeCells(oDrwTable.NumberOfRows, oDrwTable.NumberOfColumns - 4, 1, 5)
        oDrwTable.SetCellString(oDrwTable.NumberOfRows, oDrwTable.NumberOfColumns - 4, "PARTS LIST")
        oDrwTable.SetCellAlignment(oDrwTable.NumberOfRows, oDrwTable.NumberOfColumns - 4, DRAFTINGITF.CatTablePosition.CatTableMiddleCenter)

        Dim SelectedTable As INFITF.Selection
        SelectedTable = CATIA.ActiveDocument.Selection
        SelectedTable.Clear()

        SelectedTable.Add(oDrwTable)
        SelectedTable.VisProperties.SetRealWidth(2, 1)

    End Sub
    Sub SelectXL()
        Dim oXL As New Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        Try
            oXL = GetObject(, "Excel.Application")
            ' oXL.Sheets(1).Cells.Clear()
        Catch ex As Exception
            'oXL = New Excel.Application
            'oXL = CreateObject("Excel.Application")
            MsgBox("Open an Excel Worksheet in order to make a selection")
        End Try

        ' oXL.DisplayAlerts = False
        'oXL.Visible = True

        oWB = oXL.ActiveWorkbook
        oSheet = oWB.ActiveSheet


        oSheet.Select()
        oSheet.Activate()

        'Task.Delay(5000)

        Dim Stradd, Endadd

        With oSheet.Selection
            Stradd = .Cells(1, 1).Address
            Endadd = .Cells(.Rows.Count, .Columns.Count).Address
        End With

        StartingXLRow = oSheet.Range(Stradd).Row
        StartingXLColumn = oSheet.Range(Stradd).Column
        EndXLRow = oSheet.Range(Endadd).Row
        EndXLColumn = oSheet.Range(Endadd).Column

        MsgBox(StartingXLColumn)
    End Sub

    Sub HTMLGenerator()
        Dim myTitle = "Hello this the comparator report"
        Dim Link As String = "D:\RHDSetup.log" '"http://www.w3schools.com"

        Dim D3HTML As XElement

        'For Each child In BOM3D
        'D3HTML = "<td>" & "1" & "</td><td><a href= " & Link & ">" & "CouldBeAVariable" & "</a></td>"
        ' For i = 1 To 5
        D3HTML =
<html>
    <table border="1" align="center">
        <tr><td>1</td><td><a href=<%= Link %>>CouldBeAVariable4</a></td></tr>
    </table>
</html>
        'D3HTML.AddAfterSelf(D3HTML)

        'Next
        ' Next



        Dim myHTML As XElement =
                    <html>
                        <head>
                            <title><%= myTitle %></title>
                        </head>
                        <body>
                            <h1>Welcome to my hood! Where the hood at?</h1>
                        </body>
                        <table border="1" align="center">
                            <tr><th>QTY</th><th>Part Number</th><th>Nomenclature</th></tr>
                            <tr><td>1</td><td><a href=<%= Link %>>CouldBeAVariable1</a></td></tr>
                        </table>
                    </html>

        Dim HTML2DReport As String


        HTML2DReport = "<html><head><title>" & myTitle & "</title> <style> table { border-width: 7px; border-style: outset; } </style></head>"     ' http://www.tizag.com/cssT/border.php
        HTML2DReport += "<body body bgcolor=""#f5f5dc""><h1 align=""center"">Welcome to my 2D hood! Where the hood at? </h1>"
        HTML2DReport += "<table border=""1"" border-style:groove align=""center""><tr><th>QTY</th><th>Part Number</th><th>Nomenclature</th><th>NPCF</th></tr>"

        For i = 86 To 134
            HTML2DReport += "<tr><td align=""center"">" & i & "</td><td><a href=" & Link & ">CouldBeAVariable1</a></td><td>" & i & "</td><td><input type=""button"" onclick=""window.location.href(" & "'http://www.google.com'" & " );"" value=""NPCF""></td></tr>"
        Next

        'HTML2DReport += "<tr><td>1</td><td><a href=" & Link & ">CouldBeAVariable1</a></td></tr>"
        HTML2DReport += "</table></body></html>"
       


        ' myHTML.Element("html").Element("body").Element("table").Element("tr").AddAfterSelf(D3HTML)

        Dim myHTMLafter As XElement = myHTML.<table>(0)
        'Dim BodyHTML As XElement =
        '        <body>
        '            <h1>Welcome to my hood! Where the hood at?</h1>
        '            <table border="1" align="center">
        '                <tr><th>QTY</th><th>Part Number</th><th>Nomenclature</th></tr>
        '                <tr><td>1</td><td><a href=<%= Link %>>CouldBeAVariable1</a></td></tr>
        '            </table>
        '        </body>

        'Dim TableHTML As XElement =
        '    <table border="1" align="center">
        '        <tr><th>QTY</th><th>Part Number</th><th>Nomenclature</th></tr>
        '        <tr><td>1</td><td><a href=<%= Link %>>CouldBeAVariable1</a></td></tr>
        '    </table>


        For i = 1 To 5
            D3HTML =
<html>
    <table border="1" align="center">
        <tr><td>1</td><td><a href=<%= Link %>>CouldBeAVariable<%= i %></a></td></tr>
    </table>
</html>

            myHTMLafter.AddAfterSelf(D3HTML)
        Next

        ' myHTMLafter.AddAfterSelf(From d3 In D3HTML.Elements() Where CInt(d3) = 4 Select d3)

        ' http://msdn.microsoft.com/en-us/library/system.xml.linq.xelement.addafterself(v=vs.110).aspx
        ' <%= D3HTML %>
        '<td>1</td><td><a href=<%= Link %>>CouldBeAVariable</a></td>
        Using writer As StreamWriter = New StreamWriter("TheComparator.html")
            writer.Write(HTML2DReport)
        End Using
    End Sub


    Sub XLto2D()

    End Sub

    Sub Write3DtoHTML()

        Dim myTitle = "Hello this the comparator report"
        Dim Link As String = "D:\RHDSetup.log" '"http://www.w3schools.com"

        Dim D3HTML As XElement

        D3HTML =
<html>
    <table border="1" align="center">
        <tr><td>1</td><td><a href=<%= Link %>>CouldBeAVariable4</a></td></tr>
    </table>
</html>

        Dim myHTML As XElement =
                    <html>
                        <head>
                            <title><%= myTitle %></title>
                        </head>
                        <body>
                            <h1>Welcome 3D to HTML report! Where the hood at?</h1>
                        </body>
                        <table border="1" align="center">
                            <tr><th>QTY</th><th>Part Number</th><th>Nomenclature</th></tr>
                            <tr><td>1</td><td><a href=<%= Link %>>CouldBeAVariable1</a></td></tr>
                        </table>
                    </html>




        Dim myHTMLafter As XElement = myHTML.<table>(0)
       
        Dim Realchildrens = From childs In Children3D.AsParallel() _
        Group childs By childs.partnumber, childs.nomenclature Into Group _
        Select qty = Group.Count, partnumber = partnumber, nomenclature = nomenclature

        For Each results In Realchildrens

            D3HTML =
<html>
    <table border="1" align="center">
        <tr><td><%= results.qty %></td><td><a href=<%= Link %>><%= results.partnumber %></a></td></tr>
    </table>
</html>

            myHTMLafter.AddAfterSelf(D3HTML)
        Next


        ' http://msdn.microsoft.com/en-us/library/system.xml.linq.xelement.addafterself(v=vs.110).aspx
        ' <%= D3HTML %>
        '<td>1</td><td><a href=<%= Link %>>CouldBeAVariable</a></td>
        Using writer As StreamWriter = New StreamWriter("TheComparator.html")
            writer.Write(myHTML)
        End Using



    End Sub
    Sub XLtoHTML()
    End Sub

    Sub Wrtite3Dvs2DtoHTML(Selected2DAssy As String)


        Dim PartNo3DChildren = From PartNo3DChild In Children3D _
        Group PartNo3DChild By PartNo3DChild.partnumber, PartNo3DChild.nomenclature Into Group _
        Select partnumber = partnumber

        Dim PartNo2DChildren = From PartNo2DChild In Children2D _
        Where PartNo2DChild.Key.DrawingNumber + PartNo2DChild.Key.Parent = Selected2DAssy _
        Select PartNo2DChild.Key.PartNumber

        Dim PartNoCompare = PartNo3DChildren.Intersect(PartNo2DChildren)
        For Each part In PartNoCompare
            MsgBox(part)
        Next


        Dim PartNoAndQty3DChildren = From PartNoAndQty3DChild In Children3D _
        Group PartNoAndQty3DChild By PartNoAndQty3DChild.partnumber Into Group _
        Select Value = Group.Count.ToString, partnumber = partnumber

        Dim PartNoAndQty2DChildren = From PartNoAndQty2DChild In Children2D _
        Where PartNoAndQty2DChild.Key.DrawingNumber + PartNoAndQty2DChild.Key.Parent = Selected2DAssy _
        Select PartNoAndQty2DChild.Value, PartNoAndQty2DChild.Key.PartNumber

        Dim PartNoAndQtyCompare = PartNoAndQty2DChildren.Except(PartNoAndQty3DChildren)

            For Each PartAndQtyChild In PartNoAndQtyCompare
                ' MsgBox(PartAndQtyChild.partnumber)
                MsgBox(PartAndQtyChild)
            Next

        'Catch ex As Exception

        If PartNoAndQty2DChildren.Count = 0 Then
            MsgBox("There is No 2D")
        End If

        If PartNoAndQty3DChildren.Count = 0 Then
            MsgBox("There is No 3D")
        End If

        MsgBox("Can't do it!!")
        'End Try

        ' Dim Real3Dchildren = From child In Children3D _
        ' Group child By child.partnumber, child.nomenclature Into Group _
        ' Select qty = Group.Count.ToString, partnumber = partnumber, nomenclature = nomenclature

        ' Dim Real2Dchildren = From Child2D In Children2D _
        ' Where Child2D.Key.DrawingNumber + Child2D.Key.Parent = Selected2DAssy _
        ' Select Child2D.Value, Child2D.Key.PartNumber, Child2D.Key.Nomenclature


    End Sub


    Sub Wrtite3Dvs2DtoXL(Selected2DAssy As String)

        Dim oXL As Excel.Application = Nothing
        Dim oWB As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing


        ' Start Excel and get Application object.
        oXL = CreateObject("Excel.Application")
        oXL.Visible = True

        ' Get a new workbook.
        'oWB=oXL.Workbooks.Add 
        oWB = oXL.Workbooks.Open("C:\Users\Al\Dropbox\ESS\The Comparator\The Comparator latest.xlsm", UpdateLinks:=0, ReadOnly:=True)
        oSheet = oWB.ActiveSheet



        '3D Excel export
        Dim XL3DRow As Integer = 14

        Dim Realchildren3D = From Child3D In Children3D.AsParallel() _
        Group Child3D By Child3D.partnumber, Child3D.nomenclature Into Group _
        Select qty = Group.Count, partnumber = partnumber, nomenclature = nomenclature

        For Each Result3D In Realchildren3D
            oXL.ActiveSheet.Cells(XL3DRow, 1).Value = Result3D.qty
            oXL.ActiveSheet.Cells(XL3DRow, 2).Value = Result3D.partnumber
            oXL.ActiveSheet.Cells(XL3DRow, 3).Value = Result3D.nomenclature
            XL3DRow += 1
        Next

        '2D Excel export
        Dim XL2DRow As Integer = 14

        Dim RealChildren2D = From Child2D In Children2D.AsParallel() _
        Where Child2D.Key.DrawingNumber + Child2D.Key.Parent = Selected2DAssy _
        Select Child2D.Value, Child2D.Key.PartNumber, Child2D.Key.Nomenclature, Child2D.Key.ParentDescription, Child2D.Key.ItemNo, Child2D.Key.MatSpec

        For Each Result2D In RealChildren2D
            oXL.ActiveSheet.Cells(XL2DRow, 5).Value = Result2D.Value
            oXL.ActiveSheet.Cells(XL2DRow, 6).Value = Result2D.PartNumber
            oXL.ActiveSheet.Cells(XL2DRow, 7).Value = Result2D.Nomenclature
            oXL.ActiveSheet.Cells(XL2DRow, 9).Value = Result2D.ItemNo
            XL2DRow += 1
        Next
    End Sub

    Sub Write2DtoHTML()

    End Sub

    Sub Is3DPartIn2D()

    End Sub
    Sub Is2DPartIn3D()

    End Sub
    Sub Is3DQtyEquals2DQty()

    End Sub
    Sub Is3DNomenclatureSameAs2D()

    End Sub
  
End Class
