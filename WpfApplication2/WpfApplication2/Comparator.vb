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
    Public ReadOnly Children2D As New Collection

    Public Property Realchildren3D As New ObservableCollection(Of String)
    Public Property Selected3DElements As New ObservableCollection(Of String)
    Public Property Available2DElements As New ObservableCollection(Of String)
    Public Property Selected2DElements As New ObservableCollection(Of String)
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
            oXL.Sheets(1).Cells(i + 3, 1).Value = result.qty
            oXL.Sheets(1).Cells(i + 3, 2).Value = result.partnumber
            oXL.Sheets(1).Cells(i + 3, 3).Value = result.nomenclature

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



        Dim oXL As Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet



        Try
            oXL = GetObject(, "Excel.Application")
            ' oXL.Sheets(1).Cells.Clear()
        Catch ex As Exception
            oXL = New Excel.Application


        End Try

        ' oXL.DisplayAlerts = False
        oXL.Visible = True




        Dim Dwg As oDrawing = New oDrawing

        Dim what(0)
        what(0) = "DrawingTable"
        Dim UserSel2D As INFITF.Selection
        UserSel2D = CATIA.ActiveDocument.Selection
        UserSel2D.Clear()

        'Dim e As catbstr
        Dim e As String

        e = UserSel2D.SelectElement3(what, "Select a Product or a Component", False, CATMultiSelTriggWhenUserValidatesSelection, True)

        'Dim MaximumOfColumnsInBigTable As Integer = 0
        'Dim MaximumOfRowsInBigTable As Integer = 0

        'MaximumOfColumnsInBigTable = 0
        'MaximumOfRowsInBigTable = 0


        Dim SelectedTable As Integer
        Dim SelectedTableCollection As New Collection '(Of DrawingTable)
        Dim ActiveTable As DrawingTable


        Dim ListBox2D As New ListBox

        Try



            For SelectedTable = 1 To UserSel2D.Count

                SelectedTableCollection.Add(UserSel2D.Item(SelectedTable).Value)
                ' ActiveTable = SelectedTableCollection(SelectedTable)

                'MaximumOfRowsInTable += ActiveTable.NumberOfRows
                MaximumOfRowsInBigTable += SelectedTableCollection(SelectedTable).NumberOfrows
            Next
        Catch ex As Exception
            MsgBox("Make sure you select a proper Drawing Table")
        End Try
        ' MsgBox(MaximumOfRowsInTable)
        MaximumOfRowsInBigTable += -1
        MaximumOfColumnsInBigTable = SelectedTableCollection(1).NumberOfColumns - 1


        ReDim Big2DTable(MaximumOfRowsInBigTable, MaximumOfColumnsInBigTable)
        ' Dim Big2DTable(MaximumOfRowsInBigTable, MaximumOfColumnsInBigTable) As String


        Dim RowIndexOfBigTable As Integer = 0
        Dim ColumnIndexOfTable As Integer = 1


        Dim ItemNo = New oDrawing.Item
        Dim MatSpec = New oDrawing.MatlSpec
        Dim Nomenclature = New oDrawing.Nomenclature
        Dim PartNo = New oDrawing.PartNo
        Dim CageCode = New oDrawing.CageCode

        ItemNo.Column = MaximumOfColumnsInBigTable + 1
        MatSpec.Column = MaximumOfColumnsInBigTable + 1 - 1
        Nomenclature.Column = MaximumOfColumnsInBigTable + 1 - 2
        PartNo.Column = MaximumOfColumnsInBigTable + 1 - 3
        CageCode.Column = MaximumOfColumnsInBigTable + 1 - 3

        Dwg.Code = "Dummy"




        For SelectedTable = SelectedTableCollection.Count To 1 Step -1


            For RowIndexOfTable As Integer = 1 To SelectedTableCollection(SelectedTable).NumberOfRows
                ActiveTable = SelectedTableCollection(SelectedTable)

                ColumnIndexOfTable = 1
                If (ActiveTable.GetCellString(RowIndexOfTable, ColumnIndexOfTable)).Contains("QTY") = True And SelectedTable > 1 Then
                    Continue For
                End If

                If ActiveTable.GetCellString(RowIndexOfTable, Nomenclature.Column).Contains("NOMENCLATURE") = True And SelectedTable > 1 Then
                    Continue For
                End If

                If Left((ActiveTable.GetCellString(RowIndexOfTable, PartNo.Column)), 2).Contains("-5") = True And SelectedTable > 1 Then
                    Dwg.ParentOf2DAssemblies.Add(ActiveTable.GetCellString(RowIndexOfTable, PartNo.Column))
                    Available2DElements.Add(Dwg.Code & ActiveTable.GetCellString(RowIndexOfTable, PartNo.Column))
                    Continue For
                End If

                For ColumnIndexOfTable = 1 To MaximumOfColumnsInBigTable + 1
                    Big2DTable(RowIndexOfBigTable, ColumnIndexOfTable - 1) = ActiveTable.GetCellString(RowIndexOfTable, ColumnIndexOfTable)
                Next
                RowIndexOfBigTable += 1
            Next
        Next

        'For j = 0 To MaximumOfColumnsInBigTable

        '    For i = 0 To MaximumOfRowsInBigTable
        '        oXL.ActiveSheet.Cells(i + 13, j + 1) = Big2DTable(i, j)
        '        oXL.ActiveSheet.Cells(i + 13, j + 1).wraptext = True


        '        If j = MaximumOfColumnsInBigTable - 3 Then
        '            oXL.ActiveSheet.Cells(i + 13, j + 1).columnwidth = 15
        '        End If

        '        If j = MaximumOfColumnsInBigTable - 2 Then
        '            oXL.ActiveSheet.Cells(i + 13, j + 1).columnwidth = 35

        '        End If


        '    Next i

        'Next


        'Call Write2DToExcel(MaximumOfColumnsInBigTable, MaximumOfRowsInBigTable, Big2DTable)

    End Sub


    'Sub Write2DToExcel()
    Sub Write2DToExcel(Selected2DAssy As Integer, Available2DAssy As Integer)


        Dim oXL As Excel.Application
        Dim oWB As Excel.Workbook
        'Dim oSheet As Excel.Worksheet



        Try
            oXL = GetObject(, "Excel.Application")
            ' oXL.Sheets(1).Cells.Clear()
        Catch ex As Exception
            oXL = New Excel.Application


        End Try

        ' oXL.DisplayAlerts = False
        oXL.Visible = True
        '  Dim Selected2DAssy As Integer

        For j As Integer = Selected2DAssy To MaximumOfColumnsInBigTable

            If j = Selected2DAssy Or j > MaximumOfColumnsInBigTable - Available2DAssy Then

                For i As Integer = 0 To MaximumOfRowsInBigTable
                    oXL.ActiveSheet.Cells(i + 13, j + 1) = Big2DTable(i, j)
                    oXL.ActiveSheet.Cells(i + 13, j + 1).wraptext = True


                    If j = MaximumOfColumnsInBigTable - 3 Then
                        oXL.ActiveSheet.Cells(i + 13, j + 1).columnwidth = 15
                    End If

                    If j = MaximumOfColumnsInBigTable - 2 Then
                        oXL.ActiveSheet.Cells(i + 13, j + 1).columnwidth = 35

                    End If

                Next i
            End If
        Next
    End Sub

    Sub Is3DPartIn2D()

    End Sub
    Sub Is2DPartIn3D()

    End Sub
    Sub Is3DQtyEquals2DQty()

    End Sub
    Sub Is3DNomenclatureSameAs2D()

    End Sub
    Public Class ChildrenList1
        Inherits ObservableCollection(Of String)
        ' Implements INotifyPropertyChanged
        ' Public Property pChildrenList As New ObservableCollection(Of String)
        ' Inherits ObservableCollection(Of Object)
        Public Sub New()

            '  Dim Item

            ' Dim Comparator As New Comparator
            ' For Each Item In Comparator.Children3D
            'For Item = 1 To 25
            MyBase.Add("nkhgkghbkhgkhklhgbkjhgjkhgkhfvk")
            MyBase.Add("kl;njkbjfkuigkhjklk")
            ' Next

            MyBase.Add("second tolast")
            MyBase.Add("kl;lsdfgd Super Last1")
        End Sub



    End Class
End Class
