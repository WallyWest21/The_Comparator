Imports DRAFTINGITF

Public Class oDrawing

    Public Code As String
    Public No As String
    Public Title As String
    Public Revision As String
    Public SheetNo As Integer
    Public SheetCount As Integer
    Public Quantity As Integer
    Public Number As String
    Public QtyReqdRows(,) As Integer
    Public FirstRowofItems(,) As Integer


    Public ParentOf2DAssemblies As New Collection
    Public Function RemoveAssembliesRow(ByVal DrawingTable As DrawingTable, ByVal RowIndexOfTable As Integer, ByVal ColumnIndexOfTable As Integer) As Boolean
        Return False
        If CInt(Right(DrawingTable.GetCellString(RowIndexOfTable, ColumnIndexOfTable), 3)) / 500 >= 1 Then
            Return True
        End If
    End Function
    Public Sub RemoveCageCodeColumn(MaximumOfRowsInBigTable As Integer, MaximumOfColumnsInBigTable As Integer, IsEnabled As Boolean)

    End Sub
    Public Sub RemoveTableFooter(MaximumOfRowsInBigTable As Integer, MaximumOfColumnsInBigTable As Integer, IsEnabled As Boolean)

    End Sub
    Public Sub RemoveEmptyItems(MaximumOfRowsInBigTable As Integer, MaximumOfColumnsInBigTable As Integer, IsEnabled As Boolean)

    End Sub
    Public Sub RemoveQtyReqdRow(MaximumOfRowsInBigTable As Integer, MaximumOfColumnsInBigTable As Integer, IsEnabled As Boolean)

    End Sub
    Public Sub RemoveAssembliesColumns(MaximumOfRowsInBigTable As Integer, MaximumOfColumnsInBigTable As Integer, IsEnabled As Boolean)

    End Sub

    Public Function IsAssembliesColumnSelected(ByVal ActiveTable As DrawingTable, ByVal RowIndexOfTable As Integer, ByVal PartNoColumn As Integer, ByVal SelectedTable As Integer) As Boolean


        Dim Assy As Integer
        Try

            If IsNumeric(Right(Trim(ActiveTable.GetCellString(RowIndexOfTable, PartNoColumn)), 3)) And Left(Trim(ActiveTable.GetCellString(RowIndexOfTable, PartNoColumn)), 1).Contains("-") = True Then 'And SelectedTable > 1 = True Then

                If CInt(Right(Trim(ActiveTable.GetCellString(RowIndexOfTable, PartNoColumn)), 3)) / 500 >= 1 Then
                    Assy = CInt(Right(Trim(ActiveTable.GetCellString(RowIndexOfTable, PartNoColumn)), 3))
                    ' MsgBox((Trim(ActiveTable.GetCellString(RowIndexOfTable, PartNoColumn))) & "    " & Assy)
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox("Wrong Table Format " & ((Trim(ActiveTable.GetCellString(RowIndexOfTable, PartNoColumn)))))
        End Try


    End Function
    Public Function Available2DElements() As String
        Available2DElements = "No Availaible 2D assemblies"
        Return Available2DElements
    End Function
    Public Class Item

        ' Dim Drawig = New Drawing
        Public Column As Integer
        Public Value As Integer
    End Class
    Public Class Nomenclature
        Public Column As Integer

        ''' <summary>
        ''' The description of the nomenclature that would be compared on the 3D.  A  spell check should performed to validate quality.
        ''' </summary>
        Public Description As String

    End Class
    Public Class MatlSpec
        Public Column As Integer
        ''' <summary>
        ''' The content of the MatSpec: it could be the manufacturer or the flag notes. In the case of the flag notes, numbers would require a space otherwise it will give wrong flag notes.
        ''' </summary>
        Public Description As String

    End Class
    Public Class PartNo
        Public Column As Integer
        ''' <summary>
        ''' The content of the PartNo: it could be the dash number or an external child. The drawing prefix should be added in case it is an internal child.
        ''' </summary>
        Public Description As String

    End Class
    Public Class CageCode
        Public Column As Integer

    End Class

    Public Class Notes
        Public Class GeneralNotes
        End Class
        Public Class FlagNotes
        End Class
    End Class
    Function Clean2DTable(DrawingTable As DrawingTable) As DrawingTable
        Return DrawingTable
    End Function

    Public Class Children2D
        Public ItemNo As Integer
        Public Nomenclature As String
        Public PartNo As String
        Public Parent As Collection
        Public DrawingNo As String
        Public MatSpec As New Collection
        Public Quantity As Integer
    End Class

    Public Structure PartsList
        Public Parent As String
        Public ParentDescription As String
        Public PartNumber As String
        Public Nomenclature As String
        Public MatSpec As Collection
        Public ItemNo As Integer
        Public DrawingNumber As String
        Public DrawingTitle As String
    End Structure

    Public Class TFList
        '  Inherits PartsList
        Public MatlCode As String

    End Class
End Class


