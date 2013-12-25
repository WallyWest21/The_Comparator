
Option Explicit On

'Imports INFITF
Imports MECMOD
Imports PARTITF
Imports KnowledgewareTypeLib
Imports HybridShapeTypeLib
Imports ProductStructureTypeLib
Imports System

Class MainWindow
    Private Sub Label_MouseDown_2(sender As Object, e As MouseButtonEventArgs)
        Dim Comparator As New Comparator

        Call Comparator.Select3D()

        '******************************************
        'Dim CATIA As Object
        'CATIA = GetObject(, "CATIA.Application")

        ''Get the current CATIA assembly

        'Dim oProdDoc As ProductDocument
        'oProdDoc = CATIA.ActiveDocument

        'Dim oRootProd As Products
        'oRootProd = oProdDoc.Product

        ''Dim Children As New Collection(Of Object)
        ''Children = New Collection

        'MsgBox("This is a CATPart with part number " & oRootProd.PartNumber)

        'Call Comparator.WalkDownTree(oRootProd)
        'Call Comparator.WriteToExcel()

        'MsgBox("Done " & Comparator.Children(Comparator.Children.Count).partnumber)


        '***************************************************************************

    End Sub

    Private Sub _2DLabel_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles _2DLabel.MouseDown
        '    Dim myCatia As INFITF.Application

        '    Try

        '        myCatia = GetObject(, "CATIA.application")

        '    Catch ex As Exception

        '        myCatia = CreateObject("CATIA.application")

        '    End Try

        '    myCatia.Visible = True

        '    myCatia.DisplayFileAlerts = True

        '    Dim mPartDoc As MECMOD.PartDocument

        '    Dim mPart As MECMOD.Part

        '    Try

        '        mPartDoc = myCatia.ActiveDocument

        '        mPart = mPartDoc.Part

        '    Catch ex As Exception

        '        MsgBox("there was no active part", MsgBoxStyle.Critical)

        '    End Try
        'MsgBox("there was no active part")

        Dim Comparator As New Comparator
        Call Comparator.Select2D()


    End Sub

    Private Sub ListBox_DragEnter(sender As Object, e As DragEventArgs)

    End Sub

    Private Sub _2DLabel_Drop(sender As Object, e As DragEventArgs) Handles _2DLabel.Drop
        _2DLabel.Content = "OK"
    End Sub

    

   
End Class
