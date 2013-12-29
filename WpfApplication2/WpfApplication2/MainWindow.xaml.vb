
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

    End Sub

    Private Sub _2DLabel_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles _2DLabel.MouseDown

        Dim Comparator As New Comparator
        Call Comparator.Select2D()

    End Sub

    Private Sub ListBox_DragEnter(sender As Object, e As DragEventArgs)

    End Sub

    Private Sub _2DLabel_Drop(sender As Object, e As DragEventArgs) Handles _2DLabel.Drop
        _2DLabel.Content = "OK"
    End Sub

    
    Private Sub Window_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Left Then
            'If e1.Key = Key.Right Then
            _2DLabel.Content = "OG"
            'End If

        End If

    End Sub
End Class
