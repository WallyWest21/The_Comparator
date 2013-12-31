
Option Explicit On

'Imports INFITF
Imports MECMOD 'Mechanical Modeler & Sketcher 
Imports PARTITF 'Part Design features ex: Pad, Split, Sweep
Imports KnowledgewareTypeLib
Imports HybridShapeTypeLib
Imports ProductStructureTypeLib

'Imports CATInstantCollabItf
'Imports INFITF
'Imports SPATypeLib

Imports System.Collections

Imports System
Imports System.Collections.ObjectModel

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
    Public Property lst1 As New ObservableCollection(Of String)
    Public Property TheRealChildren As New ObservableCollection(Of String)
    'Public Sub New()
    '    ' This call is required by the designer.
    '    InitializeComponent()

    '    'lst1.Add("one")
    '    'lst1.Add("two")
    '    'lst1.Add("three")

    '    'Dim comp As New Comparator
    '    '  ListBox1.ItemsSource = lst1

    '    ' Add any initialization after the InitializeComponent() call.
    '    '  ListBox1.ItemsSource = lst1

    '    '  Me.DataContext = Me

    'End Sub


    
    Private Sub HTML_Label_MouseDown(sender As Object, e As MouseButtonEventArgs)

    End Sub

    Private Sub HTMLLabel_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles HTMLLabel.MouseDown

        ' Dim comp As ChildrenList

        '    lst1.Add("New Item")

        'ListBox1.ItemsSource = ChildrenList.



    End Sub
End Class
