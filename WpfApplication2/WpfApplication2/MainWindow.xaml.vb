
Option Explicit On

'Imports INFITF
Imports MECMOD      'Mechanical Modeler & Sketcher 
Imports PARTITF     'Part Design features ex: Pad, Split, Sweep
Imports KnowledgewareTypeLib
Imports HybridShapeTypeLib
Imports ProductStructureTypeLib

'Imports CATInstantCollabItf

'Imports SPATypeLib

Imports System.Collections

Imports System
Imports System.Collections.ObjectModel
Imports INFITF

Public Class MainWindow


    Public Is2DSelected As Boolean
    Public Is3DSelected As Boolean



    Private Sub Label_MouseDown_2(sender As Object, e As MouseButtonEventArgs)


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
    Public Property UpdateSourceTrigger As UpdateSourceTrigger
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
        Dim comp As ChildrenList
        Dim comparator As Comparator


        '    lst1.Add("New Item")

        'ListBox1.ItemsSource = ChildrenList.
        ' comp.Add("456")
        ' comp.Add(comparator.Realchildren3D(1))

        ListBox3D.ItemsSource = comp

    End Sub
    Public Sub SelectionEvents(UIElements As Object, IsSelected As Boolean)
        ' Dim Is2DSelected As Boolean
        'Is2DSelected = Not Is2DSelected

        If IsSelected = True Then
            UIElements.BorderThickness = New System.Windows.Thickness(5)
            UIElements.BorderBrush = New SolidColorBrush(Colors.LightGreen)

        Else
            UIElements.BorderThickness = New System.Windows.Thickness(0)
            UIElements.BorderBrush = _2DLabel.Background
        End If
    End Sub







    Private Sub Label_Drop(sender As Object, e As DragEventArgs)

        'Dim theFiles() As String = CType(e.Data.GetData("FileDrop", True), String())
        'For Each theFile As String In theFiles
        '    MsgBox(theFile)
        'Next

        'Dim files As Object = e.Data.GetData("Object")

        Dim CATIA As Object
        CATIA = GetObject(, "CATIA.Application")


        Dim theFiles = e.Data.GetDataPresent("Part", True)
        'For Each theFile As String In theFiles
        Try
            MsgBox(theFiles)


        Catch ex As Exception
            MsgBox("It is empty!")
        End Try

        ' Next


        '' For Each path In files
        'MsgBox(files.Name)
        ''Next

        'Dim theFiles() As String = CType(e.Data.GetData("FileDrop", True), String())
        ''  For Each theFile As String In theFiles
        'MsgBox(theFiles)
        'Next

    End Sub



    Private Sub _2DLabel_MouseRightButtonDown(sender As Object, e As MouseButtonEventArgs) Handles _2DLabel.MouseRightButtonDown
        Is2DSelected = False
        Call SelectionEvents(_2DLabel, Is2DSelected)
        ListBox2D.ItemsSource = Nothing
        ListBox2D.Items.Clear()
        '  ListBox2D.Items.Remove(ListBox2D.SelectedIndex)
        ' ListBox2D.Items.Refresh()

    End Sub

    Private Sub _2DLabel_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles _2DLabel.MouseLeftButtonDown
        Is2DSelected = True


        Call SelectionEvents(_2DLabel, Is2DSelected)
        Dim Comparator As New Comparator
        Call Comparator.Select2D()
        ListBox2D.ItemsSource = Comparator.Available2DElements
    End Sub

    Public Sub _3DLabel_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles _3DLabel.MouseLeftButtonDown

        Is3DSelected = True

        Call SelectionEvents(_3DLabel, Is3DSelected)

        Dim Comparator As New Comparator
        Call Comparator.Select3D()
        Call SelectionEvents(_3DLabel, Is3DSelected)

        ListBox3D.ItemsSource = Comparator.Selected3DElements
        DataGrid1.ItemsSource = Comparator.Realchildren3D
        DataGrid1.DataContext = Comparator.Children3D

        ' ListBox1.ItemsSource = Comparator.Realchildren3D
    End Sub

    Private Sub _3DLabel_MouseRightButtonDown(sender As Object, e As MouseButtonEventArgs) Handles _3DLabel.MouseRightButtonDown
        Is3DSelected = False
        Call SelectionEvents(_3DLabel, Is2DSelected)
        ListBox3D.ItemsSource = Nothing
        ListBox3D.Items.Clear()
    End Sub

    Private Sub Image_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        For Each item In ListBox2D.SelectedItems
            MsgBox(item.ToString())

        Next
        
    End Sub
End Class
