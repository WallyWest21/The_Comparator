
Option Explicit On

'Imports INFITF
Imports MECMOD      'Mechanical Modeler & Sketcher 
Imports PARTITF     'Part Design features ex: Pad, Split, Sweep
Imports KnowledgewareTypeLib
Imports HybridShapeTypeLib
Imports ProductStructureTypeLib

'Imports CATInstantCollabItf

'Imports SPATypeLib

'Imports System.Collections

'Imports System
Imports System.Collections.ObjectModel
Imports INFITF
Imports System.Windows.Media.Animation

Public Class MainWindow


    Private myStoryboard As Storyboard

    'Public Sub New()

    '    ' This call is required by the designer.
    '    InitializeComponent()
    '    '*******************Animation*************************

    '    Dim myDoubleAnimation As New DoubleAnimation()

    '    myDoubleAnimation.From = ComparatorWindow.Height
    '    myDoubleAnimation.To = ComparatorWindow.Height + 150
    '    myDoubleAnimation.Duration = New Duration(TimeSpan.FromSeconds(5))
    '    myDoubleAnimation.AutoReverse = False
    '    myDoubleAnimation.RepeatBehavior = RepeatBehavior.Forever

    '    myStoryboard = New Storyboard()
    '    myStoryboard.Children.Add(myDoubleAnimation)
    '    Storyboard.SetTargetName(myDoubleAnimation, ComparatorWindow.Name)
    '    Storyboard.SetTargetProperty(myDoubleAnimation, New PropertyPath(Grid1.Height))

    '    AddHandler ComparatorWindow.Loaded, AddressOf ComparatorWindow_Loaded
    '    Me.Content = ComparatorWindow



    '    '*******************Animation*************************
    '    ' Add any initialization after the InitializeComponent() call.

    'End Sub

    Public Is2DSelected As Boolean = False
    Public Is3DSelected As Boolean = False
    Public IsXLSelected As Boolean = False

    Dim Comparator As New Comparator

    'Public Shared MaximumOfColumnsInBigTable As Integer = 0
    'Public Shared MaximumOfRowsInBigTable As Integer = 0
    'Public Shared Big2DTable(MaximumOfRowsInBigTable, MaximumOfColumnsInBigTable) As String




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
        'Dim comp As ChildrenList
        'Dim comparator As Comparator


        ''    lst1.Add("New Item")

        ''ListBox1.ItemsSource = ChildrenList.
        '' comp.Add("456")
        '' comp.Add(comparator.Realchildren3D(1))

        'ListBox3D.ItemsSource = comp

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
        SelectionEvents(_2DLabel, Is2DSelected)
        ListBox2D.ItemsSource = Nothing
        ListBox2D.Items.Clear()
        '  ListBox2D.Items.Remove(ListBox2D.SelectedIndex)
        ' ListBox2D.Items.Refresh()

    End Sub
    Private Sub _2DLabel_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles _2DLabel.MouseLeftButtonDown
        Is2DSelected = True

        SelectionEvents(_2DLabel, Is2DSelected)
        'Dim Comparator As New Comparator
        Comparator.Select2D()

        ListBox2D.ItemsSource = Comparator.Available2DElements

        ' Await Task.Delay(30000)


        ' Is2DSelected = False
        'SelectionEvents(_2DLabel, Is2DSelected)

        Exit Sub
    End Sub
    Public Sub _3DLabel_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles _3DLabel.MouseLeftButtonDown

        Is3DSelected = True

        'MsgBox(4 ^ Convert.ToUInt32(Is3DSelected))

        'For counter = 0 To 155 Step 0.001

        '    Grid1.RowDefinitions(4).Height = New GridLength(counter)

        'Next

        Call SelectionEvents(_3DLabel, Is3DSelected)

        'Dim Comparator As New Comparator
        Call Comparator.Select3D()
        Call SelectionEvents(_3DLabel, Is3DSelected)



        'myStoryboard.Begin(Me)



        ListBox3D.ItemsSource = Comparator.Selected3DElements
        DataGrid1.ItemsSource = Comparator.Realchildren3D
        DataGrid1.DataContext = Comparator.Children3D
        ListBox3D.SelectAll()









        ' ListBox1.ItemsSource = Comparator.Realchildren3D
    End Sub
    Private Sub _3DLabel_MouseRightButtonDown(sender As Object, e As MouseButtonEventArgs) Handles _3DLabel.MouseRightButtonDown
        Is3DSelected = False
        Call SelectionEvents(_3DLabel, Is2DSelected)
        ListBox3D.ItemsSource = Nothing
        ListBox3D.Items.Clear()
    End Sub
    Private Sub Image_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        'Dim Comparator As New Comparator

        For Each item In ListBox2D.SelectedItems
            MsgBox(item.ToString())
        Next

        MsgBox(ListBox2D.Items.Count)

        Call Comparator.Write2DToExcel(ListBox2D.SelectedIndex, ListBox2D.Items.Count)
    End Sub
    Private Sub XLLabelOutput_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles XLLabelOutput.MouseLeftButtonDown

        MsgBox("3D: " & Is3DSelected & "2D: " & Is2DSelected & InputSum())

        Select Case InputSum()

            Case 0
                MsgBox("Choose at least one input")
            Case 2
                Call Comparator.Write2DToExcel(ListBox2D.SelectedIndex, ListBox2D.Items.Count)
            Case 4
                Call Comparator.Write3DToExcel()
            Case 6
                Call Comparator.Write3DToExcel()
                Call Comparator.Write2DToExcel(ListBox2D.SelectedIndex, ListBox2D.Items.Count)
        End Select

        Call Comparator.Write2DToExcel(ListBox2D.SelectedIndex, ListBox2D.Items.Count)
    End Sub
    Private Sub ComparatorWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles ComparatorWindow.Loaded

        'For row As Integer = 4 To 2 Step -1
        '    Grid1.RowDefinitions(row).Height = New GridLength(0, GridUnitType.Star)
        'Next
        'ComparatorWindow.Height = 335

        Dim CATIA As Object
        Try
            CATIA = GetObject(, "CATIA.Application")

        Dim oDocuments As Documents
        oDocuments = CATIA.Documents

        Dim oDocument As Document

            Dim strType As String

        Dim AvailableDocsPartNo() As String
        For Each oDocument In oDocuments
            strType = TypeName(oDocument)

            Select Case strType
                Case "ProductDocument", "DrawingDocument"

                    AvailableDocsPartNo = oDocument.Name.Split(".")
                    Comparator.ActiveDocuments.Add(AvailableDocsPartNo(0))
                  
            End Select
        Next
        ListBoxDocs.ItemsSource = Comparator.ActiveDocuments
        Catch ex As Exception

        End Try
    End Sub
    Function InputSum() As Integer

        InputSum = (4 * Convert.ToUInt32(Is3DSelected)) + (2 * Convert.ToUInt32(Is2DSelected)) + (1 * Convert.ToUInt32(IsXLSelected))

        Return InputSum

    End Function

    Private Sub _2DOutputLabel_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles _2DOutputLabel.MouseLeftButtonDown
        Select Case InputSum()

            Case 0
                MsgBox("Choose at least one input")
            Case 1
                Call Comparator.XLto2D()
            Case 4
                Call Comparator.Write3Dto2D()
            Case 6
                Call Comparator.Write3DToExcel()
                Call Comparator.Write2DToExcel(ListBox2D.SelectedIndex, ListBox2D.Items.Count)
        End Select
    End Sub

    Private Sub XLLabel_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles XLLabel.MouseLeftButtonDown
        Comparator.SelectXL()
    End Sub

    Private Sub XLLabel_DragEnter(sender As Object, e As DragEventArgs) Handles XLLabel.DragEnter
        Comparator.SelectXL()
    End Sub
End Class
