Imports System.Collections.ObjectModel
Imports ProductStructureTypeLib
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading.Tasks
Imports ProductStructureTypeLib.CatWorkModeType
Imports INFITF.CatFileSelectionMode
Imports System.ComponentModel


Public Class ChildrenList
    Inherits ObservableCollection(Of String)
    ' Implements INotifyPropertyChanged
    ' Public Property pChildrenList As New ObservableCollection(Of String)
    ' Inherits ObservableCollection(Of Object)
    Public Sub New()

        '  Dim Item

        ' Dim Comparator As New Comparator
        ' For Each Item In Comparator.Children3D
        'For Item = 1 To 25
        MyBase.Add("No item selected")
        'MyBase.Add("kl;njkbjfkuigkhjklk")
        ' Next

        'MyBase.Add("second tolast")
        'MyBase.Add("kl;lsdfgd Super Last")
    End Sub



End Class