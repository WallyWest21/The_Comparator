Public Class oDrawing

    Public DrawingCode As String

    Public Class Item

        ' Dim Drawig = New Drawing
        Public Column As Integer

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
    Public Class Notes
        Public Class GeneralNotes
        End Class
        Public Class FlagNotes
        End Class
    End Class

End Class


