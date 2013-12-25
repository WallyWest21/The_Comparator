Public Class Validation
    ''' <summary>
    ''' Returns a True Value if object is a Product
    ''' </summary>
    Function IsProduct(oProduct) As Boolean
        IsProduct = False

        Dim TypeOfObject As String
        TypeOfObject = TypeName(oProduct)

        If oProduct.Products.Count > 0 Then 'Assuming that all procucts contain at least 1 child(Part or another product)
            If IsComponent(oProduct) = False Then

                If TypeOfObject = "Product" Then
                    IsProduct = True
                End If

            End If
        End If
    End Function


    ''' <summary>
    ''' Returns a True Value if object is a Component
    ''' </summary>
    Function IsComponent(oComponent) As Boolean
        IsComponent = False
        Dim Component() As String

        Component = Split(oComponent.ReferenceProduct.Parent.Name, ".")
        'If IsProduct(oComponent) = True Then
        If oComponent.ReferenceProduct.Name <> Component(0) Then
            IsComponent = True
            'End If
        End If

    End Function

    ''' <summary>
    ''' Returns a True Value if object is a Part
    ''' </summary>
    Public Function IsPart() As Boolean
        Return True
    End Function

    Function IsValid2DTable() As Boolean
        Return True
    End Function
End Class
