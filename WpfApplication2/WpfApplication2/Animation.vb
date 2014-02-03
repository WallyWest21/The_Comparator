Imports System.Windows.Media.Animation



Public Class Animation
    Sub Shazam()
        Dim myDoubleAnimation As New DoubleAnimation()
        myDoubleAnimation.From = 1.0
        myDoubleAnimation.To = 0.0
        myDoubleAnimation.Duration = New Duration(TimeSpan.FromSeconds(5))
        myDoubleAnimation.AutoReverse = True
        myDoubleAnimation.RepeatBehavior = RepeatBehavior.Forever
    End Sub
End Class
