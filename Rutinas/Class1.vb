Public Class Contabilidad
    Private mNombreBalance As String

    Public Property NombreBalance() As String
        Get
            Return mNombreBalance
        End Get
        Set(ByVal Value As String)
            mNombreBalance = Value
        End Set
    End Property

    Public Sub MostrarBalance()
        Console.WriteLine("Balance actual: {0}", mNombreBalance)
    End Sub
End Class

Namespace Personal
    Public Class Empleado
        Public Function FechaActual() As Date
            Dim Fecha As Date
            Fecha = Now
            Return Fecha
        End Function
    End Class
End Namespace