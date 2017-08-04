
' importar espacios de nombres de la librería
Imports Rutinas
Imports Rutinas.Personal

Module Module1

    Sub Main()
        ' instanciar y trabajar con objetos de la librería
        Dim oContab As Contabilidad
        Dim oEmple As Empleado
        oContab = New Contabilidad()
        oContab.NombreBalance = "Saldos"
        oContab.MostrarBalance()

        oEmple = New Empleado()
        Console.WriteLine("La fecha es: {0}", oEmple.FechaActual())

        Console.Read()
    End Sub

End Module
