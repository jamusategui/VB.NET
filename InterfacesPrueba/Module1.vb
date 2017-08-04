Module Module1
    Sub Main()
        Dim loEmple As Empleado = New Empleado()
        loEmple.Nombre = "Raquel Rodrigo"
        Console.WriteLine("El nombre del empleado tiene {0} caracteres", _
            loEmple.Longitud)
        Console.WriteLine("Valor del empleado: {0}", loEmple.ObtenerValor())

        Dim loCuenta As Cuenta = New Cuenta()
        loCuenta.Codigo = 700256
        Console.WriteLine("El código de cuenta {0} tiene una longitud de {1}", _
            loCuenta.Codigo, loCuenta.Longitud)
        Console.WriteLine("Información de la cuenta: {0}", loCuenta.ObtenerValor())

        Console.Read()
    End Sub
End Module

' las clases que implementen este interfaz
' deberán crear la propiedad Longitud y el
' método ObtenerValor(); la codificación de
' dichos miembros será particular a cada clase
Public Interface ICadena
    ReadOnly Property Longitud() As Integer
    Function ObtenerValor() As String
End Interface

Public Class Empleado
    Implements ICadena

    Private msNombre As String

    Public Property Nombre() As String
        Get
            Return msNombre
        End Get
        Set(ByVal Value As String)
            msNombre = Value
        End Set
    End Property

    ' devolvemos la longitud de la propiedad Nombre,
    ' al especificar la interfaz que se implementa,
    ' se puede poner todo el espacio de nombres
    Public ReadOnly Property Longitud() As Integer _
        Implements InterfacesPrueba.ICadena.Longitud
        Get
            Return Len(Me.Nombre)
        End Get
    End Property

    ' devolvemos el nombre en mayúscula;
    ' no es necesario poner todo el espacio
    ' de nombres calificados, basta con el
    ' nombre de interfaz y el miembro a implementar
    Public Function ObtenerValor() As String Implements ICadena.ObtenerValor
        Return UCase(Me.Nombre)
    End Function
End Class

Public Class Cuenta
    Implements ICadena

    Private miCodigo As Integer

    Public Property Codigo() As Integer
        Get
            Return miCodigo
        End Get
        Set(ByVal Value As Integer)
            miCodigo = Value
        End Set
    End Property

    ' en esta implementación del miembro, devolvemos
    ' el número de caracteres que tiene el campo 
    ' miCodigo de la clase
    Public ReadOnly Property Longitud() As Integer _
        Implements InterfacesPrueba.ICadena.Longitud
        Get
            Return Len(CStr(miCodigo))
        End Get
    End Property

    ' en este método de la interfaz, devolvemos valores
    ' diferentes en función del contenido de la 
    ' propiedad Codigo
    Public Function ObtenerValor() As String _
        Implements InterfacesPrueba.ICadena.ObtenerValor

        Select Case Me.Codigo
            Case 0 To 1000
                Return "Código en intervalo hasta 1000"
            Case 1001 To 2000
                Return "El código es: " & Me.Codigo
            Case Else
                Return "El código no está en el intervalo"
        End Select
    End Function
End Class