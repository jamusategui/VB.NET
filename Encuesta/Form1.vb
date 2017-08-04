
Imports System.IO

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub
    Friend WithEvents btnSalir As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNombre As System.Windows.Forms.TextBox
    Friend WithEvents txtObservaciones As System.Windows.Forms.TextBox
    Friend WithEvents lblNumLetras As System.Windows.Forms.Label
    Friend WithEvents chkCompra As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbtPropio As System.Windows.Forms.RadioButton
    Friend WithEvents rbtPublico As System.Windows.Forms.RadioButton
    Friend WithEvents lstSecciones As System.Windows.Forms.ListBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboEstiloMusical As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents mnuPrincipal As System.Windows.Forms.MainMenu
    Friend WithEvents mnuArchivo As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDatos As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGrabar As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSalir As System.Windows.Forms.MenuItem
    Friend WithEvents rbtNuevo As System.Windows.Forms.RadioButton
    Friend WithEvents rbtHabitual As System.Windows.Forms.RadioButton

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.Container

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboEstiloMusical = New System.Windows.Forms.ComboBox()
        Me.mnuArchivo = New System.Windows.Forms.MenuItem()
        Me.mnuDatos = New System.Windows.Forms.MenuItem()
        Me.mnuGrabar = New System.Windows.Forms.MenuItem()
        Me.mnuSalir = New System.Windows.Forms.MenuItem()
        Me.txtObservaciones = New System.Windows.Forms.TextBox()
        Me.lstSecciones = New System.Windows.Forms.ListBox()
        Me.chkCompra = New System.Windows.Forms.CheckBox()
        Me.rbtNuevo = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbtHabitual = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rbtPropio = New System.Windows.Forms.RadioButton()
        Me.rbtPublico = New System.Windows.Forms.RadioButton()
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.mnuPrincipal = New System.Windows.Forms.MainMenu()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.lblNumLetras = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(36, 240)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Secciones visitadas"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(83, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Observaciones:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Nombre completo:"
        '
        'cboEstiloMusical
        '
        Me.cboEstiloMusical.DropDownWidth = 124
        Me.cboEstiloMusical.Location = New System.Drawing.Point(196, 264)
        Me.cboEstiloMusical.Name = "cboEstiloMusical"
        Me.cboEstiloMusical.Size = New System.Drawing.Size(124, 21)
        Me.cboEstiloMusical.TabIndex = 10
        '
        'mnuArchivo
        '
        Me.mnuArchivo.Index = 0
        Me.mnuArchivo.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuDatos, Me.mnuGrabar, Me.mnuSalir})
        Me.mnuArchivo.Text = "Archivo"
        '
        'mnuDatos
        '
        Me.mnuDatos.Index = 0
        Me.mnuDatos.Text = "Datos introducidos"
        '
        'mnuGrabar
        '
        Me.mnuGrabar.Index = 1
        Me.mnuGrabar.Text = "&Grabar"
        '
        'mnuSalir
        '
        Me.mnuSalir.Index = 2
        Me.mnuSalir.Shortcut = System.Windows.Forms.Shortcut.CtrlS
        Me.mnuSalir.Text = "&Salir"
        '
        'txtObservaciones
        '
        Me.txtObservaciones.Location = New System.Drawing.Point(116, 42)
        Me.txtObservaciones.Multiline = True
        Me.txtObservaciones.Name = "txtObservaciones"
        Me.txtObservaciones.Size = New System.Drawing.Size(266, 65)
        Me.txtObservaciones.TabIndex = 2
        Me.txtObservaciones.Text = ""
        '
        'lstSecciones
        '
        Me.lstSecciones.IntegralHeight = False
        Me.lstSecciones.Items.AddRange(New Object() {"Libros", "Música", "Informática", "Papelería", "Deportes", "Calzado", "Electrónica"})
        Me.lstSecciones.Location = New System.Drawing.Point(17, 262)
        Me.lstSecciones.Name = "lstSecciones"
        Me.lstSecciones.Size = New System.Drawing.Size(143, 65)
        Me.lstSecciones.TabIndex = 9
        '
        'chkCompra
        '
        Me.chkCompra.Location = New System.Drawing.Point(120, 119)
        Me.chkCompra.Name = "chkCompra"
        Me.chkCompra.Size = New System.Drawing.Size(170, 17)
        Me.chkCompra.TabIndex = 4
        Me.chkCompra.Text = "¿Ha efectuado compra?"
        '
        'rbtNuevo
        '
        Me.rbtNuevo.Location = New System.Drawing.Point(16, 50)
        Me.rbtNuevo.Name = "rbtNuevo"
        Me.rbtNuevo.Size = New System.Drawing.Size(65, 17)
        Me.rbtNuevo.TabIndex = 6
        Me.rbtNuevo.Text = "Nuevo"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.rbtNuevo, Me.rbtHabitual})
        Me.GroupBox1.Location = New System.Drawing.Point(44, 146)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(116, 79)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Tipo de cliente"
        '
        'rbtHabitual
        '
        Me.rbtHabitual.Location = New System.Drawing.Point(16, 20)
        Me.rbtHabitual.Name = "rbtHabitual"
        Me.rbtHabitual.Size = New System.Drawing.Size(65, 17)
        Me.rbtHabitual.TabIndex = 6
        Me.rbtHabitual.Text = "Habitual"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.rbtPropio, Me.rbtPublico})
        Me.GroupBox2.Location = New System.Drawing.Point(192, 145)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(156, 81)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Medio de desplazamiento"
        '
        'rbtPropio
        '
        Me.rbtPropio.Location = New System.Drawing.Point(12, 20)
        Me.rbtPropio.Name = "rbtPropio"
        Me.rbtPropio.Size = New System.Drawing.Size(104, 17)
        Me.rbtPropio.TabIndex = 5
        Me.rbtPropio.Text = "Vehículo propio"
        '
        'rbtPublico
        '
        Me.rbtPublico.Location = New System.Drawing.Point(12, 48)
        Me.rbtPublico.Name = "rbtPublico"
        Me.rbtPublico.Size = New System.Drawing.Size(122, 17)
        Me.rbtPublico.TabIndex = 5
        Me.rbtPublico.Text = "Transporte público"
        '
        'txtNombre
        '
        Me.txtNombre.Location = New System.Drawing.Point(116, 12)
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.Size = New System.Drawing.Size(127, 20)
        Me.txtNombre.TabIndex = 2
        Me.txtNombre.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(200, 242)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(121, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Estilo musical preferido"
        '
        'mnuPrincipal
        '
        Me.mnuPrincipal.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuArchivo})
        '
        'btnSalir
        '
        Me.btnSalir.Location = New System.Drawing.Point(272, 312)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(106, 28)
        Me.btnSalir.TabIndex = 0
        Me.btnSalir.Text = "Salir"
        '
        'lblNumLetras
        '
        Me.lblNumLetras.Location = New System.Drawing.Point(257, 14)
        Me.lblNumLetras.Name = "lblNumLetras"
        Me.lblNumLetras.Size = New System.Drawing.Size(79, 17)
        Me.lblNumLetras.TabIndex = 3
        Me.lblNumLetras.Text = "NumLetras"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(390, 346)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label4, Me.cboEstiloMusical, Me.Label3, Me.lstSecciones, Me.GroupBox2, Me.GroupBox1, Me.chkCompra, Me.lblNumLetras, Me.txtObservaciones, Me.txtNombre, Me.Label2, Me.Label1, Me.btnSalir})
        Me.Menu = Me.mnuPrincipal
        Me.Name = "Form1"
        Me.Text = "Encuesta a clientes"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnSalir_MouseEnter(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles btnSalir.MouseEnter

        Me.btnSalir.Text = "¡¡SORPRESA!!"

    End Sub

    Private Sub btnSalir_MouseLeave(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles btnSalir.MouseLeave

        Me.btnSalir.Text = "Salir"

    End Sub

    Private Sub txtNombre_TextChanged(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles txtNombre.TextChanged

        Me.lblNumLetras.Text = Me.txtNombre.Text.Length

    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.cboEstiloMusical.Items.AddRange(New String() {"Pop", "Rock", "Clásica", "New Age"})
    End Sub

    Private Sub mnuSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSalir.Click
        Me.Close()
    End Sub

    Private Sub mnuDatos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDatos.Click
        Dim Texto As String

        Texto = Me.txtNombre.Text & ControlChars.CrLf
        Texto &= Me.txtObservaciones.Text & ControlChars.CrLf
        Texto &= IIf(Me.chkCompra.Checked, "Ha realizado compra", "No ha comprado") & ControlChars.CrLf

        If Me.rbtHabitual.Checked Then
            Texto &= "Es un cliente habitual" & ControlChars.CrLf
        End If

        If Me.rbtNuevo.Checked Then
            Texto &= "Es un nuevo cliente" & ControlChars.CrLf
        End If

        If Me.rbtPropio.Checked Then
            Texto &= "Ha venido en vehículo propio" & ControlChars.CrLf
        End If

        If Me.rbtPublico.Checked Then
            Texto &= "Ha venido utilizando transporte público" & ControlChars.CrLf
        End If

        Texto &= "Ha visitado la sección: " & Me.lstSecciones.SelectedItem & ControlChars.CrLf
        Texto &= "Su música preferida es: " & Me.cboEstiloMusical.Text

        MessageBox.Show(Texto, "Datos introducidos en el formulario")

    End Sub

    Private Sub mnuGrabar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGrabar.Click
        Dim Archivo As String
        Dim Escritor As StreamWriter

        Archivo = InputBox("Ruta y nombre de archivo")
        Escritor = New StreamWriter(Archivo)

        Escritor.WriteLine(Me.txtNombre.Text)
        Escritor.WriteLine(Me.txtObservaciones.Text)
        Escritor.WriteLine(IIf(Me.chkCompra.Checked, "Ha realizado compra", "No ha comprado"))

        If Me.rbtHabitual.Checked Then
            Escritor.WriteLine("Es un cliente habitual")
        End If

        If Me.rbtNuevo.Checked Then
            Escritor.WriteLine("Es un nuevo cliente")
        End If

        If Me.rbtPropio.Checked Then
            Escritor.WriteLine("Ha venido en vehículo propio")
        End If

        If Me.rbtPublico.Checked Then
            Escritor.WriteLine("Ha venido utilizando transporte público")
        End If

        Escritor.WriteLine("Ha visitado la sección: " & Me.lstSecciones.SelectedItem)
        Escritor.WriteLine("Su música preferida es: " & Me.cboEstiloMusical.Text)

        Escritor.Close()

        MessageBox.Show("Archivo de datos grabado")

    End Sub
End Class
