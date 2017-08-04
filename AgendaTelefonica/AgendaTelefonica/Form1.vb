Public Class Form1

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim escritor As System.IO.StreamWriter
        escritor = New System.IO.StreamWriter("C:/archivos/agenda/agenda.txt", True)
        escritor.WriteLine("'" & TextBox1.Text & "';'" & TextBox2.Text & "';'" & TextBox3.Text & "';'" & ComboBox1.SelectedItem & "'")
        escritor.Close()
    End Sub
End Class
