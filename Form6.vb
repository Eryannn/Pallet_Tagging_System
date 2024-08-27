Imports ZXing

Public Class Form6
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Dim writer As New BarcodeWriter
        writer.Format = BarcodeFormat.EAN_13

        PictureBox1.Image = writer.Write(txtjob.Text)
        'PictureBox2.Image = writer.Write(TextBox1.Text)
        'PictureBox3.Image = writer.Write(TextBox2.Text)
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
        Me.Dispose()
    End Sub
End Class