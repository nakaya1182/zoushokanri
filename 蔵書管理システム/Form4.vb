Public Class Form4
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Con As New System.Data.SqlClient.SqlConnection
        Dim id As String = Form1.ComboBox1.SelectedValue
        Dim txt As String
        '接続先指定
        Con.ConnectionString =
            "Data Source = DESKTOP-Q1GQLT9\SQLEXPRESS;" &
            "Initial Catalog = BookManagement;" &
            "Integrated Security = SSPI"
        Dim command As System.Data.SqlClient.SqlCommand = Con.CreateCommand()
        'データベースをオープン
        Con.Open()
        command.CommandText = "EXEC BookDelete'" & id & "';"
        command.ExecuteNonQuery()
        MessageBox.Show("処理が完了しました。",
        "完了",
        MessageBoxButtons.OK)
        Con.Close()
        Dim form2 As New Form2
        form2.Form1DbOpen()
        Me.Close()
    End Sub
End Class