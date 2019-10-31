Imports System.Data.SqlClient

Public Class Form3
    Dim evaluation As String = 1
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        command.CommandText = "select タイトル from BookInfoView WHERE id =" & id
        txt = command.ExecuteScalar()
        TextBox1.Text = txt
        command.CommandText = "select 著者 from BookInfoView WHERE id =" & id
        txt = command.ExecuteScalar()
        TextBox2.Text = txt
        command.CommandText = "select 出版社 from BookInfoView WHERE id =" & id
        txt = command.ExecuteScalar()
        TextBox3.Text = txt
        command.CommandText = "select 言語 from BookInfoView WHERE id =" & id
        txt = command.ExecuteScalar()
        TextBox4.Text = txt
        command.CommandText = "select ページ from BookInfoView WHERE id =" & id
        txt = command.ExecuteScalar()
        TextBox5.Text = txt
        command.CommandText = "select 評価 from BookInfoView WHERE id =" & id
        txt = command.ExecuteScalar()
        evaluation = txt
        command.CommandText = "select 場所 from BookInfoView WHERE id =" & id
        txt = command.ExecuteScalar()
        TextBox7.Text = txt
        command.CommandText = "select 備考 from BookInfoView WHERE id =" & id
        txt = command.ExecuteScalar()
        TextBox8.Text = txt
        If evaluation = 2 Then
            Button4.ForeColor = System.Drawing.Color.Yellow
            Button5.ForeColor = System.Drawing.Color.White
            Button4.Text = "★"
            Button5.Text = "☆"
        ElseIf evaluation = 3 Then
            Button4.ForeColor = System.Drawing.Color.Yellow
            Button5.ForeColor = System.Drawing.Color.Yellow
            Button4.Text = "★"
            Button5.Text = "★"
        End If

        Con.Close()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Con As New System.Data.SqlClient.SqlConnection
        '接続先指定
        Con.ConnectionString =
            "Data Source = DESKTOP-Q1GQLT9\SQLEXPRESS;" &
            "Initial Catalog =BookManagement;" &
            "Integrated Security = SSPI"
        'SQLクエリー文が格納される変数commandの作成
        Dim command As System.Data.SqlClient.SqlCommand = Con.CreateCommand()
        Dim id As Integer = Form1.ComboBox1.SelectedValue
        Dim title As String = TextBox1.Text
        Dim author As String = TextBox2.Text
        Dim publisher As String = TextBox3.Text
        Dim language As String = TextBox4.Text
        Dim page As String = TextBox5.Text
        Dim place As String = TextBox7.Text
        Dim remarks As String = TextBox8.Text

        If title = "" Then
            MessageBox.Show("名前に値を入力してください。",
            "エラー",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error)
            Return
        End If
        'データベースをオープン
        Con.Open()
        command.CommandText = "EXEC BookUpdate'" & id & "','" & title & "','" & author & "','" & publisher & "','" & language & "','" & page & "','" & evaluation & "','" & place & "','" & remarks & "';"
        command.ExecuteNonQuery()
        MessageBox.Show("処理が完了しました。",
        "完了",
        MessageBoxButtons.OK)
        Con.Close()
        Dim form2 As New Form2
        form2.Form1DbOpen()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button4.ForeColor = System.Drawing.Color.White
        Button5.ForeColor = System.Drawing.Color.White
        Button4.Text = "☆"
        Button5.Text = "☆"
        evaluation = 1
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button4.ForeColor = System.Drawing.Color.Yellow
        Button5.ForeColor = System.Drawing.Color.White
        Button4.Text = "★"
        Button5.Text = "☆"
        evaluation = 2
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Button4.ForeColor = System.Drawing.Color.Yellow
        Button5.ForeColor = System.Drawing.Color.Yellow
        Button4.Text = "★"
        Button5.Text = "★"
        evaluation = 3
    End Sub
End Class