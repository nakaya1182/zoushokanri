Imports System.Data.SqlClient

Public Class Form2
    Dim evaluation As String = 1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Con As New System.Data.SqlClient.SqlConnection
        '接続先指定
        Con.ConnectionString =
            "Data Source = DESKTOP-Q1GQLT9\SQLEXPRESS;" &
            "Initial Catalog =BookManagement;" &
            "Integrated Security = SSPI"
        'SQLクエリー文が格納される変数commandの作成
        Dim command As System.Data.SqlClient.SqlCommand = Con.CreateCommand()
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
        command.CommandText = "EXEC BookInsert'" & title & "','" & author & "','" & publisher & "','" & language & "','" & page & "','" & evaluation & "','" & place & "','" & remarks & "';"
        command.ExecuteNonQuery()
        MessageBox.Show("処理が完了しました。",
        "完了",
        MessageBoxButtons.OK)
        Con.Close()
        Form1DbOpen()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Public Sub Form1DbOpen()
        Dim Con As New System.Data.SqlClient.SqlConnection
        '接続先指定
        Con.ConnectionString =
            "Data Source = DESKTOP-Q1GQLT9\SQLEXPRESS;" &
            "Initial Catalog = BookManagement;" &
            "Integrated Security = SSPI"
        Con.Open()
        'SQLクエリー文が格納される変数commandの作成
        'テーブル作成
        Dim command As System.Data.SqlClient.SqlCommand = Con.CreateCommand()
        Dim dt As DataTable = New DataTable()
        Dim da As SqlDataAdapter = New SqlDataAdapter("select * from BookInfoView", Con)
        da.Fill(dt)
        Form1.DataGridView1.DataSource = dt
        'コンボボックス作成
        command.CommandText = "SELECT id FROM BookInfoView"
        da.SelectCommand = command
        Form1.ComboBox1.DataSource = dt
        Form1.ComboBox1.DisplayMember = "タイトル"
        Form1.ComboBox1.ValueMember = "id"
        Form1.ComboBox1.SelectedIndex = -1
        Form1.ComboBox1.Text = "編集"
        Form1.ComboBox2.DataSource = dt
        Form1.ComboBox2.DisplayMember = "タイトル"
        Form1.ComboBox2.ValueMember = "id"
        Form1.ComboBox2.SelectedIndex = -1
        Form1.ComboBox2.Text = "削除"
        Con.Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs)

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