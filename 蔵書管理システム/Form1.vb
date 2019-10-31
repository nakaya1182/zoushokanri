Imports System.Data.SqlClient

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Con As New System.Data.SqlClient.SqlConnection
        '接続先指定
        Con.ConnectionString =
            "Data Source = DESKTOP-Q1GQLT9\SQLEXPRESS;" &
            "Initial Catalog = BookManagement;" &
            "Integrated Security = SSPI"

        'データベースをオープン
        Con.Open()
        'SQLクエリー文が格納される変数commandの作成
        'テーブル作成
        Dim command As System.Data.SqlClient.SqlCommand = Con.CreateCommand()
        Dim dt As DataTable = New DataTable()
        Dim da As SqlDataAdapter = New SqlDataAdapter("select * from BookInfoView", Con)
        da.Fill(dt)
        DataGridView1.DataSource = dt
        'コンボボックス作成
        command.CommandText = "SELECT id FROM BookInfoView"
        da.SelectCommand = command
        ComboBox1.DataSource = dt
        ComboBox1.DisplayMember = "タイトル"
        ComboBox1.ValueMember = "id"
        ComboBox1.SelectedIndex = -1
        ComboBox1.Text = "選択"
        ComboBox2.DataSource = dt
        ComboBox2.DisplayMember = "タイトル"
        ComboBox2.ValueMember = "id"
        ComboBox2.SelectedIndex = -1
        ComboBox2.Text = "選択"
        Con.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim f As New Form2()
        f.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim f As New Form3()
        f.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim f As New Form4()
        f.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim Con As New System.Data.SqlClient.SqlConnection
        Dim total As Object
        '接続先指定
        Con.ConnectionString =
            "Data Source = DESKTOP-Q1GQLT9\SQLEXPRESS;" &
            "Initial Catalog = BookManagement;" &
            "Integrated Security = SSPI"
        Dim command As System.Data.SqlClient.SqlCommand = Con.CreateCommand()
        'データベースをオープン
        Con.Open()
        command.CommandText = "declare @result as integer
            exec @result = BookPageSum
            select @result as result"
        total = command.ExecuteScalar()
        Con.Close()
        Label1.Text = total.ToString
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim Con As New System.Data.SqlClient.SqlConnection
        Dim sum As Object
        '接続先指定
        Con.ConnectionString =
            "Data Source = DESKTOP-Q1GQLT9\SQLEXPRESS;" &
            "Initial Catalog = BookManagement;" &
            "Integrated Security = SSPI"
        'データベースをオープン
        Con.Open()

        Dim command As System.Data.SqlClient.SqlCommand = Con.CreateCommand()
        Dim dt As DataTable = New DataTable()
        Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT 著者,SUM(ページ) FROM BookInfoView GROUP BY 著者", Con)
        da.Fill(dt)
        DataGridView1.DataSource = dt
        Con.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim Con As New System.Data.SqlClient.SqlConnection
        Dim sum As Object
        '接続先指定
        Con.ConnectionString =
            "Data Source = DESKTOP-Q1GQLT9\SQLEXPRESS;" &
            "Initial Catalog = BookManagement;" &
            "Integrated Security = SSPI"
        'データベースをオープン
        Con.Open()

        Dim command As System.Data.SqlClient.SqlCommand = Con.CreateCommand()
        Dim dt As DataTable = New DataTable()
        Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT 著者,AVG(ページ) FROM BookInfoView GROUP BY 著者", Con)
        da.Fill(dt)
        DataGridView1.DataSource = dt
        Con.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim Con As New System.Data.SqlClient.SqlConnection
        Dim sum As Object
        '接続先指定
        Con.ConnectionString =
            "Data Source = DESKTOP-Q1GQLT9\SQLEXPRESS;" &
            "Initial Catalog = BookManagement;" &
            "Integrated Security = SSPI"
        'データベースをオープン
        Con.Open()

        Dim command As System.Data.SqlClient.SqlCommand = Con.CreateCommand()
        Dim dt As DataTable = New DataTable()
        Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT 著者,MAX(ページ) FROM BookInfoView GROUP BY 著者", Con)
        da.Fill(dt)
        DataGridView1.DataSource = dt
        Con.Close()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim Con As New System.Data.SqlClient.SqlConnection
        Dim sum As Object
        '接続先指定
        Con.ConnectionString =
            "Data Source = DESKTOP-Q1GQLT9\SQLEXPRESS;" &
            "Initial Catalog = BookManagement;" &
            "Integrated Security = SSPI"
        'データベースをオープン
        Con.Open()

        Dim command As System.Data.SqlClient.SqlCommand = Con.CreateCommand()
        Dim dt As DataTable = New DataTable()
        Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT 著者,MIN(ページ) FROM BookInfoView GROUP BY 著者", Con)
        da.Fill(dt)
        DataGridView1.DataSource = dt
        Con.Close()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim Con As New System.Data.SqlClient.SqlConnection
        '接続先指定
        Con.ConnectionString =
            "Data Source = DESKTOP-Q1GQLT9\SQLEXPRESS;" &
            "Initial Catalog = BookManagement;" &
            "Integrated Security = SSPI"

        'データベースをオープン
        Con.Open()
        'SQLクエリー文が格納される変数commandの作成
        'テーブル作成
        Dim command As System.Data.SqlClient.SqlCommand = Con.CreateCommand()
        Dim dt As DataTable = New DataTable()
        Dim da As SqlDataAdapter = New SqlDataAdapter("select * from BookInfoView", Con)
        da.Fill(dt)
        DataGridView1.DataSource = dt
        Con.Close()
    End Sub
End Class
