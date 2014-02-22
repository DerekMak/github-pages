Imports System
Imports System.Data
Imports System.Data.OleDb

Public Class Main
    Dim dataReader As OleDbDataReader
    Dim connection As OleDbConnection
    Dim connectionString As String
    Dim command As OleDbCommand
    Dim ds As New DataSet

    Private Shared Function getConnectionString() As String
        Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
            & "C:\temp\Datamart.mdb"
    End Function

    Private Sub refreshSQL(ByVal item As Integer)
        connectionString = getConnectionString()
        connection = New OleDbConnection(connectionString)

        Try
            connection.Open()

            Select Case item
                Case 0
                    Dim queryString As String = "select distinct(country) from location"
                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "lcountry")
                    ctrycbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("lcountry").Rows.Count - 1
                        ctrycbox.Items.Add(ds.Tables("lcountry").Rows.Item(i).Item(0))
                    Next

                Case 1
                    Dim queryString As String = "select distinct(state_or_province) from location " _
                                                + "where (country = '" + ctrycbox.Text + "')"
                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "lsp")
                    provcbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("lsp").Rows.Count - 1
                        provcbox.Items.Add(ds.Tables("lsp").Rows.Item(i).Item(0))
                    Next

                Case 2
                    Dim queryString As String = "select distinct(city) from location " _
                                                + "where (country = '" + ctrycbox.Text + "') " _
                                                + "and (state_or_province = '" + provcbox.Text + "')"

                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "lcity")
                    citycbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("lcity").Rows.Count - 1
                        citycbox.Items.Add(ds.Tables("lcity").Rows.Item(i).Item(0))
                    Next

                Case 3
                    Dim queryString As String = "select distinct(street) from location " _
                                                + "where (country = '" + ctrycbox.Text + "') " _
                                                + "and (state_or_province = '" + provcbox.Text + "') " _
                                                + "and (city = '" + citycbox.Text + "')"
                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "lst")
                    stcbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("lst").Rows.Count - 1
                        stcbox.Items.Add(ds.Tables("lst").Rows.Item(i).Item(0))
                    Next

                Case 4
                    Dim queryString As String = "select distinct(supplier_type) from item"
                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "isupp")
                    suppcbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("isupp").Rows.Count - 1
                        suppcbox.Items.Add(ds.Tables("isupp").Rows.Item(i).Item(0))
                    Next

                Case 5
                    Dim queryString As String = "select distinct(brand) from item " _
                                                + "where (supplier_type = '" + suppcbox.Text + "')"
                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "ibrand")
                    brancbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("ibrand").Rows.Count - 1
                        brancbox.Items.Add(ds.Tables("ibrand").Rows.Item(i).Item(0))
                    Next

                Case 6
                    Dim queryString As String = "select distinct(type) from item " _
                                                + "where (supplier_type = '" + suppcbox.Text + "') " _
                                                + "and (brand = '" + brancbox.Text + "')"
                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "itype")
                    typecbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("itype").Rows.Count - 1
                        typecbox.Items.Add(ds.Tables("itype").Rows.Item(i).Item(0))
                    Next

                Case 7
                    Dim queryString As String = "select distinct(t_year) from time_date"
                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "tyear")
                    yrcbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("tyear").Rows.Count - 1
                        yrcbox.Items.Add(ds.Tables("tyear").Rows.Item(i).Item(0))
                    Next

                Case 8
                    Dim queryString As String = "select distinct(quarter) from time_date " _
                                                + "where (t_year = " + yrcbox.Text + ")"
                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "tquarter")
                    qtrcbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("tquarter").Rows.Count - 1
                        qtrcbox.Items.Add(ds.Tables("tquarter").Rows.Item(i).Item(0))
                    Next

                Case 9
                    Dim queryString As String = "select distinct(t_month) from time_date " _
                                                + "where (t_year = " + yrcbox.Text + ") " _
                                                + "and (quarter = " + qtrcbox.Text + ")"
                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "tmonth")
                    moncbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("tmonth").Rows.Count - 1
                        moncbox.Items.Add(ds.Tables("tmonth").Rows.Item(i).Item(0))
                    Next

                Case 10
                    Dim queryString As String = "select distinct(day) from time_date " _
                                                + "where (t_year = " + yrcbox.Text + ") " _
                                                + "and (quarter = " + qtrcbox.Text + ") " _
                                                + "and (t_month = '" + moncbox.Text + "')"
                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "tday")
                    daycbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("tday").Rows.Count - 1
                        daycbox.Items.Add(ds.Tables("tday").Rows.Item(i).Item(0))
                    Next

                Case 11
                    Dim queryString As String = "select distinct(branch_type) from branch"
                    Dim da As OleDbDataAdapter = New OleDb.OleDbDataAdapter(queryString, connection)

                    ds.Clear()
                    da.Fill(ds, "btype")
                    btcbox.Items.Clear()

                    command = connection.CreateCommand
                    command.CommandText = queryString

                    For i As Integer = 0 To ds.Tables("btype").Rows.Count - 1
                        btcbox.Items.Add(ds.Tables("btype").Rows.Item(i).Item(0))
                    Next
            End Select

        Catch ex As Exception
            connection.Close()
        End Try
        connection.Close()
    End Sub

    Private Function addToString(ByVal ParamArray nullval() As Integer) As String
        Dim str As String = ""
        Dim startWhen As Boolean = False

        For i As Integer = 0 To nullval.Length - 1
            If nullval(i) <> 0 And startWhen = False Then
                str = " where "
                startWhen = True
            End If

            If nullval(i) <> 0 Then
                Select Case i
                    Case 0
                        str = str + "(fact_table.location_key = location.location_key)" _
                            + " and (location.country = '" + ctrycbox.Text + "')"

                        If provcbox.Text.Length <> 0 Then
                            str = str + " and (location.state_or_province = '" + provcbox.Text + "')"
                        End If

                        If citycbox.Text.Length <> 0 Then
                            str = str + " and (location.city = '" + citycbox.Text + "')"
                        End If

                        If stcbox.Text.Length <> 0 Then
                            str = str + " and (location.street = '" + stcbox.Text + "')"
                        End If

                    Case 1
                        If nullval(i - 1) <> 0 Then
                            str = str + " and"
                        End If

                        str = str + " (fact_table.item_key = item.item_key)" _
                            + " and (item.supplier_type = '" + suppcbox.Text + "')"

                        If brancbox.Text.Length <> 0 Then
                            str = str + " and (item.brand = '" + brancbox.Text + "')"
                        End If

                        If typecbox.Text.Length <> 0 Then
                            str = str + " and (item.type = '" + typecbox.Text + "')"
                        End If

                    Case 2
                        If nullval(i - 1) <> 0 Or nullval(i - 2) <> 0 Then
                            str = str + " and"
                        End If

                        str = str + " (fact_table.time_key = time_date.time_key)" _
                            + " and (time_date.t_year = " + yrcbox.Text + ")"

                        If qtrcbox.Text.Length <> 0 Then
                            str = str + " and (time_date.quarter = " + qtrcbox.Text + ")"
                        End If

                        If moncbox.Text.Length <> 0 Then
                            str = str + " and (time_date.t_month = '" + moncbox.Text + "')"
                        End If

                        If daycbox.Text.Length <> 0 Then
                            str = str + " and (time_date.day = " + daycbox.Text + ")"
                        End If

                    Case 3
                        If nullval(i - 1) <> 0 Or nullval(i - 2) <> 0 Or nullval(i - 3) <> 0 Then
                            str = str + " and"
                        End If

                        str = str + " (fact_table.branch_key = branch.branch_key)" _
                            + " and (branch.branch_type = '" + btcbox.Text + "')"
                End Select
            End If
        Next

        Return str
    End Function

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        connectionString = getConnectionString()
        connection = New OleDbConnection(connectionString)

        Try
            connection.Open()

            refreshSQL(0)
            refreshSQL(4)
            refreshSQL(7)
            refreshSQL(11)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        connection.Close()

    End Sub

    Private Sub Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Search.Click
        Dim queryString As String = "select distinct(sum(units_sold)), sum(dollars_sold), sum(avg_sales) from fact_table"

        Dim nullval = New Integer() {0, 0, 0, 0}

        If ctrycbox.Text.Length <> 0 Then
            nullval(0) = 1
            queryString = queryString + ", location"
        End If

        If suppcbox.Text.Length <> 0 Then
            nullval(1) = 1
            queryString = queryString + ", item"
        End If

        If yrcbox.Text.Length <> 0 Then
            nullval(2) = 1
            queryString = queryString + ", time_date"
        End If

        If btcbox.Text.Length <> 0 Then
            nullval(3) = 1
            queryString = queryString + ", branch"
        End If

        queryString = queryString + addToString(nullval)

        Try
            connection.Open()

            Dim da As OleDbDataAdapter = New OleDbDataAdapter(queryString, connection)

            ds.Clear()
            da.Fill(ds, "sales")
            resulttxtbox.Clear()

            command = connection.CreateCommand
            command.CommandText = queryString

            resulttxtbox.AppendText("Units sold" + vbTab + vbTab + "Dollars sold" + vbTab + "Average sales" + vbNewLine)

            For i As Integer = 0 To ds.Tables("sales").Rows.Count - 1
                resulttxtbox.AppendText(ds.Tables("sales").Rows(i).Item(0).ToString + vbTab + vbTab)
                resulttxtbox.AppendText(ds.Tables("sales").Rows(i).Item(1).ToString + vbTab + vbTab)
                resulttxtbox.AppendText(ds.Tables("sales").Rows(i).Item(2).ToString + vbNewLine)
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
            connection.Close()
    End Sub

    Private Sub ctrycbox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctrycbox.SelectedIndexChanged
        If ctrycbox.Text.Length <> 0 And provcbox.Visible = False Then
            ctryd.Visible = True
        End If

        refreshSQL(1)
        provcbox.Text = ""
        citycbox.Text = ""
        stcbox.Text = ""
    End Sub

    Private Sub ctryd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctryd.Click
        provcbox.Visible = True
        provu.Visible = True
        ctryd.Visible = False
    End Sub

    Private Sub provcbox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles provcbox.SelectedIndexChanged
        If provcbox.Text.Length <> 0 And citycbox.Visible = False Then
            provd.Visible = True
        End If

        refreshSQL(2)
        citycbox.Text = ""
        stcbox.Text = ""
    End Sub

    Private Sub provu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles provu.Click
        provcbox.Text = ""
        provcbox.Visible = False
        provd.Visible = False
        provu.Visible = False
        ctryd.Visible = True
    End Sub

    Private Sub provd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles provd.Click
        citycbox.Visible = True
        cityu.Visible = True
        provu.Visible = False
        provd.Visible = False
    End Sub

    Private Sub citycbox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles citycbox.SelectedIndexChanged
        If citycbox.Text.Length <> 0 And stcbox.Visible = False Then
            cityd.Visible = True
        End If

        refreshSQL(3)
        stcbox.Text = ""
    End Sub

    Private Sub cityu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cityu.Click
        citycbox.Text = ""
        citycbox.Visible = False
        cityd.Visible = False
        cityu.Visible = False
        provu.Visible = True
        provd.Visible = True
    End Sub

    Private Sub cityd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cityd.Click
        stcbox.Visible = True
        stu.Visible = True
        cityu.Visible = False
        cityd.Visible = False
    End Sub

    Private Sub stu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stu.Click
        stcbox.Text = ""
        stcbox.Visible = False
        stcbox.Visible = False
        stu.Visible = False
        cityu.Visible = True
        cityd.Visible = True
    End Sub

    Private Sub suppcbox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles suppcbox.SelectedIndexChanged
        If suppcbox.Text.Length <> 0 And brancbox.Visible = False Then
            suppd.Visible = True
        End If

        refreshSQL(5)
        brancbox.Text = ""
        typecbox.Text = ""
    End Sub

    Private Sub suppd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles suppd.Click
        brancbox.Visible = True
        branu.Visible = True
        suppd.Visible = False
    End Sub

    Private Sub brancbox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles brancbox.SelectedIndexChanged
        If brancbox.Text.Length <> 0 And typecbox.Visible = False Then
            brand.Visible = True
        End If

        refreshSQL(6)
        typecbox.Text = ""
    End Sub

    Private Sub branu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles branu.Click
        brancbox.Text = ""
        brancbox.Visible = False
        branu.Visible = False
        brand.Visible = False
        suppd.Visible = True
    End Sub

    Private Sub brand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles brand.Click
        typecbox.Visible = True
        typeu.Visible = True
        branu.Visible = False
        brand.Visible = False
    End Sub

    Private Sub typeu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles typeu.Click
        typecbox.Text = ""
        typecbox.Visible = False
        typeu.Visible = False
        branu.Visible = True
        brand.Visible = True
    End Sub

    Private Sub yrcbox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yrcbox.SelectedIndexChanged
        If yrcbox.Text.Length <> 0 And qtrcbox.Visible = False Then
            yrd.Visible = True
        End If

        refreshSQL(8)
        qtrcbox.Text = ""
        moncbox.Text = ""
        daycbox.Text = ""
    End Sub

    Private Sub yrd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yrd.Click
        qtrcbox.Visible = True
        qtru.Visible = True
        yrd.Visible = False
    End Sub

    Private Sub qtrcbox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles qtrcbox.SelectedIndexChanged
        If qtrcbox.Text.Length <> 0 And moncbox.Visible = False Then
            qtrd.Visible = True
        End If

        refreshSQL(9)
        moncbox.Text = ""
        daycbox.Text = ""
    End Sub

    Private Sub qtru_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles qtru.Click
        qtrcbox.Text = ""
        qtrcbox.Visible = False
        qtru.Visible = False
        qtrd.Visible = False
        yrd.Visible = True
    End Sub

    Private Sub qtrd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles qtrd.Click
        moncbox.Visible = True
        monu.Visible = True
        qtru.Visible = False
        qtrd.Visible = False
    End Sub

    Private Sub moncbox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles moncbox.SelectedIndexChanged
        If moncbox.Text.Length <> 0 And daycbox.Visible = False Then
            mond.Visible = True
        End If

        refreshSQL(10)
        daycbox.Text = ""
    End Sub

    Private Sub monu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles monu.Click
        moncbox.Text = ""
        moncbox.Visible = False
        monu.Visible = False
        mond.Visible = False
        qtru.Visible = True
        qtrd.Visible = True
    End Sub

    Private Sub mond_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mond.Click
        daycbox.Visible = True
        dayu.Visible = True
        monu.Visible = False
        mond.Visible = False
    End Sub

    Private Sub dayu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dayu.Click
        daycbox.Text = ""
        daycbox.Visible = False
        dayu.Visible = False
        monu.Visible = True
        mond.Visible = True
    End Sub

    'Rests to the original settings
    Private Sub Reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Reset.Click
        ctrycbox.Text = ""
        provcbox.Text = ""
        citycbox.Text = ""
        stcbox.Text = ""

        suppcbox.Text = ""
        brancbox.Text = ""
        typecbox.Text = ""

        yrcbox.Text = ""
        qtrcbox.Text = ""
        moncbox.Text = ""
        daycbox.Text = ""

        btcbox.Text = ""

        ctryd.Visible = False
        provu.Visible = False
        provd.Visible = False
        cityu.Visible = False
        cityd.Visible = False
        stu.Visible = False

        suppd.Visible = False
        branu.Visible = False
        brand.Visible = False
        typeu.Visible = False

        yrd.Visible = False
        qtru.Visible = False
        qtrd.Visible = False
        monu.Visible = False
        mond.Visible = False
        dayu.Visible = False

        provcbox.Visible = False
        citycbox.Visible = False
        stcbox.Visible = False

        brancbox.Visible = False
        typecbox.Visible = False

        qtrcbox.Visible = False
        moncbox.Visible = False
        daycbox.Visible = False

        resulttxtbox.Text = ""
    End Sub
End Class
