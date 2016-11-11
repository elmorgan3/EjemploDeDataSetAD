Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class Form1
    Private con As SqlConnection
    Private ds As DataSet
    Private ada As SqlDataAdapter
    Private iClients, iVClients As Integer

    Private dvClients As DataView


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con = New SqlConnection
        con.ConnectionString = "Data Source=.\SQLEXPRESS;Initial Catalog=MAGATZEM; Trusted_Connection=True;"
        con.Open()

        ' Crear el dataset, base de dades virtual en memòria
        ds = New DataSet

        ' Crear el adapter, un per cada taula del dataset
        ada = New SqlDataAdapter("select * from clientes", con)

        ' ACTE DE FE, li donem la inteligencia al adapter perque sàpiga fer Insert, Delete i Update
        Dim cmBase As SqlCommandBuilder = New SqlCommandBuilder(ada)

        ' Construïm la taula virtual en memòria
        Try
            ada.Fill(ds, "CLIENTS")
        Catch
            MsgBox("No s'ha pogut recuperar la informació")
        End Try

        'Definir la clau primaria
        Dim pk(1) As DataColumn
        pk(0) = ds.Tables("CLIENTS").Columns("IDCLIENTE")
        ds.Tables("CLIENTS").PrimaryKey = pk

        'ds.Tables("CLIENTS").Columns.Add("Disponible", Type.GetType("System.Double"), "CREDITOMAXIMO - DEUDA")
        'ds.Tables("CLIENTS").Columns.Add("NombreTotal", Type.GetType("System.String"), "NOMBRECOMPANIA + ', ' + NOMBRECONTACTO")
        ' ds.Tables("CLIENTS").Columns.Add("TieneCredito", Type.GetType("System.String"), "IIf(Disponible>0, 'TIENE CREDITO', 'NO TIENE CREDITO')")

        dvClients = ds.Tables("CLIENTS").DefaultView

        iClients = 0
        iVClients = 0

        MostrarDades()
        MostrarDadesVista()
    End Sub

    Private Sub MostrarDades()
        edtId.Text = ds.Tables("CLIENTS").Rows(iClients)("IDCLIENTE").ToString()
        edtCliente.Text = ds.Tables("CLIENTS").Rows(iClients)("NOMBRECOMPANIA").ToString()

        If IsDBNull(ds.Tables("CLIENTS").Rows(iClients)("NOMBRECONTACTO")) Then
            edtContacto.BackColor = Color.Coral
        Else
            edtContacto.BackColor = Color.White
        End If

        edtContacto.Text = ds.Tables("CLIENTS").Rows(iClients)("NOMBRECONTACTO").ToString()
        'edtCredito.Text = ds.Tables("CLIENTS").Rows(iClients)("CREDITOMAXIMO").ToString()
        'edtDeuda.Text = ds.Tables("CLIENTS").Rows(iClients)("DEUDA").ToString()

        'lblDisponible.Text = ds.Tables("CLIENTS").Rows(iClients)("DISPONIBLE").ToString()
        'lblNombreTotal.Text = ds.Tables("CLIENTS").Rows(iClients)("NOMBRETOTAL").ToString()
        'lblCredito.Text = ds.Tables("CLIENTS").Rows(iClients)("TIENECREDITO").ToString()



        Select Case ds.Tables("CLIENTS").Rows(iClients).RowState
            Case DataRowState.Added
                lblEstat.Text = "NOU"
            Case DataRowState.Modified
                lblEstat.Text = "MODIFICAT"
            Case DataRowState.Unchanged
                lblEstat.Text = "NO MODIFICAT"
        End Select

    End Sub

    Private Sub MostrarDadesVista()
        edtIdVista.Text = dvClients(iVClients)("IDCLIENTE").ToString()
        edtClienteVista.Text = dvClients(iVClients)("NOMBRECOMPANIA").ToString()
    End Sub


    Private Sub butPrimer_Click(sender As Object, e As EventArgs) Handles butPrimer.Click
        iClients = 0
        MostrarDades()
    End Sub

    Private Sub butAnterior_Click(sender As Object, e As EventArgs) Handles butAnterior.Click
        If iClients > 0 Then
            iClients = iClients - 1
            MostrarDades()
        End If
    End Sub

    Private Sub butSeguent_Click(sender As Object, e As EventArgs) Handles butSeguent.Click
        If iClients <> ds.Tables("CLIENTS").Rows.Count - 1 Then
            iClients = iClients + 1
            MostrarDades()
        End If
    End Sub

    Private Sub butUltim_Click(sender As Object, e As EventArgs) Handles butUltim.Click
        iClients = ds.Tables("CLIENTS").Rows.Count - 1
        MostrarDades()
    End Sub

    Private Sub GuardarDades()
        Try
            ds.Tables("CLIENTS").Rows(iClients)("IDCLIENTE") = edtId.Text
        Catch ex As Exception
            MsgBox("El id de cliente debe estar informado y no puede ser vacío")
            edtId.Text = ds.Tables("CLIENTS").Rows(iClients)("IDCLIENTE").ToString

            Exit Sub
        End Try
        ds.Tables("CLIENTS").Rows(iClients)("NOMBRECOMPANIA") = edtCliente.Text

        If edtContacto.Text = "" Then
            ds.Tables("CLIENTS").Rows(iClients)("NOMBRECONTACTO") = DBNull.Value
        Else
            ds.Tables("CLIENTS").Rows(iClients)("NOMBRECONTACTO") = edtContacto.Text
        End If

        ds.Tables("CLIENTS").Rows(iClients)("CREDITOMAXIMO") = edtCredito.Text
        ds.Tables("CLIENTS").Rows(iClients)("DEUDA") = edtDeuda.Text

    End Sub

    Private Sub GuardarDadesVista()
        Try
            dvClients(iVClients)("IDCLIENTE") = edtIdVista.Text
        Catch ex As Exception
            MsgBox("El id de cliente debe estar informado y no puede ser vacío")
            edtId.Text = dvClients(iVClients)("IDCLIENTE")

            Exit Sub
        End Try
        dvClients(iVClients)("NOMBRECOMPANIA") = edtClienteVista.Text

    End Sub

    Private Sub btnGuardarCanvis_Click(sender As Object, e As EventArgs) Handles btnGuardarCanvis.Click
        GuardarDades()
        MostrarDades()
    End Sub

    Private Sub btnNouClient_Click(sender As Object, e As EventArgs) Handles btnNouClient.Click
        Dim dr As DataRow

        dr = ds.Tables("CLIENTS").NewRow()
        dr("CreditoMaximo") = 0
        dr("Deuda") = 0
        dr("idcliente") = ""

        ds.Tables("CLIENTS").Rows.Add(dr)

        iClients = ds.Tables("CLIENTS").Rows.Count - 1
        MostrarDades()
    End Sub

    Private Sub btnBorrar_Click(sender As Object, e As EventArgs) Handles btnBorrar.Click
        If MsgBox("Segur que desitja eliminar?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
        End If

        ds.Tables("CLIENTS").Rows(iClients).Delete()
        iClients = 0
        MostrarDades()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ds.Tables("CLIENTS").GetChanges Is Nothing Then
            MsgBox("no hi ha res pendent de actualitzar")
            Exit Sub
        End If

        ' Metode 1. Actualizem tot
        ada.Update(ds, "CLIENTS")

        ' Metode 2. Actualitzem només el que hi ha per actualtizar
        'Dim dt As DataTable

        'dt = ds.Tables("CLIENTS").GetChanges
        'ada.Update(dt)

        ' Metode 3. Actualitzar el que jo vulgui
        'Dim dt As DataTable

        'dt = ds.Tables("CLIENTS").GetChanges(DataRowState.Deleted)
        'ada.Update(dt)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        iVClients = 0
        MostrarDadesVista()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If iVClients > 0 Then
            iVClients = iVClients - 1
            MostrarDadesVista()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If iVClients < dvClients.Count - 1 Then
            iVClients = iVClients + 1
            MostrarDadesVista()

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        iVClients = dvClients.Count - 1
        MostrarDadesVista()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        GuardarDadesVista()
        MostrarDadesVista()
        MostrarDades()  ' aixo ho fem per si en els dos llocs estem mostrant el mateix registre
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        dvClients.Sort = "IDCLIENTE"
        iVClients = 0
        MostrarDadesVista()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        dvClients.Sort = "NOMBRECOMPANIA"
        iVClients = 0
        MostrarDadesVista()

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If edtFiltre.Text <> "" Then
            dvClients.RowFilter = "nombrecompania like '%" + edtFiltre.Text + "%'"
        Else
            dvClients.RowFilter = ""
        End If

        If dvClients.Count = 0 Then
            MsgBox("no hi ha registre esto va a petar")
        Else
            iVClients = 0
            MostrarDadesVista()
        End If
    End Sub
End Class
