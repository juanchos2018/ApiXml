Imports CapaDatos
Public Class Ntbl_Lic_Obras_Procesos
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idobras As Integer
    Public Property nombreobra As String
    Public Property nroentregas_obra As Integer
    Public Property lugarentrega As String
    Public Property fechaentregaobra As System.DateTime
    Public Property idproceso As Long
#End Region
#Region "Constructors"
    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As Ntbl_Lic_Obras_Procesos)

        Dim parametros() As Object = {"@nombreobra", "@nroentregas_obra", "@lugarentrega", "@fechaentregaobra", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.NVarChar, SqlDbType.Int, SqlDbType.NVarChar, SqlDbType.DateTime, SqlDbType.BigInt}
        Dim valores() As Object = {d.nombreobra, d.nroentregas_obra, d.lugarentrega, d.fechaentregaobra, d.idproceso}
        sql.EjecutarProcedure("Str_tbl_Lic_Obras_Procesos_I", parametros, valores, tipoParametro, 5)
    End Sub
    Public Sub Actualizar(d As Ntbl_Lic_Obras_Procesos)
        Dim parametros() As Object = {"@idobras", "@nombreobra", "@nroentregas_obra", "@lugarentrega", "@fechaentregaobra", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.NVarChar, SqlDbType.Int, SqlDbType.NVarChar, SqlDbType.DateTime, SqlDbType.BigInt}
        Dim valores() As Object = {d.idobras, d.nombreobra, d.nroentregas_obra, d.lugarentrega, d.fechaentregaobra, d.idproceso}
        sql.EjecutarProcedure("Str_tbl_Lic_Obras_Procesos_U", parametros, valores, tipoParametro, 6)
    End Sub
    Public Function Agregar(d As Ntbl_Lic_Obras_Procesos, Retornatable As Boolean) As Ntbl_Lic_Obras_Procesos
        Dim parametros() As Object = {"@nombreobra", "@nroentregas_obra", "@lugarentrega", "@fechaentregaobra", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.NVarChar, SqlDbType.Int, SqlDbType.NVarChar, SqlDbType.DateTime, SqlDbType.BigInt}
        Dim valores() As Object = {d.nombreobra, d.nroentregas_obra, d.lugarentrega, d.fechaentregaobra, d.idproceso}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_Obras_Procesos_I_S", parametros, valores, tipoParametro, 5).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idobras = IIf(dt.Rows(0).Item("idobras") Is DBNull.Value, Nothing, dt.Rows(0).Item("idobras"))
            d.nombreobra = IIf(dt.Rows(0).Item("nombreobra") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreobra"))
            d.nroentregas_obra = IIf(dt.Rows(0).Item("nroentregas_obra") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroentregas_obra"))
            d.lugarentrega = IIf(dt.Rows(0).Item("lugarentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarentrega"))
            d.fechaentregaobra = IIf(dt.Rows(0).Item("fechaentregaobra") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaentregaobra"))
            d.idproceso = IIf(dt.Rows(0).Item("idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproceso"))
        Else
            d.idobras = Nothing
            d.nombreobra = Nothing
            d.nroentregas_obra = Nothing
            d.lugarentrega = Nothing
            d.fechaentregaobra = Nothing
            d.idproceso = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_Lic_Obras_Procesos, Retornatable As Boolean) As Ntbl_Lic_Obras_Procesos
        Dim parametros() As Object = {"@idobras", "@nombreobra", "@nroentregas_obra", "@lugarentrega", "@fechaentregaobra", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.NVarChar, SqlDbType.Int, SqlDbType.NVarChar, SqlDbType.DateTime, SqlDbType.BigInt}
        Dim valores() As Object = {d.idobras, d.nombreobra, d.nroentregas_obra, d.lugarentrega, d.fechaentregaobra, d.idproceso}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_Obras_Procesos_U_S", parametros, valores, tipoParametro, 6).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idobras = IIf(dt.Rows(0).Item("idobras") Is DBNull.Value, Nothing, dt.Rows(0).Item("idobras"))
            d.nombreobra = IIf(dt.Rows(0).Item("nombreobra") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreobra"))
            d.nroentregas_obra = IIf(dt.Rows(0).Item("nroentregas_obra") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroentregas_obra"))
            d.lugarentrega = IIf(dt.Rows(0).Item("lugarentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarentrega"))
            d.fechaentregaobra = IIf(dt.Rows(0).Item("fechaentregaobra") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaentregaobra"))
            d.idproceso = IIf(dt.Rows(0).Item("idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproceso"))
        Else
            d.idobras = Nothing
            d.nombreobra = Nothing
            d.nroentregas_obra = Nothing
            d.lugarentrega = Nothing
            d.fechaentregaobra = Nothing
            d.idproceso = Nothing

        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_Lic_Obras_Procesos)
        Dim parametros() As Object = {"@idobras"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idobras}
        sql.EjecutarProcedure("Str_tbl_Lic_Obras_Procesos_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_tbl_Lic_Obras_Procesos(d As Ntbl_Lic_Obras_Procesos) As Boolean
        Dim parametros() As Object = {"@idobras"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.idobras}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_tbl_Lic_Obras_Procesos", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idobras"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_Obras_Procesos_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_Lic_Obras_Procesos) As DataTable
        Dim parametros() As Object = {"@idobras"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.idobras}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_Obras_Procesos_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista_Pk(d As Ntbl_Lic_Obras_Procesos) As DataTable
        Dim sParametro As Object() = {"@IdProceso"}
        Dim typeParam As Object() = {SqlDbType.Int}
        Dim vParametro As Object() = {d.idproceso}
        Dim dataTable As DataTable = New DataTable()
        Return Me.sql.ProcedureSQL("Str_tbl_Lic_Obras_Procesos_PK", sParametro, vParametro, typeParam, 1).Tables(0)
    End Function
    Public Function Registro(d As Ntbl_Lic_Obras_Procesos) As Ntbl_Lic_Obras_Procesos
        Dim parametros() As Object = {"@idobras"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.idobras}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_Obras_Procesos_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idobras = IIf(dt.Rows(0).Item("idobras") Is DBNull.Value, Nothing, dt.Rows(0).Item("idobras"))
            d.nombreobra = IIf(dt.Rows(0).Item("nombreobra") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreobra"))
            d.nroentregas_obra = IIf(dt.Rows(0).Item("nroentregas_obra") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroentregas_obra"))
            d.lugarentrega = IIf(dt.Rows(0).Item("lugarentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarentrega"))
            d.fechaentregaobra = IIf(dt.Rows(0).Item("fechaentregaobra") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaentregaobra"))
            d.idproceso = IIf(dt.Rows(0).Item("idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproceso"))
        Else
            d.idobras = Nothing
            d.nombreobra = Nothing
            d.nroentregas_obra = Nothing
            d.lugarentrega = Nothing
            d.fechaentregaobra = Nothing
            d.idproceso = Nothing

        End If
        Return d
    End Function
#End Region


End Class
