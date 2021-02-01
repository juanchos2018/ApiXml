Imports CapaDatos
Public Class Ntbl_TipoMoneda
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idtipomoneda As String
    Public Property tipomoneda As String
    Public Property met_conver As String
    Public Property tipomoneda_alter As String
    Public Property fechacrea As System.DateTime
    Public Property usuariocrea As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As Ntbl_TipoMoneda)

        Dim parametros() As Object = {"@idtipomoneda", "@tipomoneda", "@met_conver", "@tipomoneda_alter", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipomoneda, d.tipomoneda, d.met_conver, d.tipomoneda_alter, d.fechacrea, d.usuariocrea}
        sql.EjecutarProcedure("Str_tbl_TipoMoneda_I", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Actualizar(d As Ntbl_TipoMoneda)
        Dim parametros() As Object = {"@idtipomoneda", "@tipomoneda", "@met_conver", "@tipomoneda_alter", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipomoneda, d.tipomoneda, d.met_conver, d.tipomoneda_alter, d.fechacrea, d.usuariocrea}
        sql.EjecutarProcedure("Str_tbl_TipoMoneda_U", parametros, valores, tipoParametro, 6)
    End Sub
    Public Function Agregar(d As Ntbl_TipoMoneda, Retornatable As Boolean) As Ntbl_TipoMoneda

        Dim parametros() As Object = {"@idtipomoneda", "@tipomoneda", "@met_conver", "@tipomoneda_alter", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipomoneda, d.tipomoneda, d.met_conver, d.tipomoneda_alter, d.fechacrea, d.usuariocrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_TipoMoneda_I_S", parametros, valores, tipoParametro, 6).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idtipomoneda = IIf(dt.Rows(0).Item("idtipomoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipomoneda"))
            d.tipomoneda = IIf(dt.Rows(0).Item("tipomoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipomoneda"))
            d.met_conver = IIf(dt.Rows(0).Item("met_conver") Is DBNull.Value, Nothing, dt.Rows(0).Item("met_conver"))
            d.tipomoneda_alter = IIf(dt.Rows(0).Item("tipomoneda_alter") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipomoneda_alter"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.idtipomoneda = Nothing
            d.tipomoneda = Nothing
            d.met_conver = Nothing
            d.tipomoneda_alter = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_TipoMoneda, Retornatable As Boolean) As Ntbl_TipoMoneda

        Dim parametros() As Object = {"@idtipomoneda", "@tipomoneda", "@met_conver", "@tipomoneda_alter", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.usuariocrea = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_TipoMoneda_U_S", parametros, valores, tipoParametro, 6).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idtipomoneda = IIf(dt.Rows(0).Item("idtipomoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipomoneda"))
            d.tipomoneda = IIf(dt.Rows(0).Item("tipomoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipomoneda"))
            d.met_conver = IIf(dt.Rows(0).Item("met_conver") Is DBNull.Value, Nothing, dt.Rows(0).Item("met_conver"))
            d.tipomoneda_alter = IIf(dt.Rows(0).Item("tipomoneda_alter") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipomoneda_alter"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.idtipomoneda = Nothing
            d.tipomoneda = Nothing
            d.met_conver = Nothing
            d.tipomoneda_alter = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_TipoMoneda)
        Dim parametros() As Object = {"@idtipomoneda"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipomoneda}
        sql.EjecutarProcedure("Str_tbl_TipoMoneda_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idtipomoneda"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_TipoMoneda_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_TipoMoneda) As DataTable
        Dim parametros() As Object = {"@idtipomoneda"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipomoneda}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_TipoMoneda_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Ntbl_TipoMoneda) As Ntbl_TipoMoneda
        Dim parametros() As Object = {"@idtipomoneda"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipomoneda}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_TipoMoneda_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idtipomoneda = IIf(dt.Rows(0).Item("idtipomoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipomoneda"))
            d.tipomoneda = IIf(dt.Rows(0).Item("tipomoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipomoneda"))
            d.met_conver = IIf(dt.Rows(0).Item("met_conver") Is DBNull.Value, Nothing, dt.Rows(0).Item("met_conver"))
            d.tipomoneda_alter = IIf(dt.Rows(0).Item("tipomoneda_alter") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipomoneda_alter"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.idtipomoneda = Nothing
            d.tipomoneda = Nothing
            d.met_conver = Nothing
            d.tipomoneda_alter = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
#End Region
End Class
