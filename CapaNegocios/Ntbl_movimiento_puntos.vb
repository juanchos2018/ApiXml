Imports CapaDatos
Public Class Ntbl_movimiento_puntos
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property id As Long
    Public Property idcliente As String
    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numerodocumento As String
    Public Property puntos As Decimal
    Public Property estado As String
    Public Property fechacrea As System.DateTime
    Public Property usuariocrea As String
#End Region
#Region "Constructors"
    Public Sub New()
    End Sub
#End Region
#Region "Metodos"
    Public Sub Agregar(d As Ntbl_movimiento_puntos)

        Dim parametros() As Object = {"@idcliente", "@idtipodocumento", "@serie", "@numerodocumento", "@puntos", "@estado", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idcliente, d.idtipodocumento, d.serie, d.numerodocumento, d.puntos, d.estado, d.fechacrea, d.usuariocrea}
        sql.EjecutarProcedure("Str_tbl_movimiento_puntos_I", parametros, valores, tipoParametro, 8)
    End Sub
    Public Sub Actualizar(d As Ntbl_movimiento_puntos)
        Dim parametros() As Object = {"@id", "@idcliente", "@idtipodocumento", "@serie", "@numerodocumento", "@puntos", "@estado", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.id, d.idcliente, d.idtipodocumento, d.serie, d.numerodocumento, d.puntos, d.estado, d.fechacrea, d.usuariocrea}
        sql.EjecutarProcedure("Str_tbl_movimiento_puntos_U", parametros, valores, tipoParametro, 9)
    End Sub
    Public Function Agregar(d As Ntbl_movimiento_puntos, Retornatable As Boolean) As Ntbl_movimiento_puntos

        Dim parametros() As Object = {"@idcliente", "@idtipodocumento", "@serie", "@numerodocumento", "@puntos", "@estado", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idcliente, d.idtipodocumento, d.serie, d.numerodocumento, d.puntos, d.estado, d.fechacrea, d.usuariocrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_movimiento_puntos_I_S", parametros, valores, tipoParametro, 8).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.idcliente = IIf(dt.Rows(0).Item("idcliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcliente"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.puntos = IIf(dt.Rows(0).Item("puntos") Is DBNull.Value, Nothing, dt.Rows(0).Item("puntos"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.id = Nothing
            d.idcliente = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.puntos = Nothing
            d.estado = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_movimiento_puntos, Retornatable As Boolean) As Ntbl_movimiento_puntos

        Dim parametros() As Object = {"@Id", "@idcliente", "@idtipodocumento", "@serie", "@numerodocumento", "@puntos", "@estado", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.id, d.idcliente, d.idtipodocumento, d.serie, d.numerodocumento, d.puntos, d.estado, d.fechacrea, d.usuariocrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_movimiento_puntos_U_S", parametros, valores, tipoParametro, 9).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.idcliente = IIf(dt.Rows(0).Item("idcliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcliente"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.puntos = IIf(dt.Rows(0).Item("puntos") Is DBNull.Value, Nothing, dt.Rows(0).Item("puntos"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.id = Nothing
            d.idcliente = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.puntos = Nothing
            d.estado = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_movimiento_puntos)
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        sql.EjecutarProcedure("Str_tbl_movimiento_puntos_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_movimiento_puntos_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_movimiento_puntos) As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_movimiento_puntos_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Ntbl_movimiento_puntos) As Ntbl_movimiento_puntos
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_movimiento_puntos_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.idcliente = IIf(dt.Rows(0).Item("idcliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcliente"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.puntos = IIf(dt.Rows(0).Item("puntos") Is DBNull.Value, Nothing, dt.Rows(0).Item("puntos"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.id = Nothing
            d.idcliente = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.puntos = Nothing
            d.estado = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
#End Region

End Class
