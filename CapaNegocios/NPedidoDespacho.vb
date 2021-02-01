Imports CapaDatos
Public Class NPedidoDespacho
    Dim sql As New ClsConexion
    Public Property Id As Integer
    Public Property IdAgencia As String
    Public Property IdAlmacen As String
    Public Property TipoDocumentoOrigen As String
    Public Property TipoDocumentoDestino As String
    Public Property SerieOrigen As String
    Public Property SerieDestino As String
    Public Property NumeroDocumentoOrigen As String
    Public Property NumeroDocumentoDestino As String
    Public Sub Agregar(d As NPedidoDespacho)
        Dim parametros() As Object = {"@IdAgencia", "@IdAlmacen", "@TipoDocumentoOrigen", "@TipoDocumentoDestino", "@SerieOrigen", "@SerieDestino", "@NumeroDocumentoOrigen", "@NumeroDocumentoDestino"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAgencia, d.IdAlmacen, d.TipoDocumentoOrigen, d.TipoDocumentoDestino, d.SerieOrigen, d.SerieDestino, d.NumeroDocumentoOrigen, d.NumeroDocumentoDestino}
        sql.EjecutarProcedure("Str_PedidoDespacho_I", parametros, valores, tipoParametro, 8)
    End Sub
    Public Sub Eliminar(d As NPedidoDespacho)
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.Id}
        sql.EjecutarProcedure("Str_PedidoDespacho_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function ListaPorDestino(d As NPedidoDespacho) As DataTable
        Dim cad As String = "select [IdAgencia], [IdAlmacen], [TipoDocumentoOrigen], [TipoDocumentoDestino],
       [SerieOrigen], [SerieDestino], [NumeroDocumentoOrigen], [NumeroDocumentoDestino] 
	   from PedidoDespacho where NumeroDocumentoDestino='" + d.NumeroDocumentoDestino + "'"
        Return sql.EjecutarConsulta("d", cad).Tables(0)
    End Function
    Public Function ListaPorOrigen(d As NPedidoDespacho) As DataTable
        Dim cad As String = "select [IdAgencia], [IdAlmacen], [TipoDocumentoOrigen], [TipoDocumentoDestino],
       [SerieOrigen], [SerieDestino], [NumeroDocumentoOrigen], [NumeroDocumentoDestino] 
	   from PedidoDespacho where NumeroDocumentoOrigen='" + d.NumeroDocumentoOrigen + "'"
        Return sql.EjecutarConsulta("d", cad).Tables(0)
    End Function
    Public Function Registro(d As NPedidoDespacho) As NPedidoDespacho
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.Id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_PedidoDespacho_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdAlmacen = IIf(dt.Rows(0).Item("IdAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.IdAgencia = IIf(dt.Rows(0).Item("IdAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("agencia"))
            d.TipoDocumentoOrigen = IIf(dt.Rows(0).Item("TipoDocumentoOrigen") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.TipoDocumentoDestino = IIf(dt.Rows(0).Item("TipoDocumentoDestino") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.SerieOrigen = IIf(dt.Rows(0).Item("SerieOrigen") Is DBNull.Value, Nothing, dt.Rows(0).Item("distrito"))
            d.SerieDestino = IIf(dt.Rows(0).Item("SerieDestino") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono"))
            d.NumeroDocumentoOrigen = IIf(dt.Rows(0).Item("NumeroDocumentoOrigen") Is DBNull.Value, Nothing, dt.Rows(0).Item("distrito"))
            d.NumeroDocumentoDestino = IIf(dt.Rows(0).Item("NumeroDocumentoDestino") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono"))
        Else
            d.IdAlmacen = Nothing
            d.IdAgencia = Nothing
            d.TipoDocumentoOrigen = Nothing
            d.TipoDocumentoDestino = Nothing
            d.SerieOrigen = Nothing
            d.SerieDestino = Nothing
            d.NumeroDocumentoOrigen = Nothing
            d.NumeroDocumentoDestino = Nothing
        End If
        Return d
    End Function
End Class
