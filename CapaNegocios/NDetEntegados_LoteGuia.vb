Imports CapaDatos
Public Class NDetEntegados_LoteGuia
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idEntegado As Long
    Public Property idAlmacen As String
    Public Property tipoDocumento As String
    Public Property numeroDocumento As String
    Public Property item As String
    Public Property idAgencia As String
    Public Property idArticulo As String
    Public Property cantidad As Decimal
    Public Property saldoEntregado As Decimal
    Public Property nroLote As String
    Public Property idEntrega As Long
#End Region
#Region "Constructors"
    Public Sub New()
    End Sub
#End Region
#Region "Metodos"
    Public Sub Agregar(d As NDetEntegados_LoteGuia)
        Dim parametros() As Object = {"@idAlmacen", "@tipoDocumento", "@numeroDocumento", "@item", "@idAgencia", "@idArticulo", "@cantidad", "@saldoEntregado", "@nroLote", "@idEntrega"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.BigInt}
        Dim valores() As Object = {d.idAlmacen, d.tipoDocumento, d.numeroDocumento, d.item, d.idAgencia, d.idArticulo, d.cantidad, d.saldoEntregado, d.nroLote, d.idEntrega}
        sql.EjecutarProcedure("Str_DetEntegados_LoteGuia_I", parametros, valores, tipoParametro, 10)
    End Sub
    Public Sub Actualizar(d As NDetEntegados_LoteGuia)
        Dim parametros() As Object = {"@idEntegado", "@idAlmacen", "@tipoDocumento", "@numeroDocumento", "@item", "@idAgencia", "@idArticulo", "@cantidad", "@saldoEntregado", "@nroLote", "@idEntrega"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.BigInt}
        Dim valores() As Object = {d.idEntegado, d.idAlmacen, d.tipoDocumento, d.numeroDocumento, d.item, d.idAgencia, d.idArticulo, d.cantidad, d.saldoEntregado, d.nroLote, d.idEntrega}
        sql.EjecutarProcedure("Str_DetEntegados_LoteGuia_U", parametros, valores, tipoParametro, 10)
    End Sub
    Public Sub Eliminar(d As NDetEntegados_LoteGuia)
        Dim parametros() As Object = {"@idEntegado"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idEntegado}
        sql.EjecutarProcedure("Str_DetEntegados_LoteGuia_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idEntegado"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetEntegados_LoteGuia_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NDetEntegados_LoteGuia) As DataTable
        Dim parametros() As Object = {"@idEntegado"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idEntegado}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetEntegados_LoteGuia_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDetEntegados_LoteGuia) As NDetEntegados_LoteGuia
        Dim parametros() As Object = {"@idEntegado"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idEntegado}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetEntegados_LoteGuia_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idEntegado = IIf(dt.Rows(0).Item("idEntegado") Is DBNull.Value, Nothing, dt.Rows(0).Item("idEntegado"))
            d.idAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.tipoDocumento = IIf(dt.Rows(0).Item("tipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoDocumento"))
            d.numeroDocumento = IIf(dt.Rows(0).Item("numeroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idAgencia = IIf(dt.Rows(0).Item("idAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAgencia"))
            d.idArticulo = IIf(dt.Rows(0).Item("idArticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idArticulo"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.saldoEntregado = IIf(dt.Rows(0).Item("saldoEntregado") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldoEntregado"))
            d.nroLote = IIf(dt.Rows(0).Item("nroLote") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroLote"))
            d.idEntrega = IIf(dt.Rows(0).Item("idEntrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("idEntrega"))
        Else
            d.idEntegado = Nothing
            d.idAlmacen = Nothing
            d.tipoDocumento = Nothing
            d.numeroDocumento = Nothing
            d.item = Nothing
            d.idAgencia = Nothing
            d.idArticulo = Nothing
            d.cantidad = Nothing
            d.saldoEntregado = Nothing
            d.nroLote = Nothing
            d.idEntrega = Nothing
        End If
        Return d
    End Function
    Public Function Entregados_LoteGuia(idagencia As String, idalmacen As String, idTipoDocumento As String, numerodocumento As String) As DataTable
        Dim cad As String = " SELECT    IdEntrega,NroLote,Cantidad,SaldoEntrega,IdArticulo from DetEntega_Lote "
        cad += " where rtrim(IdAlmacen)='" & idalmacen & "' AND rtrim(IdAgencia)='" & idagencia & "' AND TipoDocumento='" & idTipoDocumento & "' "
        cad += " AND rtrim(NumeroDocumento)='" & numerodocumento & "'"
        Return sql.EjecutarConsulta("lote", cad).Tables(0)
    End Function
#End Region

End Class
