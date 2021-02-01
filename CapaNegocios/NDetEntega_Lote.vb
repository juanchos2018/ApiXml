Imports CapaDatos
Public Class NDetEntega_Lote
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idEntrega As Long
    Public Property idAlmacen As String
    Public Property tipoDocumento As String
    Public Property numeroDocumento As String
    Public Property item As String
    Public Property idAgencia As String
    Public Property idArticulo As String
    Public Property cantidad As Decimal
    Public Property saldoEntrega As Decimal
    Public Property nroLote As String
    Public Property idLote As Long

#End Region

#Region "Constructors"
    Public Sub New()
    End Sub
#End Region
#Region "Metodos"
    Public Sub Agregar(d As NDetEntega_Lote)
        Dim parametros() As Object = {"@idAlmacen", "@tipoDocumento", "@numeroDocumento", "@item", "@idAgencia", "@idArticulo", "@cantidad", "@saldoEntrega", "@nroLote", "@idLote"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.BigInt}
        Dim valores() As Object = {d.idAlmacen, d.tipoDocumento, d.numeroDocumento, d.item, d.idAgencia, d.idArticulo, d.cantidad, d.saldoEntrega, d.nroLote, d.idLote}
        sql.EjecutarProcedure("Str_DetEntega_Lote_I", parametros, valores, tipoParametro, 10)
    End Sub
    Public Sub Actualizar(d As NDetEntega_Lote)
        Dim parametros() As Object = {"@IdEntrega", "@idAlmacen", "@tipoDocumento", "@numeroDocumento", "@item", "@idAgencia", "@idArticulo", "@cantidad", "@saldoEntrega", "@nroLote", "@idLote"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.BigInt}
        Dim valores() As Object = {d.idEntrega, d.idAlmacen, d.tipoDocumento, d.numeroDocumento, d.item, d.idAgencia, d.idArticulo, d.cantidad, d.saldoEntrega, d.nroLote, d.idLote}
        sql.EjecutarProcedure("Str_DetEntega_Lote_U", parametros, valores, tipoParametro, 11)
    End Sub
    Public Sub Eliminar(d As NDetEntega_Lote)
        Dim parametros() As Object = {"@idEntrega"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idEntrega}
        sql.EjecutarProcedure("Str_DetEntega_Lote_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idEntrega"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetEntega_Lote_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NDetEntega_Lote) As DataTable
        Dim parametros() As Object = {"@idEntrega"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idEntrega}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetEntega_Lote_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDetEntega_Lote) As NDetEntega_Lote
        Dim parametros() As Object = {"@idEntrega"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idEntrega}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetEntega_Lote_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idEntrega = IIf(dt.Rows(0).Item("idEntrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("idEntrega"))
            d.idAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.tipoDocumento = IIf(dt.Rows(0).Item("tipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoDocumento"))
            d.numeroDocumento = IIf(dt.Rows(0).Item("numeroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idAgencia = IIf(dt.Rows(0).Item("idAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAgencia"))
            d.idArticulo = IIf(dt.Rows(0).Item("idArticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idArticulo"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.saldoEntrega = IIf(dt.Rows(0).Item("saldoEntrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldoEntrega"))
            d.nroLote = IIf(dt.Rows(0).Item("nroLote") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroLote"))
            d.idLote = IIf(dt.Rows(0).Item("idLote") Is DBNull.Value, Nothing, dt.Rows(0).Item("idLote"))
        Else
            d.idEntrega = Nothing
            d.idAlmacen = Nothing
            d.tipoDocumento = Nothing
            d.numeroDocumento = Nothing
            d.item = Nothing
            d.idAgencia = Nothing
            d.idArticulo = Nothing
            d.cantidad = Nothing
            d.saldoEntrega = Nothing
            d.nroLote = Nothing
            d.idLote = Nothing
        End If
        Return d
    End Function
#End Region

End Class
