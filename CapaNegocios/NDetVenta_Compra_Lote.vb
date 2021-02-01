Imports CapaDatos
Public Class NDetVenta_Compra_Lote
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idLote As Long
    Public Property idAlmacen As String
    Public Property idAgencia As String
    Public Property idProveedor As String
    Public Property idTipoDocumento As String
    Public Property serie As String
    Public Property numeroDocumento As String
    Public Property idArticulo As String
    Public Property item As String
    Public Property cantidad As Decimal
    Public Property nroLote As String
    Public Property idAlmacenCL As String
    Public Property idAgenciaCL As String
    Public Property idCliente As String
    Public Property idTipoDocumentoCL As String
    Public Property serieCL As String
    Public Property numeroDocumentoCL As String
    Public Property proyecto As String
    Public Property itemCL As String
    Public Property idArticuloCL As String
    Public Property saldoLote As Decimal
    Public Property idLoteCompra As Long

#End Region
#Region "Constructors"
    Public Sub New()
    End Sub
#End Region
#Region "Metodos"
    Public Sub Agregar(d As NDetVenta_Compra_Lote)
        Dim parametros() As Object = {"@idAlmacen", "@idAgencia", "@idProveedor", "@idTipoDocumento", "@serie", "@numeroDocumento", "@idArticulo", "@item", "@cantidad", "@nroLote", "@idAlmacenCL", "@idAgenciaCL", "@idCliente", "@idTipoDocumentoCL", "@serieCL", "@numeroDocumentoCL", "@proyecto", "@itemCL", "@idArticuloCL", "@saldoLote", "@idLoteCompra"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.BigInt}
        Dim valores() As Object = {d.idAlmacen, d.idAgencia, d.idProveedor, d.idTipoDocumento, d.serie, d.numeroDocumento, d.idArticulo, d.item, d.cantidad, d.nroLote, d.idAlmacenCL, d.idAgenciaCL, d.idCliente, d.idTipoDocumentoCL, d.serieCL, d.numeroDocumentoCL, d.proyecto, d.itemCL, d.idArticuloCL, d.saldoLote, d.idLoteCompra}
        sql.EjecutarProcedure("Str_DetVenta_Compra_Lote_I", parametros, valores, tipoParametro, 21)
    End Sub
    Public Sub Actualizar(d As NDetVenta_Compra_Lote)
        Dim parametros() As Object = {"@idLote", "@idAlmacen", "@idAgencia", "@idProveedor", "@idTipoDocumento", "@serie", "@numeroDocumento", "@idArticulo", "@item", "@cantidad", "@nroLote", "@idAlmacenCL", "@idAgenciaCL", "@idCliente", "@idTipoDocumentoCL", "@serieCL", "@numeroDocumentoCL", "@proyecto", "@itemCL", "@idArticuloCL", "@saldoLote", "@idLoteCompra"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.BigInt}
        Dim valores() As Object = {d.idLote, d.idAlmacen, d.idAgencia, d.idProveedor, d.idTipoDocumento, d.serie, d.numeroDocumento, d.idArticulo, d.item, d.cantidad, d.nroLote, d.idAlmacenCL, d.idAgenciaCL, d.idCliente, d.idTipoDocumentoCL, d.serieCL, d.numeroDocumentoCL, d.proyecto, d.itemCL, d.idArticuloCL, d.saldoLote, d.idLoteCompra}
        sql.EjecutarProcedure("Str_DetVenta_Compra_Lote_U", parametros, valores, tipoParametro, 22)
    End Sub
    Public Sub Eliminar(d As NDetVenta_Compra_Lote)
        Dim parametros() As Object = {"@idLote"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idLote}
        sql.EjecutarProcedure("Str_DetVenta_Compra_Lote_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idLote"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetVenta_Compra_Lote_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NDetVenta_Compra_Lote) As DataTable
        Dim parametros() As Object = {"@idLote"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idLote}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetVenta_Compra_Lote_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDetVenta_Compra_Lote) As NDetVenta_Compra_Lote
        Dim parametros() As Object = {"@idLote"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idLote}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetVenta_Compra_Lote_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idLote = IIf(dt.Rows(0).Item("idLote") Is DBNull.Value, Nothing, dt.Rows(0).Item("idLote"))
            d.idAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.idAgencia = IIf(dt.Rows(0).Item("idAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAgencia"))
            d.idProveedor = IIf(dt.Rows(0).Item("idProveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idProveedor"))
            d.idTipoDocumento = IIf(dt.Rows(0).Item("idTipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoDocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numeroDocumento = IIf(dt.Rows(0).Item("numeroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento"))
            d.idArticulo = IIf(dt.Rows(0).Item("idArticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idArticulo"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.nroLote = IIf(dt.Rows(0).Item("nroLote") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroLote"))
            d.idAlmacenCL = IIf(dt.Rows(0).Item("idAlmacenCL") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacenCL"))
            d.idAgenciaCL = IIf(dt.Rows(0).Item("idAgenciaCL") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAgenciaCL"))
            d.idCliente = IIf(dt.Rows(0).Item("idCliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCliente"))
            d.idTipoDocumentoCL = IIf(dt.Rows(0).Item("idTipoDocumentoCL") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoDocumentoCL"))
            d.serieCL = IIf(dt.Rows(0).Item("serieCL") Is DBNull.Value, Nothing, dt.Rows(0).Item("serieCL"))
            d.numeroDocumentoCL = IIf(dt.Rows(0).Item("numeroDocumentoCL") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumentoCL"))
            d.proyecto = IIf(dt.Rows(0).Item("proyecto") Is DBNull.Value, Nothing, dt.Rows(0).Item("proyecto"))
            d.itemCL = IIf(dt.Rows(0).Item("itemCL") Is DBNull.Value, Nothing, dt.Rows(0).Item("itemCL"))
            d.idArticuloCL = IIf(dt.Rows(0).Item("idArticuloCL") Is DBNull.Value, Nothing, dt.Rows(0).Item("idArticuloCL"))
            d.saldoLote = IIf(dt.Rows(0).Item("saldoLote") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldoLote"))
            d.idLoteCompra = IIf(dt.Rows(0).Item("idLoteCompra") Is DBNull.Value, Nothing, dt.Rows(0).Item("idLoteCompra"))
        Else
            d.idLote = Nothing
            d.idAlmacen = Nothing
            d.idAgencia = Nothing
            d.idProveedor = Nothing
            d.idTipoDocumento = Nothing
            d.serie = Nothing
            d.numeroDocumento = Nothing
            d.idArticulo = Nothing
            d.item = Nothing
            d.cantidad = Nothing
            d.nroLote = Nothing
            d.idAlmacenCL = Nothing
            d.idAgenciaCL = Nothing
            d.idCliente = Nothing
            d.idTipoDocumentoCL = Nothing
            d.serieCL = Nothing
            d.numeroDocumentoCL = Nothing
            d.proyecto = Nothing
            d.itemCL = Nothing
            d.idArticuloCL = Nothing
            d.saldoLote = Nothing
            d.idLoteCompra = Nothing
        End If
        Return d
    End Function
    Public Function DetalleLote(idagencia As String, idalmacen As String, idTipoDocumento As String, numerodocumento As String, idarticulo As String) As DataTable
        Dim cad As String = " SELECT    IdLote,NroLote,Cantidad,SaldoLote FROM  DetVenta_Compra_Lote "
        cad += " where rtrim(IdAlmacenCL)='" & idalmacen & "' AND rtrim(IdAgenciaCL)='" & idagencia & "' AND IdTipoDocumentoCL='" & idTipoDocumento & "' "
        cad += " AND rtrim(SerieCL)+rtrim(NumeroDocumentoCL)='" & numerodocumento & "' AND rtrim(IdArticuloCL)='" & idarticulo & "' "
        Return sql.EjecutarConsulta("lote", cad).Tables(0)
    End Function
    Public Function DetalleLote(idagencia As String, idalmacen As String, idTipoDocumento As String, numerodocumento As String) As DataTable
        Dim cad As String = " SELECT    IdLote,NroLote,Cantidad,SaldoLote,IdArticulo FROM  DetVenta_Compra_Lote "
        cad += " where rtrim(IdAlmacenCL)='" & idalmacen & "' AND rtrim(IdAgenciaCL)='" & idagencia & "' AND IdTipoDocumentoCL='" & idTipoDocumento & "' "
        cad += " AND rtrim(SerieCL)+rtrim(NumeroDocumentoCL)='" & numerodocumento & "' and SaldoLote<>0"
        Return sql.EjecutarConsulta("lote", cad).Tables(0)
    End Function
    Public Function DetalleLote(idagencia As String, idalmacen As String, idTipoDocumento As String, numerodocumento As String, idarticulo As String, serie As String) As DataTable
        Dim cad As String = " SELECT    IdLote, IdAlmacen, IdAgencia, IdProveedor, IdTipoDocumento, Serie, NumeroDocumento, IdArticulo, item, Cantidad, NroLote, IdAlmacenCL, IdAgenciaCL, IdCliente, IdTipoDocumentoCL, SerieCL,  NumeroDocumentoCL, Proyecto, ItemCL, IdArticuloCL, SaldoLote,IdLoteCompra FROM  DetVenta_Compra_Lote "
        cad += " where rtrim(IdAlmacenCL)='" & idalmacen & "' AND rtrim(IdAgenciaCL)='" & idagencia & "' AND IdTipoDocumentoCL='" & idTipoDocumento & "' "
        cad += " AND rtrim(SerieCL)='" & serie & "' and rtrim(NumeroDocumentoCL)='" & numerodocumento & "'  and SaldoLote<>0 "
        Return sql.EjecutarConsulta("lote", cad).Tables(0)
    End Function
    Public Function DetalleLote(idagencia As String, idalmacen As String, idTipoDocumento As String, numerodocumento As String, idarticulo As String, serie As String, NroLote As String) As DataTable
        Dim cad As String = " SELECT    IdLote, IdAlmacen, IdAgencia, IdProveedor, IdTipoDocumento, Serie, NumeroDocumento, IdArticulo, item, Cantidad, NroLote, IdAlmacenCL, IdAgenciaCL, IdCliente, IdTipoDocumentoCL, SerieCL,  NumeroDocumentoCL, Proyecto, ItemCL, IdArticuloCL, SaldoLote ,IdLoteCompra FROM  DetVenta_Compra_Lote "
        cad += " where rtrim(IdAlmacenCL)='" & idalmacen & "' AND rtrim(IdAgenciaCL)='" & idagencia & "' AND IdTipoDocumentoCL='" & idTipoDocumento & "' "
        cad += " AND rtrim(SerieCL)='" & serie & "' and rtrim(NumeroDocumentoCL)='" & numerodocumento & "' and SaldoLote<>0"
        Return sql.EjecutarConsulta("lote", cad).Tables(0)
    End Function

#End Region
End Class
