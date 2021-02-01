Imports CapaDatos
Public Class NDetalleComprobanteCompra
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idAgencia As String
    Public Property idTipoDocumento As String
    Public Property serie As String
    Public Property numeroDocumento As String
    Public Property item As String
    Public Property idArticulo As String
    Public Property descripcion As String
    Public Property cantidad As Decimal
    Public Property unidad As String
    Public Property precioUnitario As Decimal
    Public Property precioSIGV As Decimal
    Public Property importeDescuento As Decimal
    Public Property utilidad As Decimal
    Public Property importeUtilidad As Decimal
    Public Property iGV As Decimal
    Public Property importeIGV As Decimal
    Public Property importeUS As Decimal
    Public Property importeMN As Decimal
    Public Property descuento1 As Decimal
    Public Property importeDescuento1 As Decimal
    Public Property descuento2 As Decimal
    Public Property importeDescuento2 As Decimal
    Public Property descuento3 As Decimal
    Public Property importeDescuento3 As Decimal
    Public Property estado As String
    Public Property idAlmacen As String
    Public Property usuarioCrea As String
    Public Property fechaCrea As System.DateTime
    Public Property usuarioMod As String
    Public Property fechaMod As System.DateTime
    Public Property idProveedor As String
    Public Property identrada As String
    Public Property saldo As Decimal
    Public Property nroLote As String
    Public Property saldoLote As Decimal

#End Region
#Region "Constructors"
    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NDetalleComprobanteCompra)
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@item", "@idArticulo", "@descripcion", "@cantidad", "@unidad", "@precioUnitario", "@precioSIGV", "@importeDescuento", "@utilidad", "@importeUtilidad", "@iGV", "@importeIGV", "@importeUS", "@importeMN", "@descuento1", "@importeDescuento1", "@descuento2", "@importeDescuento2", "@descuento3", "@importeDescuento3", "@estado", "@idAlmacen", "@usuarioCrea", "@fechaCrea", "@usuarioMod", "@fechaMod", "@idProveedor", "@identrada", "@saldo", "@nroLote", "@saldoLote"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal}
        Dim valores() As Object = {d.idAgencia, d.idTipoDocumento, d.serie, d.numeroDocumento, d.item, d.idArticulo, d.descripcion, d.cantidad, d.unidad, d.precioUnitario, d.precioSIGV, d.importeDescuento, d.utilidad, d.importeUtilidad, d.iGV, d.importeIGV, d.importeUS, d.importeMN, d.descuento1, d.importeDescuento1, d.descuento2, d.importeDescuento2, d.descuento3, d.importeDescuento3, d.estado, d.idAlmacen, d.usuarioCrea, d.fechaCrea, d.usuarioMod, d.fechaMod, d.idProveedor, d.identrada, d.saldo, d.nroLote, d.saldoLote}
        sql.EjecutarProcedure("Str_detallecomprobantecompra_I", parametros, valores, tipoParametro, 35)
    End Sub
    Public Sub Actualizar(d As NDetalleComprobanteCompra)
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@item", "@idArticulo", "@descripcion", "@cantidad", "@unidad", "@precioUnitario", "@precioSIGV", "@importeDescuento", "@utilidad", "@importeUtilidad", "@iGV", "@importeIGV", "@importeUS", "@importeMN", "@descuento1", "@importeDescuento1", "@descuento2", "@importeDescuento2", "@descuento3", "@importeDescuento3", "@estado", "@idAlmacen", "@usuarioCrea", "@fechaCrea", "@usuarioMod", "@fechaMod", "@idProveedor", "@identrada", "@saldo", "@nroLote", "@saldoLote"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal}
        Dim valores() As Object = {d.idAgencia, d.idTipoDocumento, d.serie, d.numeroDocumento, d.item, d.idArticulo, d.descripcion, d.cantidad, d.unidad, d.precioUnitario, d.precioSIGV, d.importeDescuento, d.utilidad, d.importeUtilidad, d.iGV, d.importeIGV, d.importeUS, d.importeMN, d.descuento1, d.importeDescuento1, d.descuento2, d.importeDescuento2, d.descuento3, d.importeDescuento3, d.estado, d.idAlmacen, d.usuarioCrea, d.fechaCrea, d.usuarioMod, d.fechaMod, d.idProveedor, d.identrada, d.saldo, d.nroLote, d.saldoLote}
        sql.EjecutarProcedure("Str_detallecomprobantecompra_U", parametros, valores, tipoParametro, 35)
    End Sub
    Public Sub Eliminar(d As NDetalleComprobanteCompra)
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@item", "@idArticulo", "@idAlmacen", "@idProveedor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char}
        Dim valores() As Object = {d.idAgencia, d.idTipoDocumento, d.serie, d.numeroDocumento, d.item, d.idArticulo, d.idAlmacen, d.idProveedor}
        sql.EjecutarProcedure("Str_detallecomprobantecompra_D", parametros, valores, tipoParametro, 8)
    End Sub

    Public Function ListaLote(idarticulo As String) As DataTable
        'Dim cadena As String = " select IdAlmacen,IdAgencia,IdTipoDocumento,Serie,NumeroDocumento,IdProveedor,IdArticulo,Cantidad,NroLote,SaldoLote,item from detallecomprobantecompra "
        'cadena += " where Nrolote is not null and idarticulo='" & idarticulo & "'"
        Dim cadena As String = " select IdAlmacen,IdAgencia,IdTipoDocumento,Serie,NumeroDocumento,IdProveedor,IdArticulo,Cantidad,NroLote,saldo as SaldoLote,item,idLote from detcompra_lote "
        cadena += " where Nrolote is not null and idarticulo='" & idarticulo & "'"
        Return sql.EjecutarConsulta("df", cadena).Tables(0)
    End Function
    Public Function ListaLote(idarticulo As String, IdAlmacen As String) As DataTable
        Dim cadena As String = " select  IdAlmacen,IdAgencia,IdTipoDocumento,Serie,NumeroDocumento,IdProveedor,IdArticulo,Cantidad,NroLote,saldo as SaldoLote,item,idLote from detcompra_lote "
        cadena += " where Nrolote is not null and idarticulo='" & idarticulo & "' and IdAlmacen='" & IdAlmacen & "'"
        Return sql.EjecutarConsulta("df", cadena).Tables(0)
    End Function
    Public Function Lista_Lote_Serie(idarticulo As String, idalmacen As String) As DataTable
        Dim cadena As String = " select    IdAlmacen, IdArticulo, NroSerie, Cantidad, TipoDocRef, NroDocRef, FechaRef, IdAgencia  from StockSerie "
        cadena += " where idarticulo='" & idarticulo.Trim & "' and idalmacen='" & idalmacen & "'"
        Return sql.EjecutarConsulta("df", cadena).Tables(0)
    End Function

    ''' <summary>
    ''' Elimina un registro en funcion a su primary key de cabecera
    ''' </summary>
    ''' <param name="d"></param>
    Public Sub Eliminar(idalmacen As String, Serie As String, idtipodocumento As String, numeroDocumento As String, Idagencia As String, idproveedor As String)
        'sql.Eliminar_Items("DetalleComprobantecompra", "IdAlmacen='" & idalmacen & "' and Serie='" & Serie & "' and IdTipoDocumento='" & idtipodocumento & "' and NumeroDocumento='" & numeroDocumento & "' and IdAgencia='" & Idagencia & "' and idproveedor='" & idproveedor & "'")
    End Sub

    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@item", "@idArticulo", "@idAlmacen", "@idProveedor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleComprobanteCompra_S", parametros, valores, tipoParametro, 8).Tables(0)
        Return dt
    End Function
    ''' <summary>
    ''' Retorna una lista de los items del detalle de la factura
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    ''' 

    Public Function Lista(d As NDetalleComprobanteCompra) As DataTable
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@item", "@idArticulo", "@idAlmacen", "@idProveedor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAgencia, d.IdTipoDocumento, d.Serie, d.NumeroDocumento, DBNull.Value, DBNull.Value, d.IdAlmacen, d.IdProveedor}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleComprobanteCompra_S", parametros, valores, tipoParametro, 8).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDetalleComprobanteCompra) As NDetalleComprobanteCompra
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@item", "@idArticulo", "@idAlmacen", "@idProveedor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char}
        Dim valores() As Object = {d.idAgencia, d.idTipoDocumento, d.serie, d.numeroDocumento, d.item, d.idArticulo, d.idAlmacen, d.idProveedor}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_detallecomprobantecompra_S", parametros, valores, tipoParametro, 8).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idAgencia = IIf(dt.Rows(0).Item("idAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAgencia"))
            d.idTipoDocumento = IIf(dt.Rows(0).Item("idTipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoDocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numeroDocumento = IIf(dt.Rows(0).Item("numeroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idArticulo = IIf(dt.Rows(0).Item("idArticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idArticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.unidad = IIf(dt.Rows(0).Item("unidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidad"))
            d.precioUnitario = IIf(dt.Rows(0).Item("precioUnitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioUnitario"))
            d.precioSIGV = IIf(dt.Rows(0).Item("precioSIGV") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioSIGV"))
            d.importeDescuento = IIf(dt.Rows(0).Item("importeDescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeDescuento"))
            d.utilidad = IIf(dt.Rows(0).Item("utilidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("utilidad"))
            d.importeUtilidad = IIf(dt.Rows(0).Item("importeUtilidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeUtilidad"))
            d.iGV = IIf(dt.Rows(0).Item("iGV") Is DBNull.Value, Nothing, dt.Rows(0).Item("iGV"))
            d.importeIGV = IIf(dt.Rows(0).Item("importeIGV") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeIGV"))
            d.importeUS = IIf(dt.Rows(0).Item("importeUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeUS"))
            d.importeMN = IIf(dt.Rows(0).Item("importeMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeMN"))
            d.descuento1 = IIf(dt.Rows(0).Item("descuento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento1"))
            d.importeDescuento1 = IIf(dt.Rows(0).Item("importeDescuento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeDescuento1"))
            d.descuento2 = IIf(dt.Rows(0).Item("descuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento2"))
            d.importeDescuento2 = IIf(dt.Rows(0).Item("importeDescuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeDescuento2"))
            d.descuento3 = IIf(dt.Rows(0).Item("descuento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento3"))
            d.importeDescuento3 = IIf(dt.Rows(0).Item("importeDescuento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeDescuento3"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.idAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.usuarioCrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.fechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.usuarioMod = IIf(dt.Rows(0).Item("usuarioMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioMod"))
            d.fechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.idProveedor = IIf(dt.Rows(0).Item("idProveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idProveedor"))
            d.identrada = IIf(dt.Rows(0).Item("identrada") Is DBNull.Value, Nothing, dt.Rows(0).Item("identrada"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.nroLote = IIf(dt.Rows(0).Item("nroLote") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroLote"))
            d.saldoLote = IIf(dt.Rows(0).Item("saldoLote") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldoLote"))
        Else
            d.idAgencia = Nothing
            d.idTipoDocumento = Nothing
            d.serie = Nothing
            d.numeroDocumento = Nothing
            d.item = Nothing
            d.idArticulo = Nothing
            d.descripcion = Nothing
            d.cantidad = Nothing
            d.unidad = Nothing
            d.precioUnitario = Nothing
            d.precioSIGV = Nothing
            d.importeDescuento = Nothing
            d.utilidad = Nothing
            d.importeUtilidad = Nothing
            d.iGV = Nothing
            d.importeIGV = Nothing
            d.importeUS = Nothing
            d.importeMN = Nothing
            d.descuento1 = Nothing
            d.importeDescuento1 = Nothing
            d.descuento2 = Nothing
            d.importeDescuento2 = Nothing
            d.descuento3 = Nothing
            d.importeDescuento3 = Nothing
            d.estado = Nothing
            d.idAlmacen = Nothing
            d.usuarioCrea = Nothing
            d.fechaCrea = Nothing
            d.usuarioMod = Nothing
            d.fechaMod = Nothing
            d.idProveedor = Nothing
            d.identrada = Nothing
            d.saldo = Nothing
            d.nroLote = Nothing
            d.saldoLote = Nothing
        End If
        Return d
    End Function
    Public Function DetalleAsiento(d As NDetalleComprobanteCompra) As DataTable
        Dim cad_asientovtas As String = " SELECT '' as IdSubdiario,(SELECT IDCOMPRA FROM CUENTAEXISTENCIA WHERE IDCUENTA=A.IDCUENTACONTABLE) AS IDCUENTA,"
        cad_asientovtas += " (SELECT  IDCUENTA FROM CUENTAEXISTENCIA WHERE IDCUENTA=A.IDCUENTACONTABLE) AS IDCTAMER, "
        cad_asientovtas += " (SELECT  IdCtaVariacion FROM CUENTAEXISTENCIA WHERE IDCUENTA=A.IDCUENTACONTABLE) AS IDCTAVAR "
        cad_asientovtas += " ,'D' AS DEBEHABER, D.IdTipoDocumento, RIGHT(D.Serie, 4) AS serie,d.NumeroDocumento,PRECIOSIGV ,(left(rtrim(d.descripcion),20)+' Cant:'+convert(varchar(10),cast(Cantidad as int),112) ) glosa,(SELECT IDCENTROCOSTO FROM CUENTAEXISTENCIA WHERE IDCUENTA=A.IDCUENTACONTABLE) AS IdCosto FROM DetalleComprobanteCompra D INNER JOIN ARTICULO A   "
        cad_asientovtas += " ON D.IDARTICULO=A.IDARTICULO "
        cad_asientovtas += " where d.IdAlmacen='" & d.idAlmacen & "' and d.IdProveedor='" & d.idProveedor & "' and d.IdTipoDocumento='" & d.idTipoDocumento & "' and d.Serie='" & d.serie & "' and d.NumeroDocumento='" & d.numeroDocumento & "'"
        Return sql.EjecutarConsulta("d", cad_asientovtas).Tables(0)
    End Function
#End Region

End Class
