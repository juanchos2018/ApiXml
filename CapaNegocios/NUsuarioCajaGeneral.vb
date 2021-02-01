Imports CapaDatos
Public Class NUsuarioCajaGeneral
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _idCaja As String
    Private _idUsuario As String
    Private _idAlmacen As String
    Private _estado As Boolean
    Private _correlativo As Integer
    Private _nombreCaja As String
    Private _habilitar As Boolean

#End Region

#Region "Properties"

    Public Property IdCaja As String
        Get
            Return _idCaja
        End Get
        Set
            _idCaja = Value
        End Set
    End Property

    Public Property IdUsuario As String
        Get
            Return _idUsuario
        End Get
        Set
            _idUsuario = Value
        End Set
    End Property

    Public Property IdAlmacen As String
        Get
            Return _idAlmacen
        End Get
        Set
            _idAlmacen = Value
        End Set
    End Property

    Public Property Estado As Boolean
        Get
            Return _estado
        End Get
        Set
            _estado = Value
        End Set
    End Property

    Public Property Correlativo As Integer
        Get
            Return _correlativo
        End Get
        Set
            _correlativo = Value
        End Set
    End Property

    Public Property NombreCaja As String
        Get
            Return _nombreCaja
        End Get
        Set
            _nombreCaja = Value
        End Set
    End Property

    Public Property Habilitar As Boolean
        Get
            Return _habilitar
        End Get
        Set
            _habilitar = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idCaja As String, ByVal idUsuario As String, ByVal idAlmacen As String, ByVal estado As Boolean, ByVal correlativo As Integer, ByVal nombreCaja As String, ByVal habilitar As Boolean)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NUsuarioCajaGeneral)
        Dim parametros() As Object = {"@idCaja", "@idUsuario", "@idAlmacen", "@estado", "@correlativo", "@nombreCaja", "@habilitar"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.IdCaja, d.IdUsuario, d.IdAlmacen, d.Estado, d.Correlativo, d.NombreCaja, d.Habilitar}
        sql.EjecutarProcedure("Str_Tbl_Usuario_Caja_General_I", parametros, valores, tipoParametro, 7)
    End Sub
    Public Sub Actualizar(d As NUsuarioCajaGeneral)
        Dim parametros() As Object = {"@idCaja", "@idUsuario", "@idAlmacen", "@estado", "@correlativo", "@nombreCaja", "@habilitar"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.IdCaja, d.IdUsuario, d.IdAlmacen, d.Estado, d.Correlativo, d.NombreCaja, d.Habilitar}
        sql.EjecutarProcedure("Str_Tbl_Usuario_Caja_General_U", parametros, valores, tipoParametro, 7)
    End Sub
    Public Sub Eliminar(d As NUsuarioCajaGeneral)
        Dim parametros() As Object = {"@idCaja"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.IdCaja}
        sql.EjecutarProcedure("Str_Tbl_Usuario_Caja_General_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idCaja"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Usuario_Caja_General_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NUsuarioCajaGeneral) As DataTable
        Dim parametros() As Object = {"@idCaja"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.IdCaja}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Usuario_Caja_General_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NUsuarioCajaGeneral) As NUsuarioCajaGeneral
        Dim parametros() As Object = {"@idCaja"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.IdCaja}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Usuario_Caja_General_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdCaja = IIf(dt.Rows(0).Item("idCaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCaja"))
            d.IdUsuario = IIf(dt.Rows(0).Item("idUsuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idUsuario"))
            d.IdAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.Estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.Correlativo = IIf(dt.Rows(0).Item("correlativo") Is DBNull.Value, Nothing, dt.Rows(0).Item("correlativo"))
            d.NombreCaja = IIf(dt.Rows(0).Item("nombreCaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreCaja"))
            d.Habilitar = IIf(dt.Rows(0).Item("habilitar") Is DBNull.Value, Nothing, dt.Rows(0).Item("habilitar"))
        Else

            d.IdAlmacen = Nothing
            d.Estado = Nothing
            d.Correlativo = Nothing
            d.NombreCaja = Nothing
            d.Habilitar = Nothing
        End If
        Return d
    End Function

    ''' <summary>
    ''' Obtiene el registro de usario caja
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    Public Function RegistroCaja(d As NUsuarioCajaGeneral) As NUsuarioCajaGeneral
        Dim parametros() As Object = {"@IdUsuario"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.IdUsuario}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Usuario_Caja_General_M", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdCaja = IIf(dt.Rows(0).Item("idCaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCaja"))
            d.IdUsuario = IIf(dt.Rows(0).Item("idUsuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idUsuario"))
            d.IdAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.Estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.Correlativo = IIf(dt.Rows(0).Item("correlativo") Is DBNull.Value, Nothing, dt.Rows(0).Item("correlativo"))
            d.NombreCaja = IIf(dt.Rows(0).Item("nombreCaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreCaja"))
            d.Habilitar = IIf(dt.Rows(0).Item("habilitar") Is DBNull.Value, Nothing, dt.Rows(0).Item("habilitar"))
        Else
            d.IdAlmacen = Nothing
            d.Estado = Nothing
            d.Correlativo = Nothing
            d.NombreCaja = Nothing
            d.Habilitar = Nothing
        End If
        Return d
    End Function
    ''' <summary>
    ''' Retorna los registros de ventas al contado y cobranzas para una sucursal y un cajero
    ''' </summary>
    ''' <returns></returns>
    Public Function DiarioCaja(IdalmI As String, IdAlmF As String, IdCajaI As String, IdCajaF As String, fechaI As DateTime, fechaF As DateTime) As DataTable
        Dim dt As New DataTable
        Dim parametros() As Object = {"@IdAlmacenI", "@IdAlmacenF", "@IdCajaI", "@IdCajaF", "@FechaMovimientoI", "@FechaMovimientoF"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime}
        Dim valores() As Object = {IdalmI, IdAlmF, IdCajaI, IdCajaF, fechaI, fechaF}
        'dt = sql.ProcedureSQL("Str_ListaPagos", parametros, valores, tipoParametro, 6).Tables(0)
        dt = sql.ProcedureSQL("Str_LiquidacionCaja", parametros, valores, tipoParametro, 6).Tables(0)
        Return dt
    End Function

    Public Function DiarioCajaG(idalmacen As String, fechai As DateTime, fechaf As DateTime, idcaja As String) As DataTable

        Dim cadena As String = " SELECT '1-VENTAS' as Grupo, c.IdAlmacen, c.IdTipoDocumento, c.Serie, c.NumeroDocumento, c.Estado, c.Glosa, c.TipoCambio, c.IdMoneda,sum(c.Pago) as Pago, "
        cadena += " sum(case when ltrim(rtrim(referencia))='1' then c.PagoMN else null end) as PagoMN,sum(case when ltrim(rtrim(referencia))='2' then c.PagoMN else null end) as TarjetaMN, "
        cadena += " sum(case when ltrim(rtrim(referencia))='3' then c.PagoMN else null end) as OtrosMN,sum(c.PagoUS) as PagoUS,'' as IdDetUsuarioCaja,'' as NroDocumentoPago, "
        cadena += " '' as TipoDocumentoPago, c.IdMonedaPago, c.FechaDocumento, d.IdCaja, tc.NombreCaja, CP.IdCliente, CP.NombreCliente, '' as IdCuenta,''as Referencia "
        cadena += " ,vc.Cajas,cp.ChkBancarizar   FROM Tbl_Caja_Venta AS c INNER JOIN Tbl_DetalleUsuarioCaja AS d ON c.IdDetUsuarioCaja = d.IdDetUsuarioCaja INNER JOIN "
        cadena += " Tbl_Usuario_Caja_General AS tc ON d.IdCaja = tc.IdCaja INNER JOIN VTipoOperacion ON d.IdTipoOperacion = VTipoOperacion.IdCodigo INNER JOIN "
        cadena += " Comprobante AS CP ON c.IdAgencia = CP.IdAgencia AND c.IdTipoDocumento = CP.IdTipoDocumento AND c.Serie = CP.Serie AND c.NumeroDocumento = CP.NumeroDocumento AND  "
        cadena += " c.IdAlmacen = CP.IdAlmacen "
        cadena += " inner join vcajas vc on d.Idcaja=vc.IdCaja "
        cadena += " where isnull(cp.estado,'V')='V' And "
        cadena += " (c.FechaDocumento between '" & fechai & "' and '" & fechaf & "') and d.IdCaja in(" & idcaja & ")"
        cadena += " group by c.IdAlmacen, c.IdTipoDocumento, c.Serie, c.NumeroDocumento, c.Estado, c.Glosa, c.TipoCambio, c.IdMoneda, "
        cadena += " c.IdMonedaPago, c.FechaDocumento, d.IdCaja, tc.NombreCaja, CP.IdCliente, CP.NombreCliente,d.idcaja,vc.Cajas,cp.ChkBancarizar    "
        cadena += " union all "
        cadena += " SELECT '2-COBRANZA' as Grupo,dtc.IdAlmacen, dtc.TipoDocumento AS IdTipoDocumento, LEFT(dtc.NroDocumento, 4) AS Serie, RIGHT(dtc.NroDocumento, 7) AS NumeroDocumento, dtc.Estado,  'COBRANZA Pl.Nro°:'+dtc.NroLiq  AS Glosa, "
        cadena += " dtc.TipoCambio, dtc.IdMonedaO AS IdMoneda, dtc.ImporteCobrado AS Pago,case when ltrim(rtrim(referencia))='1' then dtc.ImpCobMN else null end as PagoMN, "
        cadena += " case when ltrim(rtrim(referencia))='2' then dtc.ImpCobMN else null end as TarjetaMN,case when ltrim(rtrim(referencia))='3' then dtc.ImpCobMN else null end as OtrosMN,"
        cadena += " dtc.ImpCobUS AS PagoUS, dtc.IdDetUsuarioCaja,  vfp.Grupo+'-'+dtc.NroDocRef AS NroDocumentoPago, dtc.FormaPago AS TipoDocumentoPago, dtc.IdMoneda AS IdMonedaPago, dtc.FechaMovimiento AS FechaDocumento, "
        cadena += " dtc.IdCaja, '' AS NombreCaja, dtc.IdCliente, cl.Nombre AS NombreCliente,dtuc.IdCuenta AS IdCuenta, VTipoOperacion.Referencia "
        cadena += " ,vc.Cajas,0 AS ChkBancarizar FROM Tbl_DetalleCobranza AS dtc INNER JOIN Cliente AS cl ON dtc.IdCliente = cl.IdCliente INNER JOIN Tbl_DetalleUsuarioCaja AS dtuc ON dtc.IdDetUsuarioCaja = dtuc.IdDetUsuarioCaja INNER JOIN "
        cadena += " VTipoOperacion ON dtuc.IdTipoOperacion = VTipoOperacion.IdCodigo "
        cadena += " inner join vcajas vc on dtc.IdCaja=vc.Idcaja "
        cadena += " left join vformapago vfp on dtc.FormaPago=vfp.IdFormaPago "
        cadena += " where (fechamovimiento between '" & fechai & "' and '" & fechaf & "') and (dtuc.IdCaja in(" & idcaja & "))"
        cadena += " union all "
        cadena += " select case when caja.IdTipoMovimiento='1' Then '3-INGRESOS' else '4-EGRESOS' end  AS Grupo,'" & idalmacen & "' as IdAlmacen,'CJ' as IdTipoDocumento,'' as serie,NumeroTransacion as NumeroDocumento,caja.Estado, "
        cadena += " Glosa,TipoCambio,IdMoneda,case when IdMoneda='MN' then importeMN*signo else IMporteMN*signo/(case when TipoCambio=0 then 1 else tipocambio end) end as Pago, "
        cadena += " ImporteMN*signo as PagoMN,0.00 as TarjetaMN,0.00 as OtrosMN,importeUS*signo as PagosUS,0 as IdDetUsuarioCaja, "
        cadena += "  IdTipoDocumentoRef +'-'+NumeroDocumentoRef as NroDocumentoPago,IdTipoMovimiento+'-'+IdMovimiento as TipoDoCumentoPago, "
        cadena += " IdMoneda as IdMonedaPago,FechaMovimiento as FechaDocumento,caja.IdCaja,'' as NombreCaja "
        cadena += " ,IdTipoAnexo+'-'+IdAnexo as IdCliente,isnull(rtrim(pv.Nombre),rtrim(beneficiario))+case when IdTipoMovimiento='2' then '' else CASE WHEN IdTipoMovimiento='3' then  '-->' else '<---'end end +isnull(vco.cajas,'') as NombreCliente,''as IdCuenta,'1' as Referencia,vc.Cajas,0 AS ChkBancarizar from caja "
        cadena += " inner join vcajas vc on  caja.IdCaja=vc.IdCaja left join vcajas vco on caja.idcajaorigen=vco.idcaja"
        cadena += " left join Proveedor pv on LTRIM(rtrim(IdTipoAnexo))+LTRIM(RTRIM(IdAnexo))=LTRIM(rtrim(pv.TipoAnexo))+LTRIM(RTRIM(pv.IdProveedor)) "
        cadena += " where (caja.IdCaja in(" & idcaja & ")) and (fechamovimiento between '" & fechai & "' and '" & fechaf & "')"

        Dim dt As DataTable = sql.EjecutarConsulta("c", cadena).Tables(0)
        Return dt
    End Function
    Public Function DiarioCajaR(idalmacen As String, fechai As DateTime, fechaf As DateTime, idcaja As String) As DataTable

        Dim cadena As String = " SELECT '1-VENTAS' as Grupo,c.IdMoneda,sum(c.Pago) as Pago,sum(case when ltrim(rtrim(referencia))='1' then c.PagoMN else null end) as PagoMN, "
        cadena += " sum(c.PagoUS) as PagoUS,c.IdMonedaPago, c.FechaDocumento, d.IdCaja, tc.NombreCaja,vc.Cajas FROM Tbl_Caja_Venta AS c INNER JOIN Tbl_DetalleUsuarioCaja AS d ON "
        cadena += " c.IdDetUsuarioCaja = d.IdDetUsuarioCaja INNER JOIN  Tbl_Usuario_Caja_General AS tc ON d.IdCaja = tc.IdCaja INNER JOIN VTipoOperacion ON d.IdTipoOperacion = "
        cadena += " VTipoOperacion.IdCodigo INNER JOIN  Comprobante AS CP ON c.IdAgencia = CP.IdAgencia AND c.IdTipoDocumento = CP.IdTipoDocumento AND c.Serie = CP.Serie AND c.NumeroDocumento = CP.NumeroDocumento AND   c.IdAlmacen = CP.IdAlmacen  "
        cadena += " inner join vcajas vc on d.Idcaja=vc.IdCaja  where isnull(cp.estado,'V')='V' And  "
        cadena += " (c.FechaDocumento between '" & fechai & "' and '" & fechaf & "') and d.IdCaja in(" & idcaja & ")"
        cadena += " group by c.IdMoneda,  c.IdMonedaPago, c.FechaDocumento, d.IdCaja, tc.NombreCaja,vc.Cajas   "
        cadena += " union all "
        cadena += " SELECT '2-COBRANZA' as Grupo,dtc.IdMonedaO AS IdMoneda, sum(dtc.ImporteCobrado) AS Pago,sum(case when ltrim(rtrim(referencia))='1' then dtc.ImpCobMN else null end) as PagoMN,  "
        cadena += " sum(dtc.ImpCobUS) AS PagoUS,dtc.IdMoneda AS IdMonedaPago, dtc.FechaMovimiento AS FechaDocumento,  dtc.IdCaja, '' AS NombreCaja,vc.Cajas FROM Tbl_DetalleCobranza AS dtc "
        cadena += " INNER JOIN Cliente AS cl ON dtc.IdCliente = cl.IdCliente INNER JOIN Tbl_DetalleUsuarioCaja AS dtuc ON dtc.IdDetUsuarioCaja = dtuc.IdDetUsuarioCaja INNER JOIN  VTipoOperacion ON dtuc.IdTipoOperacion = VTipoOperacion.IdCodigo  "
        cadena += " inner join vcajas vc on dtc.IdCaja=vc.Idcaja  left join vformapago vfp on dtc.FormaPago=vfp.IdFormaPago  "
        cadena += " where (fechamovimiento between '" & fechai & "' and '" & fechaf & "') and (dtuc.IdCaja in(" & idcaja & ")) "
        cadena += " and ltrim(rtrim(referencia))='1' group by IdMonedaO, dtc.IdMoneda,FechaMovimiento, dtc.IdCaja,vc.Cajas "
        cadena += " union all "
        cadena += " select case when caja.IdTipoMovimiento='1' Then '3-INGRESOS' else '4-EGRESOS' end  AS Grupo,'' as IdMoneda,sum(case when IdMoneda='MN' then importeMN*signo else IMporteMN*signo/(case when TipoCambio=0 then 1 else tipocambio end) end) as Pago,  "
        cadena += " sum(ImporteMN*signo) as PagoMN,sum(importeUS*signo) as PagosUS,IdMoneda as IdMonedaPago,FechaMovimiento as FechaDocumento,caja.IdCaja,'' as  "
        cadena += " NombreCaja,vc.Cajas from caja  inner join vcajas vc on  caja.IdCaja=vc.IdCaja left join vcajas vco on  caja.idcajaorigen=vco.idcaja left join Proveedor pv on LTRIM(rtrim(IdTipoAnexo))+LTRIM(RTRIM(IdAnexo))=LTRIM(rtrim(pv.TipoAnexo))+LTRIM(RTRIM(pv.IdProveedor))"
        cadena += " where (caja.IdCaja in(" & idcaja & ")) and (fechamovimiento between '" & fechai & "' and '" & fechaf & "')"
        cadena += " group by IdMoneda,IdMoneda,FechaMovimiento, caja.IdCaja,vc.Cajas,caja.IdTipoMovimiento "
        Dim dt As DataTable = sql.EjecutarConsulta("c", cadena).Tables(0)
        Return dt
    End Function
#End Region


End Class
