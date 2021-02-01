Imports CapaDatos
Public Class NAduana_DetalleCobranza
    Property sql As New ClsConexion
#Region "Propiedades"

    Public Property correlativo As Long
    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numerodocumento As String
    Public Property item As String
    Public Property tipomovimiento As String
    Public Property idmovimiento As String
    Public Property fechadocumento As System.DateTime
    Public Property idtipodocumentoref As String
    Public Property numerodocumentoref As String
    Public Property idproveedor As String
    Public Property concepto As String
    Public Property tipocambio As Decimal
    Public Property idmoneda As String
    Public Property importe As Decimal
    Public Property importemn As Decimal
    Public Property importeus As Decimal
    Public Property fechacrea As System.DateTime
    Public Property usuariocrea As String
    Public Property idcaja As String
    Public Property fechapago As DateTime?
    Public Property idtipodocumentoref1 As String
    Public Property numerodocumentoref1 As String
    Public Property idcuenta As String
    Public Property idsubdiario As String
    Public Property nrocomprobante As String
    Public Property idarea As String
    Public Property IdCuentaCaja As String
    Public Property EsVenta As Boolean
#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NAduana_DetalleCobranza)
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item", "@tipomovimiento", "@idmovimiento", "@fechadocumento", "@idtipodocumentoref", "@numerodocumentoref", "@idcaja", "@fechapago", "@idtipodocumentoref1", "@numerodocumentoref1", "@idproveedor", "@concepto", "@tipocambio", "@idmoneda", "@importe", "@fechacrea", "@usuariocrea", "@idcuenta", "@idsubdiario", "@nrocomprobante", "@idarea", "@EsVenta", "@IdCuentaCaja"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.tipomovimiento, d.idmovimiento, d.fechadocumento, d.idtipodocumentoref, d.numerodocumentoref, d.idcaja, d.fechapago, d.idtipodocumentoref1, d.numerodocumentoref1, d.idproveedor, d.concepto, d.tipocambio, d.idmoneda, d.importe, d.fechacrea, d.usuariocrea, d.idcuenta, d.idsubdiario, d.nrocomprobante, d.idarea, d.EsVenta, d.IdCuentaCaja}
        sql.EjecutarProcedure("Str_Aduana_DetalleCobranza_I", parametros, valores, tipoParametro, 26)
    End Sub
    Public Sub Actualizar(d As NAduana_DetalleCobranza)
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item", "@tipomovimiento", "@idmovimiento", "@fechadocumento", "@idtipodocumentoref", "@numerodocumentoref", "@idcaja", "@fechapago", "@idtipodocumentoref1", "@numerodocumentoref1", "@idproveedor", "@concepto", "@tipocambio", "@idmoneda", "@importe", "@fechacrea", "@usuariocrea", "@idcuenta", "@idsubdiario", "@nrocomprobante", "@idarea", "@IdCuentaCaja", "@EsVenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.tipomovimiento, d.idmovimiento, d.fechadocumento, d.idtipodocumentoref, d.numerodocumentoref, d.idcaja, d.fechapago, d.idtipodocumentoref1, d.numerodocumentoref1, d.idproveedor, d.concepto, d.tipocambio, d.idmoneda, d.importe, d.fechacrea, d.usuariocrea, d.idcuenta, d.idsubdiario, d.nrocomprobante, d.idarea, d.IdCuentaCaja, d.EsVenta}
        sql.EjecutarProcedure("Str_Aduana_DetalleCobranza_U", parametros, valores, tipoParametro, 26)
    End Sub

    Public Function Agregar(d As NAduana_DetalleCobranza, Retornatable As Boolean) As NAduana_DetalleCobranza

        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item", "@tipomovimiento", "@idmovimiento", "@fechadocumento", "@idtipodocumentoref", "@numerodocumentoref", "@idcaja", "@fechapago", "@idtipodocumentoref1", "@numerodocumentoref1", "@idproveedor", "@concepto", "@tipocambio", "@idmoneda", "@importe", "@fechacrea", "@usuariocrea", "@idcuenta", "@idsubdiario", "@nrocomprobante", "@idarea", "@EsVenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.tipomovimiento, d.idmovimiento, d.fechadocumento, d.idtipodocumentoref, d.numerodocumentoref, d.idcaja, d.fechapago, d.idtipodocumentoref1, d.numerodocumentoref1, d.idproveedor, d.concepto, d.tipocambio, d.idmoneda, d.importe, d.fechacrea, d.usuariocrea, d.idcuenta, d.idsubdiario, d.nrocomprobante, d.idarea, d.EsVenta}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Aduana_DetalleCobranza_I_S", parametros, valores, tipoParametro, 25).Tables(0)
        If dt.Rows.Count > 0 Then
            d.correlativo = IIf(dt.Rows(0).Item("correlativo") Is DBNull.Value, Nothing, dt.Rows(0).Item("correlativo"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.tipomovimiento = IIf(dt.Rows(0).Item("tipomovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipomovimiento"))
            d.idmovimiento = IIf(dt.Rows(0).Item("idmovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmovimiento"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.idtipodocumentoref = IIf(dt.Rows(0).Item("idtipodocumentoref") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumentoref"))
            d.numerodocumentoref = IIf(dt.Rows(0).Item("numerodocumentoref") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumentoref"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.concepto = IIf(dt.Rows(0).Item("concepto") Is DBNull.Value, Nothing, dt.Rows(0).Item("concepto"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.correlativo = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.item = Nothing
            d.tipomovimiento = Nothing
            d.idmovimiento = Nothing
            d.fechadocumento = Nothing
            d.idtipodocumentoref = Nothing
            d.numerodocumentoref = Nothing
            d.idproveedor = Nothing
            d.concepto = Nothing
            d.tipocambio = Nothing
            d.idmoneda = Nothing
            d.importe = Nothing
            d.importemn = Nothing
            d.importeus = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NAduana_DetalleCobranza, Retornatable As Boolean) As NAduana_DetalleCobranza

        Dim parametros() As Object = {"@correlativo", "@idtipodocumento", "@serie", "@numerodocumento", "@item", "@tipomovimiento", "@idmovimiento", "@fechadocumento", "@idtipodocumentoref", "@numerodocumentoref", "@idproveedor", "@concepto", "@tipocambio", "@idmoneda", "@importe", "@importemn", "@importeus", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.Char}
        Dim valores() As Object = {d.usuariocrea = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Aduana_DetalleCobranza_U_S", parametros, valores, tipoParametro, 57).Tables(0)
        If dt.Rows.Count > 0 Then
            d.correlativo = IIf(dt.Rows(0).Item("correlativo") Is DBNull.Value, Nothing, dt.Rows(0).Item("correlativo"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.tipomovimiento = IIf(dt.Rows(0).Item("tipomovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipomovimiento"))
            d.idmovimiento = IIf(dt.Rows(0).Item("idmovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmovimiento"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.idtipodocumentoref = IIf(dt.Rows(0).Item("idtipodocumentoref") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumentoref"))
            d.numerodocumentoref = IIf(dt.Rows(0).Item("numerodocumentoref") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumentoref"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.concepto = IIf(dt.Rows(0).Item("concepto") Is DBNull.Value, Nothing, dt.Rows(0).Item("concepto"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.correlativo = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.item = Nothing
            d.tipomovimiento = Nothing
            d.idmovimiento = Nothing
            d.fechadocumento = Nothing
            d.idtipodocumentoref = Nothing
            d.numerodocumentoref = Nothing
            d.idproveedor = Nothing
            d.concepto = Nothing
            d.tipocambio = Nothing
            d.idmoneda = Nothing
            d.importe = Nothing
            d.importemn = Nothing
            d.importeus = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NAduana_DetalleCobranza)
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item}
        sql.EjecutarProcedure("Str_Aduana_DetalleCobranza_D", parametros, valores, tipoParametro, 4)
    End Sub
    Public Function Existe_Aduana_DetalleCobranza(d As NAduana_DetalleCobranza) As Boolean
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_Aduana_DetalleCobranza", parametros, valores, tipoParametro, 4)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Aduana_DetalleCobranza_S", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NAduana_DetalleCobranza) As DataTable
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Aduana_DetalleCobranza_S", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NAduana_DetalleCobranza) As NAduana_DetalleCobranza
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Aduana_DetalleCobranza_S", parametros, valores, tipoParametro, 4).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.tipomovimiento = IIf(dt.Rows(0).Item("tipomovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipomovimiento"))
            d.idmovimiento = IIf(dt.Rows(0).Item("idmovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmovimiento"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.idtipodocumentoref = IIf(dt.Rows(0).Item("idtipodocumentoref") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumentoref"))
            d.numerodocumentoref = IIf(dt.Rows(0).Item("numerodocumentoref") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumentoref"))
            d.idcaja = IIf(dt.Rows(0).Item("idcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcaja"))
            d.fechapago = IIf(dt.Rows(0).Item("fechapago") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechapago"))
            d.idtipodocumentoref1 = IIf(dt.Rows(0).Item("idtipodocumentoref1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumentoref1"))
            d.numerodocumentoref1 = IIf(dt.Rows(0).Item("numerodocumentoref1") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumentoref1"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.concepto = IIf(dt.Rows(0).Item("concepto") Is DBNull.Value, Nothing, dt.Rows(0).Item("concepto"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.idcuenta = IIf(dt.Rows(0).Item("idcuenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcuenta"))
            d.idsubdiario = IIf(dt.Rows(0).Item("idsubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsubdiario"))
            d.nrocomprobante = IIf(dt.Rows(0).Item("nrocomprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrocomprobante"))
            d.idarea = IIf(dt.Rows(0).Item("idarea") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarea"))
            d.IdCuentaCaja = IIf(dt.Rows(0).Item("IdCuentaCaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("IdCuentaCaja"))
            d.EsVenta = IIf(dt.Rows(0).Item("EsVenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("EsVenta"))
        Else
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.item = Nothing
            d.tipomovimiento = Nothing
            d.idmovimiento = Nothing
            d.fechadocumento = Nothing
            d.idtipodocumentoref = Nothing
            d.numerodocumentoref = Nothing
            d.idcaja = Nothing
            d.fechapago = Nothing
            d.idtipodocumentoref1 = Nothing
            d.numerodocumentoref1 = Nothing
            d.idproveedor = Nothing
            d.concepto = Nothing
            d.tipocambio = Nothing
            d.idmoneda = Nothing
            d.importe = Nothing
            d.importemn = Nothing
            d.importeus = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
            d.idcuenta = Nothing
            d.idsubdiario = Nothing
            d.nrocomprobante = Nothing
            d.idarea = Nothing
            d.IdCuentaCaja = Nothing
            d.EsVenta = Nothing
        End If
        Return d
    End Function
#End Region
    Public Function Lista_Cobranza(fechaI As DateTime, fechaF As DateTime) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime}
        Dim valores() As Object = {fechaI, fechaF}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ListaCobranzas", parametros, valores, tipoParametro, 2, True)
        Return dt
    End Function
    Public Function Lista_Asiento(idproveedor As String, idtipodocumento As String, numerodocumento As String, item As String) As DataTable
        Dim parametros() As Object = {"@IdProveedor", "@IdTipoDocumento", "@NumeroDocumento", "@Item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {idproveedor, idtipodocumento, numerodocumento, item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Asiento_cobranza", parametros, valores, tipoParametro, 4, True)
        Return dt
    End Function
    Public Function Lista_Asiento_01(serieliq As String, nroliq As String, idproveedor As String, idtipodocumento As String, numerodocumento As String, item As String) As DataTable
        Dim parametros() As Object = {"@SerieLiq", "@NroLiq", "@IdProveedor", "@IdTipoDocumento", "@NumeroDocumento", "@Item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {serieliq, nroliq, idproveedor, idtipodocumento, numerodocumento, item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Asiento_cobranza_01", parametros, valores, tipoParametro, 6, True)
        Return dt
    End Function
End Class
