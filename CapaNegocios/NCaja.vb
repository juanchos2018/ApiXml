Imports CapaDatos

Public Class NCaja
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idagencia As String
    Public Property idcaja As String
    Public Property idtipomovimiento As String
    Public Property numerotransacion As String
    Public Property idmovimiento As String
    Public Property fechamovimiento As System.DateTime
    Public Property idsubdiario As String
    Public Property tipocomprobante As String
    Public Property comprobante As String
    Public Property tipocambio As Decimal
    Public Property beneficiario As String
    Public Property glosa As String
    Public Property concepto As String
    Public Property idmoneda As String
    Public Property idtipoanexo As String
    Public Property idanexo As String
    Public Property idtipodocumentoref As String
    Public Property numerodocumentoref As String
    Public Property estado As String
    Public Property usuariocrea As String
    Public Property usuariomod As String
    Public Property fechacrea As System.DateTime
    Public Property fechamod As System.DateTime
    Public Property registromovimiento As String
    Public Property fechamovimiento2 As String
    Public Property idalmacen As String
    Public Property importe As Decimal
    Public Property signo As Integer
    Public Property importemn As Decimal
    Public Property importeus As Decimal
    Public Property idcajaorigen As String
    Public Property idtipomovimientoorigen As String
    Public Property numerotransaccionorigen As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region


#Region "Metodos"
    Public Sub Agregar(d As NCaja)

        Dim parametros() As Object = {"@idagencia", "@idcaja", "@idtipomovimiento", "@numerotransacion", "@idmovimiento", "@fechamovimiento", "@idsubdiario", "@tipocomprobante", "@comprobante", "@tipocambio", "@beneficiario", "@glosa", "@concepto", "@idmoneda", "@idtipoanexo", "@idanexo", "@idtipodocumentoref", "@numerodocumentoref", "@estado", "@usuariocrea", "@usuariomod", "@fechacrea", "@fechamod", "@registromovimiento", "@fechamovimiento2", "@idalmacen", "@importe", "@signo", "@importemn", "@importeus", "@idcajaorigen", "@idtipomovimientoorigen", "@numerotransaccionorigen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Int, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idcaja, d.idtipomovimiento, d.numerotransacion, d.idmovimiento, d.fechamovimiento, d.idsubdiario, d.tipocomprobante, d.comprobante, d.tipocambio, d.beneficiario, d.glosa, d.concepto, d.idmoneda, d.idtipoanexo, d.idanexo, d.idtipodocumentoref, d.numerodocumentoref, d.estado, d.usuariocrea, d.usuariomod, d.fechacrea, d.fechamod, d.registromovimiento, d.fechamovimiento2, d.idalmacen, d.importe, d.signo, d.importemn, d.importeus, d.idcajaorigen, d.idtipomovimientoorigen, d.numerotransaccionorigen}
        sql.EjecutarProcedure("Str_Caja_I", parametros, valores, tipoParametro, 33)
    End Sub
    Public Sub Actualizar(d As NCaja)
        Dim parametros() As Object = {"@idagencia", "@idcaja", "@idtipomovimiento", "@numerotransacion", "@idmovimiento", "@fechamovimiento", "@idsubdiario", "@tipocomprobante", "@comprobante", "@tipocambio", "@beneficiario", "@glosa", "@concepto", "@idmoneda", "@idtipoanexo", "@idanexo", "@idtipodocumentoref", "@numerodocumentoref", "@estado", "@usuariocrea", "@usuariomod", "@fechacrea", "@fechamod", "@registromovimiento", "@fechamovimiento2", "@idalmacen", "@importe", "@signo", "@importemn", "@importeus", "@idcajaorigen", "@idtipomovimientoorigen", "@numerotransaccionorigen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Int, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idcaja, d.idtipomovimiento, d.numerotransacion, d.idmovimiento, d.fechamovimiento, d.idsubdiario, d.tipocomprobante, d.comprobante, d.tipocambio, d.beneficiario, d.glosa, d.concepto, d.idmoneda, d.idtipoanexo, d.idanexo, d.idtipodocumentoref, d.numerodocumentoref, d.estado, d.usuariocrea, d.usuariomod, d.fechacrea, d.fechamod, d.registromovimiento, d.fechamovimiento2, d.idalmacen, d.importe, d.signo, d.importemn, d.importeus, d.idcajaorigen, d.idtipomovimientoorigen, d.numerotransaccionorigen}
        sql.EjecutarProcedure("Str_Caja_U", parametros, valores, tipoParametro, 33)
    End Sub
    Public Sub Eliminar(d As NCaja)
        Dim parametros() As Object = {"@idcaja", "@idtipomovimiento", "@numerotransacion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idcaja, d.idtipomovimiento, d.numerotransacion}
        sql.EjecutarProcedure("Str_Caja_D", parametros, valores, tipoParametro, 3)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idcaja", "@idtipomovimiento", "@numerotransacion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Caja_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NCaja) As DataTable
        Dim parametros() As Object = {"@idcaja", "@idtipomovimiento", "@numerotransacion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idcaja, d.idtipomovimiento, d.numerotransacion}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Caja_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NCaja) As NCaja
        Dim parametros() As Object = {"@idcaja", "@idtipomovimiento", "@numerotransacion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idcaja, d.idtipomovimiento, d.numerotransacion}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Caja_S", parametros, valores, tipoParametro, 3).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idcaja = IIf(dt.Rows(0).Item("idcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcaja"))
            d.idtipomovimiento = IIf(dt.Rows(0).Item("idtipomovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipomovimiento"))
            d.numerotransacion = IIf(dt.Rows(0).Item("numerotransacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerotransacion"))
            d.idmovimiento = IIf(dt.Rows(0).Item("idmovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmovimiento"))
            d.fechamovimiento = IIf(dt.Rows(0).Item("fechamovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamovimiento"))
            d.idsubdiario = IIf(dt.Rows(0).Item("idsubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsubdiario"))
            d.tipocomprobante = IIf(dt.Rows(0).Item("tipocomprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocomprobante"))
            d.comprobante = IIf(dt.Rows(0).Item("comprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("comprobante"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.beneficiario = IIf(dt.Rows(0).Item("beneficiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("beneficiario"))
            d.glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.concepto = IIf(dt.Rows(0).Item("concepto") Is DBNull.Value, Nothing, dt.Rows(0).Item("concepto"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.idtipoanexo = IIf(dt.Rows(0).Item("idtipoanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipoanexo"))
            d.idanexo = IIf(dt.Rows(0).Item("idanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idanexo"))
            d.idtipodocumentoref = IIf(dt.Rows(0).Item("idtipodocumentoref") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumentoref"))
            d.numerodocumentoref = IIf(dt.Rows(0).Item("numerodocumentoref") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumentoref"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.registromovimiento = IIf(dt.Rows(0).Item("registromovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("registromovimiento"))
            d.fechamovimiento2 = IIf(dt.Rows(0).Item("fechamovimiento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamovimiento2"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))
            d.signo = IIf(dt.Rows(0).Item("signo") Is DBNull.Value, Nothing, dt.Rows(0).Item("signo"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.idcajaorigen = IIf(dt.Rows(0).Item("idcajaorigen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcajaorigen"))
            d.idtipomovimientoorigen = IIf(dt.Rows(0).Item("idtipomovimientoorigen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipomovimientoorigen"))
            d.numerotransaccionorigen = IIf(dt.Rows(0).Item("numerotransaccionorigen") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerotransaccionorigen"))
        Else
            d.idagencia = Nothing
            d.idcaja = Nothing
            d.idtipomovimiento = Nothing
            d.numerotransacion = Nothing
            d.idmovimiento = Nothing
            d.fechamovimiento = Nothing
            d.idsubdiario = Nothing
            d.tipocomprobante = Nothing
            d.comprobante = Nothing
            d.tipocambio = Nothing
            d.beneficiario = Nothing
            d.glosa = Nothing
            d.concepto = Nothing
            d.idmoneda = Nothing
            d.idtipoanexo = Nothing
            d.idanexo = Nothing
            d.idtipodocumentoref = Nothing
            d.numerodocumentoref = Nothing
            d.estado = Nothing
            d.usuariocrea = Nothing
            d.usuariomod = Nothing
            d.fechacrea = Nothing
            d.fechamod = Nothing
            d.registromovimiento = Nothing
            d.fechamovimiento2 = Nothing
            d.idalmacen = Nothing
            d.importe = Nothing
            d.signo = Nothing
            d.importemn = Nothing
            d.importeus = Nothing
            d.idcajaorigen = Nothing
            d.idtipomovimientoorigen = Nothing
            d.numerotransaccionorigen = Nothing
        End If
        Return d
    End Function
    Public Function AsientoEgresos(nroTransaccion As String) As DataTable
        Dim ca As String = " select c.IdCaja, IdTipoMovimiento, NumeroTransacion, c.IdMovimiento, FechaMovimiento, TipoCambio, Beneficiario, Concepto, IdMoneda,  "
        ca += " ImporteMN, ImporteUS, IdTipoAnexo, c.IdAnexo, IdTipoDocumentoRef, NumeroDocumentoRef, Estado, c.UsuarioCrea, c.RegistroMovimiento, FechaMovimiento2,  "
        ca += " IdAlmacen, Importe, Signo,m.IdCuenta as IdCuentaGasto,IdCentroCosto,c.Glosa,vc.idcuenta as IdCuentaCaja,c.IdSubdiario,Comprobante,n.idsubdiario as IdSubdiarioCaja from caja c "
        ca += " inner join movimientocaja m on c.idtipomovimiento=m.tipomovimiento and c.IdMovimiento=m.idmovimiento inner join  "
        ca += " (select IdCaja,ltrim(rtrim(idCuenta))as IdCuenta from vcajas) as vc on c.IdCaja=vc.IdCaja inner join numeracion n "
        ca += " on left(c.NumeroTransacion,4)=n.serie and n.IdTipoDocumento='CJ' "
        ca += " where NumeroTransacion='" & nroTransaccion & "'"
        Return sql.EjecutarConsulta("df", ca).Tables(0)
    End Function
    Public Function liquidacionDiario(fecha As DateTime, Ai As String, Af As String, mon As String) As DataTable
        Dim parametros() As Object = {"@FechaDia", "@AlmacenI", "@AlmacenF", "@IdMoneda"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {fecha, Ai, Af, mon}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Liquidacion_Caja_serie", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function liquidacionDiario(fecha As DateTime, Ai As String, Af As String, mon As String, fechaf As DateTime) As DataTable
        Dim parametros() As Object = {"@FechaDia", "@AlmacenI", "@AlmacenF", "@IdMoneda", "@FechaFinal"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {fecha, Ai, Af, mon, fechaf}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Liquidacion_Caja_serie", parametros, valores, tipoParametro, 5).Tables(0)
        Return dt
    End Function

#End Region
End Class
