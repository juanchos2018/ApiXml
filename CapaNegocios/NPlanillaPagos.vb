Imports CapaDatos

Public Class NPlanillaPagos
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idplanilla As Long
    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numeroliquidacion As String
    Public Property idmoneda As String
    Public Property idcaja As String
    Public Property fechaliquidacion As System.DateTime
    Public Property pagototal As Decimal
    Public Property usuariocrea As String
    Public Property fechacrea As System.DateTime
    Public Property tipocambio As Decimal
    Public Property pagototalmn As Decimal
    Public Property pagototalus As Decimal
    Public Property idsubdiario As String
    Public Property nrocomprobante As String
    Public Property Glosa As String
    Public Property observacion As String
    Public Property MedioPago As String
    Public Property nrooperacion As String
    Public Property serieoperacion As String
    Public Property idproveedor As String
    Public Property idfpago As String
    Public Property idbanco As String

    Public Property isplanillamixta As Boolean

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NPlanillaPagos)

        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numeroliquidacion", "@idmoneda", "@idcaja", "@fechaliquidacion", "@pagototal", "@usuariocrea", "@fechacrea", "@tipocambio", "@pagototalmn", "@pagototalus", "@idsubdiario", "@nrocomprobante", "@glosa", "@observacion", "@mediopago", "@nrooperacion", "@serieoperacion", "@idproveedor", "@idfpago", "@idbanco", "@isplanillamixta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numeroliquidacion, d.idmoneda, d.idcaja, d.fechaliquidacion, d.pagototal, d.usuariocrea, d.fechacrea, d.tipocambio, d.pagototalmn, d.pagototalus, d.idsubdiario, d.nrocomprobante, d.Glosa, d.observacion, d.MedioPago, d.nrooperacion, d.serieoperacion, d.idproveedor, d.idfpago, d.idbanco, d.isplanillamixta}
        sql.EjecutarProcedure("Str_PlanillaPagos_I", parametros, valores, tipoParametro, 23)
    End Sub
    Public Sub Actualizar(d As NPlanillaPagos)
        Dim parametros() As Object = {"@idplanilla", "@idtipodocumento", "@serie", "@numeroliquidacion", "@idmoneda", "@idcaja", "@fechaliquidacion", "@pagototal", "@usuariocrea", "@fechacrea", "@tipocambio", "@pagototalmn", "@pagototalus", "@idsubdiario", "@nrocomprobante", "@glosa", "@observacion", "@mediopago", "@nrooperacion", "@serieoperacion", "@idproveedor", "@idfpago", "@idbanco", "@isplanillamixta"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.idplanilla, d.idtipodocumento, d.serie, d.numeroliquidacion, d.idmoneda, d.idcaja, d.fechaliquidacion, d.pagototal, d.usuariocrea, d.fechacrea, d.tipocambio, d.pagototalmn, d.pagototalus, d.idsubdiario, d.nrocomprobante, d.Glosa, d.observacion, d.MedioPago, d.nrooperacion, d.serieoperacion, d.idproveedor, d.idfpago, d.idbanco, d.isplanillamixta}
        sql.EjecutarProcedure("Str_PlanillaPagos_U", parametros, valores, tipoParametro, 24)
    End Sub
    Public Function Agregar(d As NPlanillaPagos, Retornatable As Boolean) As NPlanillaPagos

        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numeroliquidacion", "@idmoneda", "@idcaja", "@fechaliquidacion", "@pagototal", "@usuariocrea", "@fechacrea", "@tipocambio", "@pagototalmn", "@pagototalus", "@idsubdiario", "@nrocomprobante", "@glosa", "@observacion", "@mediopago", "@nrooperacion", "@serieoperacion", "@idproveedor", "@idfpago", "@idbanco", "@isplanillamixta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numeroliquidacion, d.idmoneda, d.idcaja, d.fechaliquidacion, d.pagototal, d.usuariocrea, d.fechacrea, d.tipocambio, d.pagototalmn, d.pagototalus, d.idsubdiario, d.nrocomprobante, d.Glosa, d.observacion, d.MedioPago, d.nrooperacion, d.serieoperacion, d.idproveedor, d.idfpago, d.idbanco, d.isplanillamixta}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_PlanillaPagos_I_S", parametros, valores, tipoParametro, 23).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idplanilla = IIf(dt.Rows(0).Item("idplanilla") Is DBNull.Value, Nothing, dt.Rows(0).Item("idplanilla"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numeroliquidacion = IIf(dt.Rows(0).Item("numeroliquidacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroliquidacion"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.idcaja = IIf(dt.Rows(0).Item("idcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcaja"))
            d.fechaliquidacion = IIf(dt.Rows(0).Item("fechaliquidacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaliquidacion"))
            d.pagototal = IIf(dt.Rows(0).Item("pagototal") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagototal"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.pagototalmn = IIf(dt.Rows(0).Item("pagototalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagototalmn"))
            d.pagototalus = IIf(dt.Rows(0).Item("pagototalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagototalus"))
            d.idsubdiario = IIf(dt.Rows(0).Item("idsubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsubdiario"))
            d.nrocomprobante = IIf(dt.Rows(0).Item("nrocomprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrocomprobante"))
            d.Glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.observacion = IIf(dt.Rows(0).Item("observacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("observacion"))
            d.MedioPago = IIf(dt.Rows(0).Item("mediopago") Is DBNull.Value, Nothing, dt.Rows(0).Item("mediopago"))
            d.nrooperacion = IIf(dt.Rows(0).Item("nrooperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrooperacion"))
            d.serieoperacion = IIf(dt.Rows(0).Item("serieoperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("serieoperacion"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.idfpago = IIf(dt.Rows(0).Item("idfpago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idfpago"))
            d.idbanco = IIf(dt.Rows(0).Item("idbanco") Is DBNull.Value, Nothing, dt.Rows(0).Item("idbanco"))
            d.isplanillamixta = IIf(dt.Rows(0).Item("isplanillamixta") Is DBNull.Value, Nothing, dt.Rows(0).Item("isplanillamixta"))
        Else
            d.idplanilla = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numeroliquidacion = Nothing
            d.idmoneda = Nothing
            d.idcaja = Nothing
            d.fechaliquidacion = Nothing
            d.pagototal = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.tipocambio = Nothing
            d.pagototalmn = Nothing
            d.pagototalus = Nothing
            d.idsubdiario = Nothing
            d.nrocomprobante = Nothing
            d.Glosa = Nothing
            d.observacion = Nothing
            d.MedioPago = Nothing
            d.nrooperacion = Nothing
            d.serieoperacion = Nothing
            d.idproveedor = Nothing
            d.idfpago = Nothing
            d.idbanco = Nothing
            d.isplanillamixta = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NPlanillaPagos, Retornatable As Boolean) As NPlanillaPagos

        Dim parametros() As Object = {"@idplanilla", "@idtipodocumento", "@serie", "@numeroliquidacion", "@idmoneda", "@idcaja", "@fechaliquidacion", "@pagototal", "@usuariocrea", "@fechacrea", "@tipocambio", "@pagototalmn", "@pagototalus", "@idsubdiario", "@nrocomprobante", "@glosa", "@observacion", "@mediopago", "@nrooperacion", "@serieoperacion", "@idproveedor", "@idfpago", "@idbanco", "@isplanillamixta"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.idplanilla, d.idtipodocumento, d.serie, d.numeroliquidacion, d.idmoneda, d.idcaja, d.fechaliquidacion, d.pagototal, d.usuariocrea, d.fechacrea, d.tipocambio, d.pagototalmn, d.pagototalus, d.idsubdiario, d.nrocomprobante, d.Glosa, d.observacion, d.MedioPago, d.nrooperacion, d.serieoperacion, d.idproveedor, d.idfpago, d.idbanco, d.isplanillamixta}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_PlanillaPagos_U_S", parametros, valores, tipoParametro, 24).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idplanilla = IIf(dt.Rows(0).Item("idplanilla") Is DBNull.Value, Nothing, dt.Rows(0).Item("idplanilla"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numeroliquidacion = IIf(dt.Rows(0).Item("numeroliquidacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroliquidacion"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.idcaja = IIf(dt.Rows(0).Item("idcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcaja"))
            d.fechaliquidacion = IIf(dt.Rows(0).Item("fechaliquidacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaliquidacion"))
            d.pagototal = IIf(dt.Rows(0).Item("pagototal") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagototal"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.pagototalmn = IIf(dt.Rows(0).Item("pagototalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagototalmn"))
            d.pagototalus = IIf(dt.Rows(0).Item("pagototalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagototalus"))
            d.idsubdiario = IIf(dt.Rows(0).Item("idsubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsubdiario"))
            d.nrocomprobante = IIf(dt.Rows(0).Item("nrocomprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrocomprobante"))
            d.Glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.observacion = IIf(dt.Rows(0).Item("observacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("observacion"))
            d.MedioPago = IIf(dt.Rows(0).Item("mediopago") Is DBNull.Value, Nothing, dt.Rows(0).Item("mediopago"))
            d.nrooperacion = IIf(dt.Rows(0).Item("nrooperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrooperacion"))
            d.serieoperacion = IIf(dt.Rows(0).Item("serieoperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("serieoperacion"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.idfpago = IIf(dt.Rows(0).Item("idfpago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idfpago"))
            d.idbanco = IIf(dt.Rows(0).Item("idbanco") Is DBNull.Value, Nothing, dt.Rows(0).Item("idbanco"))
            d.isplanillamixta = IIf(dt.Rows(0).Item("isplanillamixta") Is DBNull.Value, Nothing, dt.Rows(0).Item("isplanillamixta"))
        Else
            d.idplanilla = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numeroliquidacion = Nothing
            d.idmoneda = Nothing
            d.idcaja = Nothing
            d.fechaliquidacion = Nothing
            d.pagototal = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.tipocambio = Nothing
            d.pagototalmn = Nothing
            d.pagototalus = Nothing
            d.idsubdiario = Nothing
            d.nrocomprobante = Nothing
            d.Glosa = Nothing
            d.observacion = Nothing
            d.MedioPago = Nothing
            d.nrooperacion = Nothing
            d.serieoperacion = Nothing
            d.idproveedor = Nothing
            d.idfpago = Nothing
            d.idbanco = Nothing
            d.isplanillamixta = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NPlanillaPagos)
        Dim parametros() As Object = {"@idplanilla"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idplanilla}
        sql.EjecutarProcedure("Str_PlanillaPagos_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idplanilla"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_PlanillaPagos_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NPlanillaPagos) As DataTable
        Dim parametros() As Object = {"@idplanilla"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idplanilla}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_PlanillaPagos_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NPlanillaPagos) As NPlanillaPagos
        Dim parametros() As Object = {"@idplanilla"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idplanilla}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_PlanillaPagos_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idplanilla = IIf(dt.Rows(0).Item("idplanilla") Is DBNull.Value, Nothing, dt.Rows(0).Item("idplanilla"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numeroliquidacion = IIf(dt.Rows(0).Item("numeroliquidacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroliquidacion"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.idcaja = IIf(dt.Rows(0).Item("idcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcaja"))
            d.fechaliquidacion = IIf(dt.Rows(0).Item("fechaliquidacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaliquidacion"))
            d.pagototal = IIf(dt.Rows(0).Item("pagototal") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagototal"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.pagototalmn = IIf(dt.Rows(0).Item("pagototalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagototalmn"))
            d.pagototalus = IIf(dt.Rows(0).Item("pagototalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagototalus"))
            d.idsubdiario = IIf(dt.Rows(0).Item("idsubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsubdiario"))
            d.nrocomprobante = IIf(dt.Rows(0).Item("nrocomprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrocomprobante"))
            d.Glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.observacion = IIf(dt.Rows(0).Item("observacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("observacion"))
            d.MedioPago = IIf(dt.Rows(0).Item("mediopago") Is DBNull.Value, Nothing, dt.Rows(0).Item("mediopago"))
            d.nrooperacion = IIf(dt.Rows(0).Item("nrooperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrooperacion"))
            d.serieoperacion = IIf(dt.Rows(0).Item("serieoperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("serieoperacion"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.idfpago = IIf(dt.Rows(0).Item("idfpago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idfpago"))
            d.idbanco = IIf(dt.Rows(0).Item("idbanco") Is DBNull.Value, Nothing, dt.Rows(0).Item("idbanco"))
            d.isplanillamixta = IIf(dt.Rows(0).Item("isplanillamixta") Is DBNull.Value, Nothing, dt.Rows(0).Item("isplanillamixta"))
        Else
            d.idplanilla = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numeroliquidacion = Nothing
            d.idmoneda = Nothing
            d.idcaja = Nothing
            d.fechaliquidacion = Nothing
            d.pagototal = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.tipocambio = Nothing
            d.pagototalmn = Nothing
            d.pagototalus = Nothing
            d.idsubdiario = Nothing
            d.nrocomprobante = Nothing
            d.Glosa = Nothing
            d.observacion = Nothing
            d.MedioPago = Nothing
            d.nrooperacion = Nothing
            d.serieoperacion = Nothing
            d.idproveedor = Nothing
            d.idfpago = Nothing
            d.idbanco = Nothing
            d.isplanillamixta = Nothing
        End If
        Return d
    End Function

    Public Function Get_Id(idcaja As String, serie As String, nro As String) As DataTable
        Dim ca As String = " select IdPlanilla from PlanillaPagos where IdCaja='" & idcaja & "' and serie='" & serie & "' and NumeroLiquidacion='" & nro & "'"
        Dim dt As New DataTable
        dt = sql.EjecutarConsulta("d", ca).Tables(0)
        Return dt
    End Function
    Public Function lista_pla() As DataTable
        Dim ca As String = " select IdPlanilla,IdTipoDocumento,Serie,NumeroLiquidacion,IdMoneda,PagoTotal,IdSubdiario,NroComprobante,TipoCambio from planillapagos "
        Return sql.EjecutarConsulta("c", ca).Tables(0)
    End Function

#End Region

End Class
