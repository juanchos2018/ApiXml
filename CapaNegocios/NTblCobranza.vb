Imports CapaDatos
Public Class NTblCobranza
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property tipoliq As String
    Public Property nroliq As String
    Public Property fechacobra As System.DateTime
    Public Property idvendedor As String
    Public Property zonaventa As String
    Public Property importecobrado As Decimal
    Public Property idmoneda As String
    Public Property tipocambio As Decimal
    Public Property situacion As String
    Public Property usuariocrea As String
    Public Property fechacrea As System.DateTime
    Public Property usuariomod As String
    Public Property fechamod As System.DateTime
    Public Property idagencia As String
    Public Property idtipoanexo As String
    Public Property glosa As String
    Public Property idanexo As String
    Public Property idcaja As String
    Public Property idalmacen As String
    Public Property importecobradomn As Decimal
    Public Property importecobradous As Decimal
    Public Property tipodocumento As String
    Public Property nrodocpla As String
    Public Property idsubdiario As String
    Public Property nrocomprobante As String
    Public Property cobservacion As String
    Public Property cidcuentaban As String
    Public Property cnombreban As String
    Public Property IdMediopago As String
    Public Property cidcliente As String
    Public Property cidbanco As String
    Public Property cnrooperacion As String
    Public Property cidarea As String
    Public Property CIdDetUsuarioCaja As String


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal tipoLiq As String, ByVal nroLiq As String, ByVal fechaCobra As System.DateTime, ByVal idVendedor As String, ByVal zonaVenta As String, ByVal importeCobrado As Decimal, ByVal idMoneda As String, ByVal tipoCambio As Decimal, ByVal situacion As String, ByVal usuarioCrea As String, ByVal fechaCrea As System.DateTime, ByVal usuarioMod As String, ByVal fechaMod As System.DateTime, ByVal idAgencia As String, ByVal idTipoAnexo As String, ByVal glosa As String, ByVal idAnexo As String, ByVal idCaja As String, ByVal idAlmacen As String, ByVal importeCobradoMN As Decimal, ByVal importeCobradoUS As Decimal, ByVal tipoDocumento As String, ByVal nroDocPla As String)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NTblCobranza)

        Dim parametros() As Object = {"@tipoliq", "@nroliq", "@fechacobra", "@idvendedor", "@zonaventa", "@importecobrado", "@idmoneda", "@tipocambio", "@situacion", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@idagencia", "@idtipoanexo", "@glosa", "@idanexo", "@idcaja", "@idalmacen", "@tipodocumento", "@nrodocpla", "@idsubdiario", "@nrocomprobante", "@cobservacion", "@cidcuentaban", "@cnombreban", "@idmediopago", "@cidcliente", "@cidbanco", "@cnrooperacion", "@ciddetusuariocaja", "@cidarea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.BigInt, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipoliq, d.nroliq, d.fechacobra, d.idvendedor, d.zonaventa, d.importecobrado, d.idmoneda, d.tipocambio, d.situacion, d.usuariocrea, d.fechacrea, d.usuariomod, d.fechamod, d.idagencia, d.idtipoanexo, d.glosa, d.idanexo, d.idcaja, d.idalmacen, d.tipodocumento, d.nrodocpla, d.idsubdiario, d.nrocomprobante, d.cobservacion, d.cidcuentaban, d.cnombreban, d.IdMediopago, d.cidcliente, d.cidbanco, d.cnrooperacion, d.CIdDetUsuarioCaja, d.cidarea}
        sql.EjecutarProcedure("Str_Tbl_Cobranza_I", parametros, valores, tipoParametro, 32)
    End Sub
    Public Sub Actualizar(d As NTblCobranza)
        Dim parametros() As Object = {"@tipoliq", "@nroliq", "@fechacobra", "@idvendedor", "@zonaventa", "@importecobrado", "@idmoneda", "@tipocambio", "@situacion", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@idagencia", "@idtipoanexo", "@glosa", "@idanexo", "@idcaja", "@idalmacen", "@tipodocumento", "@nrodocpla", "@idsubdiario", "@nrocomprobante", "@cobservacion", "@cidcuentaban", "@cnombreban", "@idmediopago", "@cidcliente", "@cidbanco", "@cnrooperacion", "@ciddetusuariocaja", "@cidarea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.BigInt, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipoliq, d.nroliq, d.fechacobra, d.idvendedor, d.zonaventa, d.importecobrado, d.idmoneda, d.tipocambio, d.situacion, d.usuariocrea, d.fechacrea, d.usuariomod, d.fechamod, d.idagencia, d.idtipoanexo, d.glosa, d.idanexo, d.idcaja, d.idalmacen, d.tipodocumento, d.nrodocpla, d.idsubdiario, d.nrocomprobante, d.cobservacion, d.cidcuentaban, d.cnombreban, d.IdMediopago, d.cidcliente, d.cidbanco, d.cnrooperacion, d.CIdDetUsuarioCaja, d.cidarea}
        sql.EjecutarProcedure("Str_Tbl_Cobranza_U", parametros, valores, tipoParametro, 32)
    End Sub
    Public Sub Eliminar(d As NTblCobranza)
        Dim parametros() As Object = {"@tipoLiq", "@nroLiq"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipoliq, d.nroliq}
        sql.EjecutarProcedure("Str_Tbl_Cobranza_D", parametros, valores, tipoParametro, 2)
    End Sub

    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@tipoLiq", "@nroLiq"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Cobranza_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTblCobranza) As NTblCobranza
        Dim parametros() As Object = {"@tipoliq", "@nroliq"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipoliq, d.nroliq}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Cobranza_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.tipoliq = IIf(dt.Rows(0).Item("tipoliq") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoliq"))
            d.nroliq = IIf(dt.Rows(0).Item("nroliq") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroliq"))
            d.fechacobra = IIf(dt.Rows(0).Item("fechacobra") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacobra"))
            d.idvendedor = IIf(dt.Rows(0).Item("idvendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idvendedor"))
            d.zonaventa = IIf(dt.Rows(0).Item("zonaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("zonaventa"))
            d.importecobrado = IIf(dt.Rows(0).Item("importecobrado") Is DBNull.Value, Nothing, dt.Rows(0).Item("importecobrado"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.situacion = IIf(dt.Rows(0).Item("situacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("situacion"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idtipoanexo = IIf(dt.Rows(0).Item("idtipoanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipoanexo"))
            d.glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.idanexo = IIf(dt.Rows(0).Item("idanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idanexo"))
            d.idcaja = IIf(dt.Rows(0).Item("idcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcaja"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.importecobradomn = IIf(dt.Rows(0).Item("importecobradomn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importecobradomn"))
            d.importecobradous = IIf(dt.Rows(0).Item("importecobradous") Is DBNull.Value, Nothing, dt.Rows(0).Item("importecobradous"))
            d.tipodocumento = IIf(dt.Rows(0).Item("tipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocumento"))
            d.nrodocpla = IIf(dt.Rows(0).Item("nrodocpla") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocpla"))
            d.idsubdiario = IIf(dt.Rows(0).Item("idsubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsubdiario"))
            d.nrocomprobante = IIf(dt.Rows(0).Item("nrocomprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrocomprobante"))
            d.cobservacion = IIf(dt.Rows(0).Item("cobservacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("cobservacion"))
            d.cidcuentaban = IIf(dt.Rows(0).Item("cidcuentaban") Is DBNull.Value, Nothing, dt.Rows(0).Item("cidcuentaban"))
            d.cnombreban = IIf(dt.Rows(0).Item("cnombreban") Is DBNull.Value, Nothing, dt.Rows(0).Item("cnombreban"))
            d.IdMediopago = IIf(dt.Rows(0).Item("idmediopago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmediopago"))
            d.cidcliente = IIf(dt.Rows(0).Item("cidcliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("cidcliente"))
            d.cidbanco = IIf(dt.Rows(0).Item("cidbanco") Is DBNull.Value, Nothing, dt.Rows(0).Item("cidbanco"))
            d.cnrooperacion = IIf(dt.Rows(0).Item("cnrooperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("cnrooperacion"))
            d.CIdDetUsuarioCaja = IIf(dt.Rows(0).Item("ciddetusuariocaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("ciddetusuariocaja"))
            d.cidarea = IIf(dt.Rows(0).Item("cidarea") Is DBNull.Value, Nothing, dt.Rows(0).Item("cidarea"))
        Else
            d.tipoliq = Nothing
            d.nroliq = Nothing
            d.fechacobra = Nothing
            d.idvendedor = Nothing
            d.zonaventa = Nothing
            d.importecobrado = Nothing
            d.idmoneda = Nothing
            d.tipocambio = Nothing
            d.situacion = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.usuariomod = Nothing
            d.fechamod = Nothing
            d.idagencia = Nothing
            d.idtipoanexo = Nothing
            d.glosa = Nothing
            d.idanexo = Nothing
            d.idcaja = Nothing
            d.idalmacen = Nothing
            d.importecobradomn = Nothing
            d.importecobradous = Nothing
            d.tipodocumento = Nothing
            d.nrodocpla = Nothing
            d.idsubdiario = Nothing
            d.nrocomprobante = Nothing
            d.cobservacion = Nothing
            d.cidcuentaban = Nothing
            d.cnombreban = Nothing
            d.IdMediopago = Nothing
            d.cidcliente = Nothing
            d.cidbanco = Nothing
            d.cnrooperacion = Nothing
            d.CIdDetUsuarioCaja = Nothing
            d.cidarea = Nothing
        End If
        Return d
    End Function
#End Region


End Class
