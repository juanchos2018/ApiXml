Imports CapaDatos
Public Class NDocMovAlmacen
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property tipoMov As String
    Public Property idMovimiento As String
    Public Property descripcion As String
    Public Property flagProveedor As String
    Public Property flagDocReferencia As String
    Public Property flagSolicitante As String
    Public Property flagCentroCostos As String
    Public Property flagGlosa As String
    Public Property flagAlmacenRef As String
    Public Property flagOrdenTrabajo As String
    Public Property flagTipoAnexo As String
    Public Property flagIdAnexo As String
    Public Property flagOrdenCompra As String
    Public Property tipoStock As String
    Public Property tipoCosteo As String
    Public Property prioridadValorizacion As String
    Public Property usuarioCrea As String
    Public Property usuarioMod As String
    Public Property fechaCrea As System.DateTime
    Public Property fechaMod As System.DateTime
    Public Property tipoAnexo As String
    Public Property tM_CFCONSU As String
    Public Property tM_CREPM01 As String
    Public Property tM_CREPM02 As String
    Public Property tM_CREPM03 As String
    Public Property tM_CREPM04 As String
    Public Property tM_CREPM05 As String
    Public Property tM_CREPM06 As String
    Public Property tM_CREPM07 As String
    Public Property tM_CREPM08 As String
    Public Property flagAgencia As String
    Public Property flagCliente As String
    Public Property tM_CFCTACO As String
    Public Property flagOcultar As Boolean
    Public Property flagTransDirecta As Boolean
    Public Property flagTransBach As Boolean
    Public Property flag_KFisico As Boolean
    Public Property flag_KValorado As Boolean

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NDocMovAlmacen)
        Dim parametros() As Object = {"@tipoMov", "@idMovimiento", "@descripcion", "@flagProveedor", "@flagDocReferencia", "@flagSolicitante", "@flagCentroCostos", "@flagGlosa", "@flagAlmacenRef", "@flagOrdenTrabajo", "@flagTipoAnexo", "@flagIdAnexo", "@flagOrdenCompra", "@tipoStock", "@tipoCosteo", "@prioridadValorizacion", "@usuarioCrea", "@usuarioMod", "@fechaCrea", "@fechaMod", "@tipoAnexo", "@tM_CFCONSU", "@tM_CREPM01", "@tM_CREPM02", "@tM_CREPM03", "@tM_CREPM04", "@tM_CREPM05", "@tM_CREPM06", "@tM_CREPM07", "@tM_CREPM08", "@flagAgencia", "@flagCliente", "@tM_CFCTACO", "@flagOcultar", "@flagTransDirecta", "@flagTransBach", "@flag_KFisico", "@flag_KValorado"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.tipoMov, d.idMovimiento, d.descripcion, d.flagProveedor, d.flagDocReferencia, d.flagSolicitante, d.flagCentroCostos, d.flagGlosa, d.flagAlmacenRef, d.flagOrdenTrabajo, d.flagTipoAnexo, d.flagIdAnexo, d.flagOrdenCompra, d.tipoStock, d.tipoCosteo, d.prioridadValorizacion, d.usuarioCrea, d.usuarioMod, d.fechaCrea, d.fechaMod, d.tipoAnexo, d.tM_CFCONSU, d.tM_CREPM01, d.tM_CREPM02, d.tM_CREPM03, d.tM_CREPM04, d.tM_CREPM05, d.tM_CREPM06, d.tM_CREPM07, d.tM_CREPM08, d.flagAgencia, d.flagCliente, d.tM_CFCTACO, d.flagOcultar, d.flagTransDirecta, d.flagTransBach, d.flag_KFisico, d.flag_KValorado}
        sql.EjecutarProcedure("Str_docmovalmacen_I", parametros, valores, tipoParametro, 38)
    End Sub
    Public Sub Actualizar(d As NDocMovAlmacen)
        Dim parametros() As Object = {"@tipoMov", "@idMovimiento", "@descripcion", "@flagProveedor", "@flagDocReferencia", "@flagSolicitante", "@flagCentroCostos", "@flagGlosa", "@flagAlmacenRef", "@flagOrdenTrabajo", "@flagTipoAnexo", "@flagIdAnexo", "@flagOrdenCompra", "@tipoStock", "@tipoCosteo", "@prioridadValorizacion", "@usuarioCrea", "@usuarioMod", "@fechaCrea", "@fechaMod", "@tipoAnexo", "@tM_CFCONSU", "@tM_CREPM01", "@tM_CREPM02", "@tM_CREPM03", "@tM_CREPM04", "@tM_CREPM05", "@tM_CREPM06", "@tM_CREPM07", "@tM_CREPM08", "@flagAgencia", "@flagCliente", "@tM_CFCTACO", "@flagOcultar", "@flagTransDirecta", "@flagTransBach", "@flag_KFisico", "@flag_KValorado"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.tipoMov, d.idMovimiento, d.descripcion, d.flagProveedor, d.flagDocReferencia, d.flagSolicitante, d.flagCentroCostos, d.flagGlosa, d.flagAlmacenRef, d.flagOrdenTrabajo, d.flagTipoAnexo, d.flagIdAnexo, d.flagOrdenCompra, d.tipoStock, d.tipoCosteo, d.prioridadValorizacion, d.usuarioCrea, d.usuarioMod, d.fechaCrea, d.fechaMod, d.tipoAnexo, d.tM_CFCONSU, d.tM_CREPM01, d.tM_CREPM02, d.tM_CREPM03, d.tM_CREPM04, d.tM_CREPM05, d.tM_CREPM06, d.tM_CREPM07, d.tM_CREPM08, d.flagAgencia, d.flagCliente, d.tM_CFCTACO, d.flagOcultar, d.flagTransDirecta, d.flagTransBach, d.flag_KFisico, d.flag_KValorado}
        sql.EjecutarProcedure("Str_docmovalmacen_U", parametros, valores, tipoParametro, 38)
    End Sub
    Public Sub Eliminar(d As NDocMovAlmacen)
        Dim parametros() As Object = {"@tipoMov", "@idMovimiento"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipoMov, d.idMovimiento}
        sql.EjecutarProcedure("Str_docmovalmacen_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@tipoMov", "@idMovimiento"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_docmovalmacen_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NDocMovAlmacen) As DataTable
        Dim parametros() As Object = {"@tipoMov", "@idMovimiento"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipoMov, d.idMovimiento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_docmovalmacen_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDocMovAlmacen) As NDocMovAlmacen
        Dim parametros() As Object = {"@tipoMov", "@idMovimiento"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipoMov, d.idMovimiento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_docmovalmacen_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.tipoMov = IIf(dt.Rows(0).Item("tipoMov") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoMov"))
            d.idMovimiento = IIf(dt.Rows(0).Item("idMovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMovimiento"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.flagProveedor = IIf(dt.Rows(0).Item("flagProveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagProveedor"))
            d.flagDocReferencia = IIf(dt.Rows(0).Item("flagDocReferencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagDocReferencia"))
            d.flagSolicitante = IIf(dt.Rows(0).Item("flagSolicitante") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagSolicitante"))
            d.flagCentroCostos = IIf(dt.Rows(0).Item("flagCentroCostos") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagCentroCostos"))
            d.flagGlosa = IIf(dt.Rows(0).Item("flagGlosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagGlosa"))
            d.flagAlmacenRef = IIf(dt.Rows(0).Item("flagAlmacenRef") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagAlmacenRef"))
            d.flagOrdenTrabajo = IIf(dt.Rows(0).Item("flagOrdenTrabajo") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagOrdenTrabajo"))
            d.flagTipoAnexo = IIf(dt.Rows(0).Item("flagTipoAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagTipoAnexo"))
            d.flagIdAnexo = IIf(dt.Rows(0).Item("flagIdAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagIdAnexo"))
            d.flagOrdenCompra = IIf(dt.Rows(0).Item("flagOrdenCompra") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagOrdenCompra"))
            d.tipoStock = IIf(dt.Rows(0).Item("tipoStock") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoStock"))
            d.tipoCosteo = IIf(dt.Rows(0).Item("tipoCosteo") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoCosteo"))
            d.prioridadValorizacion = IIf(dt.Rows(0).Item("prioridadValorizacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("prioridadValorizacion"))
            d.usuarioCrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.usuarioMod = IIf(dt.Rows(0).Item("usuarioMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioMod"))
            d.fechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.fechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.tipoAnexo = IIf(dt.Rows(0).Item("tipoAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoAnexo"))
            d.tM_CFCONSU = IIf(dt.Rows(0).Item("tM_CFCONSU") Is DBNull.Value, Nothing, dt.Rows(0).Item("tM_CFCONSU"))
            d.tM_CREPM01 = IIf(dt.Rows(0).Item("tM_CREPM01") Is DBNull.Value, Nothing, dt.Rows(0).Item("tM_CREPM01"))
            d.tM_CREPM02 = IIf(dt.Rows(0).Item("tM_CREPM02") Is DBNull.Value, Nothing, dt.Rows(0).Item("tM_CREPM02"))
            d.tM_CREPM03 = IIf(dt.Rows(0).Item("tM_CREPM03") Is DBNull.Value, Nothing, dt.Rows(0).Item("tM_CREPM03"))
            d.tM_CREPM04 = IIf(dt.Rows(0).Item("tM_CREPM04") Is DBNull.Value, Nothing, dt.Rows(0).Item("tM_CREPM04"))
            d.tM_CREPM05 = IIf(dt.Rows(0).Item("tM_CREPM05") Is DBNull.Value, Nothing, dt.Rows(0).Item("tM_CREPM05"))
            d.tM_CREPM06 = IIf(dt.Rows(0).Item("tM_CREPM06") Is DBNull.Value, Nothing, dt.Rows(0).Item("tM_CREPM06"))
            d.tM_CREPM07 = IIf(dt.Rows(0).Item("tM_CREPM07") Is DBNull.Value, Nothing, dt.Rows(0).Item("tM_CREPM07"))
            d.tM_CREPM08 = IIf(dt.Rows(0).Item("tM_CREPM08") Is DBNull.Value, Nothing, dt.Rows(0).Item("tM_CREPM08"))
            d.flagAgencia = IIf(dt.Rows(0).Item("flagAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagAgencia"))
            d.flagCliente = IIf(dt.Rows(0).Item("flagCliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagCliente"))
            d.tM_CFCTACO = IIf(dt.Rows(0).Item("tM_CFCTACO") Is DBNull.Value, Nothing, dt.Rows(0).Item("tM_CFCTACO"))
            d.flagOcultar = IIf(dt.Rows(0).Item("flagOcultar") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagOcultar"))
            d.flagTransDirecta = IIf(dt.Rows(0).Item("flagTransDirecta") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagTransDirecta"))
            d.flagTransBach = IIf(dt.Rows(0).Item("flagTransBach") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagTransBach"))
            d.flag_KFisico = IIf(dt.Rows(0).Item("flag_KFisico") Is DBNull.Value, Nothing, dt.Rows(0).Item("flag_KFisico"))
            d.flag_KValorado = IIf(dt.Rows(0).Item("flag_KValorado") Is DBNull.Value, Nothing, dt.Rows(0).Item("flag_KValorado"))
        Else
            d.tipoMov = Nothing
            d.idMovimiento = Nothing
            d.descripcion = Nothing
            d.flagProveedor = Nothing
            d.flagDocReferencia = Nothing
            d.flagSolicitante = Nothing
            d.flagCentroCostos = Nothing
            d.flagGlosa = Nothing
            d.flagAlmacenRef = Nothing
            d.flagOrdenTrabajo = Nothing
            d.flagTipoAnexo = Nothing
            d.flagIdAnexo = Nothing
            d.flagOrdenCompra = Nothing
            d.tipoStock = Nothing
            d.tipoCosteo = Nothing
            d.prioridadValorizacion = Nothing
            d.usuarioCrea = Nothing
            d.usuarioMod = Nothing
            d.fechaCrea = Nothing
            d.fechaMod = Nothing
            d.tipoAnexo = Nothing
            d.tM_CFCONSU = Nothing
            d.tM_CREPM01 = Nothing
            d.tM_CREPM02 = Nothing
            d.tM_CREPM03 = Nothing
            d.tM_CREPM04 = Nothing
            d.tM_CREPM05 = Nothing
            d.tM_CREPM06 = Nothing
            d.tM_CREPM07 = Nothing
            d.tM_CREPM08 = Nothing
            d.flagAgencia = Nothing
            d.flagCliente = Nothing
            d.tM_CFCTACO = Nothing
            d.flagOcultar = Nothing
            d.flagTransDirecta = Nothing
            d.flagTransBach = Nothing
            d.flag_KFisico = Nothing
            d.flag_KValorado = Nothing
        End If
        Return d
    End Function
#End Region

End Class
