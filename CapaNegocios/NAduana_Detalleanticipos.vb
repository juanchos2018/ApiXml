Imports CapaDatos
Public Class NAduana_Detalleanticipos
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numerodocumento As String
    Public Property item As String
    Public Property tipomovimiento As String
    Public Property idmovimiento As String
    Public Property fechadocumento As System.DateTime
    Public Property idtipodocumentoref As String
    Public Property numerodocumentoref As String
    Public Property idcliente As String
    Public Property concepto As String
    Public Property tipocambio As Decimal
    Public Property idmoneda As String
    Public Property importe As Decimal
    Public Property importemn As Decimal
    Public Property importeus As Decimal
    Public Property fechacrea As System.DateTime
    Public Property usuariocrea As String
    Public Property IdCuenta As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NAduana_Detalleanticipos)
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item", "@tipomovimiento", "@idmovimiento", "@fechadocumento", "@idtipodocumentoref", "@numerodocumentoref", "@idcliente", "@concepto", "@tipocambio", "@idmoneda", "@importe", "@fechacrea", "@usuariocrea", "IdCuenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.tipomovimiento, d.idmovimiento, d.fechadocumento, d.idtipodocumentoref, d.numerodocumentoref, d.idcliente, d.concepto, d.tipocambio, d.idmoneda, d.importe, d.fechacrea, d.usuariocrea, d.IdCuenta}
        sql.EjecutarProcedure("Str_Aduana_Detalleanticipos_I", parametros, valores, tipoParametro, 17)
    End Sub
    Public Sub Actualizar(d As NAduana_Detalleanticipos)
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item", "@tipomovimiento", "@idmovimiento", "@fechadocumento", "@idtipodocumentoref", "@numerodocumentoref", "@idcliente", "@concepto", "@tipocambio", "@idmoneda", "@importe", "@fechacrea", "@usuariocrea", "IdCuenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.tipomovimiento, d.idmovimiento, d.fechadocumento, d.idtipodocumentoref, d.numerodocumentoref, d.idcliente, d.concepto, d.tipocambio, d.idmoneda, d.importe, d.fechacrea, d.usuariocrea, d.IdCuenta}
        sql.EjecutarProcedure("Str_Aduana_Detalleanticipos_U", parametros, valores, tipoParametro, 17)
    End Sub
    Public Function Agregar(d As NAduana_Detalleanticipos, Retornatable As Boolean) As NAduana_Detalleanticipos
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item", "@tipomovimiento", "@idmovimiento", "@fechadocumento", "@idtipodocumentoref", "@numerodocumentoref", "@idcliente", "@concepto", "@tipocambio", "@idmoneda", "@importe", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.Char}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.tipomovimiento, d.idmovimiento, d.fechadocumento, d.idtipodocumentoref, d.numerodocumentoref, d.idcliente, d.concepto, d.tipocambio, d.idmoneda, d.importe, d.fechacrea, d.usuariocrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Aduana_Detalleanticipos_I_S", parametros, valores, tipoParametro, 16).Tables(0)
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

            d.idcliente = IIf(dt.Rows(0).Item("idcliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcliente"))

            d.concepto = IIf(dt.Rows(0).Item("concepto") Is DBNull.Value, Nothing, dt.Rows(0).Item("concepto"))

            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))

            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))

            d.importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))

            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))

            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))

            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))

            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))

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
            d.idcliente = Nothing
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
    Public Function Actualizar(d As NAduana_Detalleanticipos, Retornatable As Boolean) As NAduana_Detalleanticipos
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item", "@tipomovimiento", "@idmovimiento", "@fechadocumento", "@idtipodocumentoref", "@numerodocumentoref", "@idcliente", "@concepto", "@tipocambio", "@idmoneda", "@importe", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.Char}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.tipomovimiento, d.idmovimiento, d.fechadocumento, d.idtipodocumentoref, d.numerodocumentoref, d.idcliente, d.concepto, d.tipocambio, d.idmoneda, d.importe, d.fechacrea, d.usuariocrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Aduana_Detalleanticipos_U_S", parametros, valores, tipoParametro, 34).Tables(0)
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
            d.idcliente = IIf(dt.Rows(0).Item("idcliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcliente"))
            d.concepto = IIf(dt.Rows(0).Item("concepto") Is DBNull.Value, Nothing, dt.Rows(0).Item("concepto"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
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
            d.idcliente = Nothing
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
    Public Sub Eliminar(d As NAduana_Detalleanticipos)
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item}
        sql.EjecutarProcedure("Str_Aduana_Detalleanticipos_D", parametros, valores, tipoParametro, 4)
    End Sub
    Public Function Existe_Aduana_Detalleanticipos(d As NAduana_Detalleanticipos) As Boolean
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_Aduana_Detalleanticipos", parametros, valores, tipoParametro, 4)
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
        dt = sql.ProcedureSQL("Str_Aduana_Detalleanticipos_S", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NAduana_Detalleanticipos) As DataTable
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Aduana_Detalleanticipos_S", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NAduana_Detalleanticipos) As NAduana_Detalleanticipos
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Aduana_Detalleanticipos_S", parametros, valores, tipoParametro, 4).Tables(0)
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
            d.idcliente = IIf(dt.Rows(0).Item("idcliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcliente"))
            d.concepto = IIf(dt.Rows(0).Item("concepto") Is DBNull.Value, Nothing, dt.Rows(0).Item("concepto"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.IdCuenta = IIf(dt.Rows(0).Item("IdCuenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("IdCuenta"))
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
            d.idcliente = Nothing
            d.concepto = Nothing
            d.tipocambio = Nothing
            d.idmoneda = Nothing
            d.importe = Nothing
            d.importemn = Nothing
            d.importeus = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
            d.IdCuenta = Nothing
        End If
        Return d
    End Function
#End Region


End Class
