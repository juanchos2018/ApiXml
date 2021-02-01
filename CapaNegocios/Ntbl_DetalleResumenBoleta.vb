Imports CapaDatos
Public Class Ntbl_DetalleResumenBoleta
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numerodocumento As String
    Public Property nrodocumento As String
    Public Property fechacrea As System.DateTime
    Public Property usuariocrea As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As Ntbl_DetalleResumenBoleta)

        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@nrodocumento", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.nrodocumento, d.fechacrea, d.usuariocrea}
        sql.EjecutarProcedure("Str_tbl_DetalleResumenBoleta_I", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Actualizar(d As Ntbl_DetalleResumenBoleta)
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@nrodocumento", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.nrodocumento, d.fechacrea, d.usuariocrea}
        sql.EjecutarProcedure("Str_tbl_DetalleResumenBoleta_U", parametros, valores, tipoParametro, 6)
    End Sub
    Public Function Agregar(d As Ntbl_DetalleResumenBoleta, Retornatable As Boolean) As Ntbl_DetalleResumenBoleta

        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@nrodocumento", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.nrodocumento, d.fechacrea, d.usuariocrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DetalleResumenBoleta_I_S", parametros, valores, tipoParametro, 6).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.nrodocumento = IIf(dt.Rows(0).Item("nrodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocumento"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.nrodocumento = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_DetalleResumenBoleta, Retornatable As Boolean) As Ntbl_DetalleResumenBoleta

        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@nrodocumento", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.usuariocrea = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DetalleResumenBoleta_U_S", parametros, valores, tipoParametro, 18).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.nrodocumento = IIf(dt.Rows(0).Item("nrodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocumento"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.nrodocumento = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_DetalleResumenBoleta)
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento}
        sql.EjecutarProcedure("Str_tbl_DetalleResumenBoleta_D", parametros, valores, tipoParametro, 3)
    End Sub
    Public Function Existe_tbl_DetalleResumenBoleta(d As Ntbl_DetalleResumenBoleta) As Boolean
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_tbl_DetalleResumenBoleta", parametros, valores, tipoParametro, 3)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@NroDocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DetalleResumenBoleta_S", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_DetalleResumenBoleta) As DataTable
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@NroDocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.nrodocumento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DetalleResumenBoleta_S", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Ntbl_DetalleResumenBoleta) As Ntbl_DetalleResumenBoleta
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DetalleResumenBoleta_S", parametros, valores, tipoParametro, 3).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.nrodocumento = IIf(dt.Rows(0).Item("nrodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocumento"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.nrodocumento = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
#End Region

End Class
