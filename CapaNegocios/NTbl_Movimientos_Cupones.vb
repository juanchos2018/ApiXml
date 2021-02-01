Imports CapaDatos
Public Class NTbl_Movimientos_Cupones
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property id As Long
    Public Property nrocupon As String
    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numerodocumento As String
    Public Property estado As String
    Public Property usuariocrea As String
    Public Property fechacrea As System.DateTime

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NTbl_Movimientos_Cupones)
        Dim parametros() As Object = {"@nrocupon", "@idtipodocumento", "@serie", "@numerodocumento", "@estado", "@usuariocrea", "@fechacrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.nrocupon, d.idtipodocumento, d.serie, d.numerodocumento, d.estado, d.usuariocrea, d.fechacrea}
        sql.EjecutarProcedure("Str_Tbl_Movimientos_Cupones_I", parametros, valores, tipoParametro, 7)
    End Sub
    Public Sub Actualizar(d As NTbl_Movimientos_Cupones)
        Dim parametros() As Object = {"@id", "@nrocupon", "@idtipodocumento", "@serie", "@numerodocumento", "@estado", "@usuariocrea", "@fechacrea"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.id, d.nrocupon, d.idtipodocumento, d.serie, d.numerodocumento, d.estado, d.usuariocrea, d.fechacrea}
        sql.EjecutarProcedure("Str_Tbl_Movimientos_Cupones_U", parametros, valores, tipoParametro, 8)
    End Sub
    Public Function Agregar(d As NTbl_Movimientos_Cupones, Retornatable As Boolean) As NTbl_Movimientos_Cupones
        Dim parametros() As Object = {"@nrocupon", "@idtipodocumento", "@serie", "@numerodocumento", "@estado", "@usuariocrea", "@fechacrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.nrocupon, d.idtipodocumento, d.serie, d.numerodocumento, d.estado, d.usuariocrea, d.fechacrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Movimientos_Cupones_I_S", parametros, valores, tipoParametro, 7).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.nrocupon = IIf(dt.Rows(0).Item("nrocupon") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrocupon"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
        Else
            d.id = Nothing
            d.nrocupon = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.estado = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NTbl_Movimientos_Cupones, Retornatable As Boolean) As NTbl_Movimientos_Cupones
        Dim parametros() As Object = {"@id", "@nrocupon", "@idtipodocumento", "@serie", "@numerodocumento", "@estado", "@usuariocrea", "@fechacrea"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.id, d.nrocupon, d.idtipodocumento, d.serie, d.numerodocumento, d.estado, d.usuariocrea, d.fechacrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Movimientos_Cupones_U_S", parametros, valores, tipoParametro, 8).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.nrocupon = IIf(dt.Rows(0).Item("nrocupon") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrocupon"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
        Else
            d.id = Nothing
            d.nrocupon = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.estado = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NTbl_Movimientos_Cupones)
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        sql.EjecutarProcedure("Str_Tbl_Movimientos_Cupones_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Movimientos_Cupones_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTbl_Movimientos_Cupones) As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Movimientos_Cupones_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTbl_Movimientos_Cupones) As NTbl_Movimientos_Cupones
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Movimientos_Cupones_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.nrocupon = IIf(dt.Rows(0).Item("nrocupon") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrocupon"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
        Else
            d.id = Nothing
            d.nrocupon = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.estado = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
        End If
        Return d
    End Function
#End Region

End Class
