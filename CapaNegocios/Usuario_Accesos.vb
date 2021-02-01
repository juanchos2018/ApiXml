Imports CapaDatos
Public Class Usuario_Accesos
    Dim sql As New ClsConexion

    Public Property idusuario As String
    Public Property Estado As String
    Public Property nombre As String
    Public Property observacion As String
    Public Property mensajes As String
    Public Property Fechapago As DateTime
    Public Property FechaCrea As DateTime

    Public Function Registro(d As Usuario_Accesos) As Usuario_Accesos
        Dim parametros() As Object = {"@idusuario"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idusuario}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_GetUsuario", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            '  d.idusuario = IIf(dt.Rows(0).Item("idusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuario"))
            d.Estado = IIf(dt.Rows(0).Item("Estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("Estado"))
            d.nombre = IIf(dt.Rows(0).Item("nombre") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombre"))
            '   d.observacion = IIf(dt.Rows(0).Item("observacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("observacion"))
            d.mensajes = IIf(dt.Rows(0).Item("mensajes") Is DBNull.Value, Nothing, dt.Rows(0).Item("mensajes"))
            d.Fechapago = IIf(dt.Rows(0).Item("Fechapago") Is DBNull.Value, Nothing, dt.Rows(0).Item("Fechapago"))
            '  d.FechaCrea = IIf(dt.Rows(0).Item("FechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("FechaCrea"))
        Else
            d.idusuario = Nothing
            d.Estado = Nothing
            d.nombre = Nothing
            d.observacion = Nothing
            d.mensajes = Nothing
            d.Fechapago = Nothing
            d.FechaCrea = Nothing
        End If
        Return d
    End Function
End Class
