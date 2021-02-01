Imports CapaDatos
Public Class NTran
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property numero As Integer
    Public Property cara As String
    Public Property manguera As String
    Public Property producto As String
    Public Property soles As Decimal
    Public Property galones As Decimal
    Public Property precio As Decimal
    Public Property fecha As String
    Public Property hora As String
    Public Property estado As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NTran)
        Dim parametros() As Object = {"@cara", "@manguera", "@producto", "@soles", "@galones", "@precio", "@fecha", "@hora", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.NChar, SqlDbType.NChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.NChar, SqlDbType.NChar, SqlDbType.NChar}
        Dim valores() As Object = {d.cara, d.manguera, d.producto, d.soles, d.galones, d.precio, d.fecha, d.hora, d.estado}
        sql.EjecutarProcedure("Str_TRAN_I", parametros, valores, tipoParametro, 9)
    End Sub
    Public Sub Actualizar(d As NTran)
        Dim parametros() As Object = {"@cara", "@manguera", "@producto", "@soles", "@galones", "@precio", "@fecha", "@hora", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.NChar, SqlDbType.NChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.NChar, SqlDbType.NChar, SqlDbType.NChar}
        Dim valores() As Object = {d.cara, d.manguera, d.producto, d.soles, d.galones, d.precio, d.fecha, d.hora, d.estado}
        sql.EjecutarProcedure("Str_TRAN_U", parametros, valores, tipoParametro, 9)
    End Sub
    Public Function Agregar(d As NTran, Retornatable As Boolean) As NTran

        Dim parametros() As Object = {"@cara", "@manguera", "@producto", "@soles", "@galones", "@precio", "@fecha", "@hora", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.NChar, SqlDbType.NChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.NChar, SqlDbType.NChar, SqlDbType.NChar}
        Dim valores() As Object = {d.cara, d.manguera, d.producto, d.soles, d.galones, d.precio, d.fecha, d.hora, d.estado}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TRAN_I_S", parametros, valores, tipoParametro, 9).Tables(0)
        If dt.Rows.Count > 0 Then
            d.numero = IIf(dt.Rows(0).Item("numero") Is DBNull.Value, Nothing, dt.Rows(0).Item("numero"))
            d.cara = IIf(dt.Rows(0).Item("cara") Is DBNull.Value, Nothing, dt.Rows(0).Item("cara"))
            d.manguera = IIf(dt.Rows(0).Item("manguera") Is DBNull.Value, Nothing, dt.Rows(0).Item("manguera"))
            d.producto = IIf(dt.Rows(0).Item("producto") Is DBNull.Value, Nothing, dt.Rows(0).Item("producto"))
            d.soles = IIf(dt.Rows(0).Item("soles") Is DBNull.Value, Nothing, dt.Rows(0).Item("soles"))
            d.galones = IIf(dt.Rows(0).Item("galones") Is DBNull.Value, Nothing, dt.Rows(0).Item("galones"))
            d.precio = IIf(dt.Rows(0).Item("precio") Is DBNull.Value, Nothing, dt.Rows(0).Item("precio"))
            d.fecha = IIf(dt.Rows(0).Item("fecha") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecha"))
            d.hora = IIf(dt.Rows(0).Item("hora") Is DBNull.Value, Nothing, dt.Rows(0).Item("hora"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
        Else
            d.numero = Nothing
            d.cara = Nothing
            d.manguera = Nothing
            d.producto = Nothing
            d.soles = Nothing
            d.galones = Nothing
            d.precio = Nothing
            d.fecha = Nothing
            d.hora = Nothing
            d.estado = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NTran, Retornatable As Boolean) As NTran

        Dim parametros() As Object = {"@cara", "@manguera", "@producto", "@soles", "@galones", "@precio", "@fecha", "@hora", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.NChar, SqlDbType.NChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.NChar, SqlDbType.NChar, SqlDbType.NChar}
        Dim valores() As Object = {d.estado = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TRAN_U_S", parametros, valores, tipoParametro, 29).Tables(0)
        If dt.Rows.Count > 0 Then
            d.numero = IIf(dt.Rows(0).Item("numero") Is DBNull.Value, Nothing, dt.Rows(0).Item("numero"))
            d.cara = IIf(dt.Rows(0).Item("cara") Is DBNull.Value, Nothing, dt.Rows(0).Item("cara"))
            d.manguera = IIf(dt.Rows(0).Item("manguera") Is DBNull.Value, Nothing, dt.Rows(0).Item("manguera"))
            d.producto = IIf(dt.Rows(0).Item("producto") Is DBNull.Value, Nothing, dt.Rows(0).Item("producto"))
            d.soles = IIf(dt.Rows(0).Item("soles") Is DBNull.Value, Nothing, dt.Rows(0).Item("soles"))
            d.galones = IIf(dt.Rows(0).Item("galones") Is DBNull.Value, Nothing, dt.Rows(0).Item("galones"))
            d.precio = IIf(dt.Rows(0).Item("precio") Is DBNull.Value, Nothing, dt.Rows(0).Item("precio"))
            d.fecha = IIf(dt.Rows(0).Item("fecha") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecha"))
            d.hora = IIf(dt.Rows(0).Item("hora") Is DBNull.Value, Nothing, dt.Rows(0).Item("hora"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
        Else
            d.numero = Nothing
            d.cara = Nothing
            d.manguera = Nothing
            d.producto = Nothing
            d.soles = Nothing
            d.galones = Nothing
            d.precio = Nothing
            d.fecha = Nothing
            d.hora = Nothing
            d.estado = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NTran)
        Dim parametros() As Object = {"@numero"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.numero}
        sql.EjecutarProcedure("Str_TRAN_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@numero"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TRAN_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTran) As DataTable
        Dim parametros() As Object = {"@numero"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.numero}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TRAN_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTran) As NTran
        Dim parametros() As Object = {"@numero"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.numero}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TRAN_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.numero = IIf(dt.Rows(0).Item("numero") Is DBNull.Value, Nothing, dt.Rows(0).Item("numero"))
            d.cara = IIf(dt.Rows(0).Item("cara") Is DBNull.Value, Nothing, dt.Rows(0).Item("cara"))
            d.manguera = IIf(dt.Rows(0).Item("manguera") Is DBNull.Value, Nothing, dt.Rows(0).Item("manguera"))
            d.producto = IIf(dt.Rows(0).Item("producto") Is DBNull.Value, Nothing, dt.Rows(0).Item("producto"))
            d.soles = IIf(dt.Rows(0).Item("soles") Is DBNull.Value, Nothing, dt.Rows(0).Item("soles"))
            d.galones = IIf(dt.Rows(0).Item("galones") Is DBNull.Value, Nothing, dt.Rows(0).Item("galones"))
            d.precio = IIf(dt.Rows(0).Item("precio") Is DBNull.Value, Nothing, dt.Rows(0).Item("precio"))
            d.fecha = IIf(dt.Rows(0).Item("fecha") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecha"))
            d.hora = IIf(dt.Rows(0).Item("hora") Is DBNull.Value, Nothing, dt.Rows(0).Item("hora"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
        Else
            d.numero = Nothing
            d.cara = Nothing
            d.manguera = Nothing
            d.producto = Nothing
            d.soles = Nothing
            d.galones = Nothing
            d.precio = Nothing
            d.fecha = Nothing
            d.hora = Nothing
            d.estado = Nothing
        End If
        Return d
    End Function
#End Region


End Class
