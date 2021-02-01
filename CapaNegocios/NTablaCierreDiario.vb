Imports CapaDatos
Public Class NTablaCierreDiario
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property fechacierre As System.DateTime
    Public Property idusuario As String
    Public Property estadobloqueo As String
    Public Property estadocaja As Boolean
    Public Property estadoinventario As Boolean
    Public Property estadoventa As Boolean
    Public Property estadocobranza As Boolean
    Public Property estadopago As Boolean
    Public Property estadoproveedor As Boolean

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NTablaCierreDiario)

        Dim parametros() As Object = {"@fechacierre", "@idusuario", "@estadobloqueo", "@estadocaja", "@estadoinventario", "@estadoventa", "@estadocobranza", "@estadopago", "@estadoproveedor"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Char, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.fechacierre, d.idusuario, d.estadobloqueo, d.estadocaja, d.estadoinventario, d.estadoventa, d.estadocobranza, d.estadopago, d.estadoproveedor}
        sql.EjecutarProcedure("Str_TablaCierreDiario_I", parametros, valores, tipoParametro, 9)
    End Sub
    Public Sub Actualizar(d As NTablaCierreDiario)
        Dim parametros() As Object = {"@fechacierre", "@idusuario", "@estadobloqueo", "@estadocaja", "@estadoinventario", "@estadoventa", "@estadocobranza", "@estadopago", "@estadoproveedor"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Char, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.fechacierre, d.idusuario, d.estadobloqueo, d.estadocaja, d.estadoinventario, d.estadoventa, d.estadocobranza, d.estadopago, d.estadoproveedor}
        sql.EjecutarProcedure("Str_TablaCierreDiario_U", parametros, valores, tipoParametro, 9)
    End Sub
    Public Sub Eliminar(d As NTablaCierreDiario)
        Dim parametros() As Object = {"@fechacierre"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime}
        Dim valores() As Object = {d.fechacierre}
        sql.EjecutarProcedure("Str_TablaCierreDiario_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@fechacierre"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TablaCierreDiario_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTablaCierreDiario) As DataTable
        Dim parametros() As Object = {"@fechacierre"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime}
        Dim valores() As Object = {d.fechacierre}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TablaCierreDiario_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTablaCierreDiario) As NTablaCierreDiario
        Dim parametros() As Object = {"@fechacierre"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime}
        Dim valores() As Object = {d.fechacierre}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TablaCierreDiario_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.fechacierre = IIf(dt.Rows(0).Item("fechacierre") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacierre"))
            d.idusuario = IIf(dt.Rows(0).Item("idusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuario"))
            d.estadobloqueo = IIf(dt.Rows(0).Item("estadobloqueo") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadobloqueo"))
            d.estadocaja = IIf(dt.Rows(0).Item("estadocaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadocaja"))
            d.estadoinventario = IIf(dt.Rows(0).Item("estadoinventario") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadoinventario"))
            d.estadoventa = IIf(dt.Rows(0).Item("estadoventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadoventa"))
            d.estadocobranza = IIf(dt.Rows(0).Item("estadocobranza") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadocobranza"))
            d.estadopago = IIf(dt.Rows(0).Item("estadopago") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadopago"))
            d.estadoproveedor = IIf(dt.Rows(0).Item("estadoproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadoproveedor"))
        Else
            d.fechacierre = Nothing
            d.idusuario = Nothing
            d.estadobloqueo = Nothing
            d.estadocaja = Nothing
            d.estadoinventario = Nothing
            d.estadoventa = Nothing
            d.estadocobranza = Nothing
            d.estadopago = Nothing
            d.estadoproveedor = Nothing
        End If
        Return d
    End Function
#End Region


End Class
