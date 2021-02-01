Imports CapaDatos

Public Class NMovimientoCaja
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property tipomovimiento As String
    Public Property idmovimiento As String
    Public Property descripcion As String
    Public Property idcuenta As String
    Public Property idanexo As String
    Public Property idcentrocosto As String
    Public Property idanexoref As String
    Public Property glosa As String
    Public Property tipocomprobante As String
    Public Property usuariocrea As String
    Public Property usuariomod As String
    Public Property fechacrea As System.DateTime
    Public Property fechamod As System.DateTime
    Public Property registromovimiento As String
    Public Property tarifaprecio As Decimal

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NMovimientoCaja)

        Dim parametros() As Object = {"@tipomovimiento", "@idmovimiento", "@descripcion", "@idcuenta", "@idanexo", "@idcentrocosto", "@idanexoref", "@glosa", "@tipocomprobante", "@usuariocrea", "@usuariomod", "@fechacrea", "@fechamod", "@registromovimiento", "@tarifaprecio"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Decimal}
        Dim valores() As Object = {d.tipomovimiento, d.idmovimiento, d.descripcion, d.idcuenta, d.idanexo, d.idcentrocosto, d.idanexoref, d.glosa, d.tipocomprobante, d.usuariocrea, d.usuariomod, d.fechacrea, d.fechamod, d.registromovimiento, d.tarifaprecio}
        sql.EjecutarProcedure("Str_MovimientoCaja_I", parametros, valores, tipoParametro, 15)
    End Sub
    Public Sub Actualizar(d As NMovimientoCaja)
        Dim parametros() As Object = {"@tipomovimiento", "@idmovimiento", "@descripcion", "@idcuenta", "@idanexo", "@idcentrocosto", "@idanexoref", "@glosa", "@tipocomprobante", "@usuariocrea", "@usuariomod", "@fechacrea", "@fechamod", "@registromovimiento", "@tarifaprecio"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Decimal}
        Dim valores() As Object = {d.tipomovimiento, d.idmovimiento, d.descripcion, d.idcuenta, d.idanexo, d.idcentrocosto, d.idanexoref, d.glosa, d.tipocomprobante, d.usuariocrea, d.usuariomod, d.fechacrea, d.fechamod, d.registromovimiento, d.tarifaprecio}
        sql.EjecutarProcedure("Str_MovimientoCaja_U", parametros, valores, tipoParametro, 15)
    End Sub
    Public Sub Eliminar(d As NMovimientoCaja)
        Dim parametros() As Object = {"@tipomovimiento", "@idmovimiento"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipomovimiento, d.idmovimiento}
        sql.EjecutarProcedure("Str_MovimientoCaja_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@tipomovimiento", "@idmovimiento"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_MovimientoCaja_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NMovimientoCaja) As DataTable
        Dim parametros() As Object = {"@tipomovimiento", "@idmovimiento"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipomovimiento, d.idmovimiento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_MovimientoCaja_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NMovimientoCaja) As NMovimientoCaja
        Dim parametros() As Object = {"@tipomovimiento", "@idmovimiento"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipomovimiento, d.idmovimiento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_MovimientoCaja_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.tipomovimiento = IIf(dt.Rows(0).Item("tipomovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipomovimiento"))
            d.idmovimiento = IIf(dt.Rows(0).Item("idmovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmovimiento"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.idcuenta = IIf(dt.Rows(0).Item("idcuenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcuenta"))
            d.idanexo = IIf(dt.Rows(0).Item("idanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idanexo"))
            d.idcentrocosto = IIf(dt.Rows(0).Item("idcentrocosto") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcentrocosto"))
            d.idanexoref = IIf(dt.Rows(0).Item("idanexoref") Is DBNull.Value, Nothing, dt.Rows(0).Item("idanexoref"))
            d.glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.tipocomprobante = IIf(dt.Rows(0).Item("tipocomprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocomprobante"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.registromovimiento = IIf(dt.Rows(0).Item("registromovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("registromovimiento"))
            d.tarifaprecio = IIf(dt.Rows(0).Item("tarifaprecio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tarifaprecio"))
        Else
            d.tipomovimiento = Nothing
            d.idmovimiento = Nothing
            d.descripcion = Nothing
            d.idcuenta = Nothing
            d.idanexo = Nothing
            d.idcentrocosto = Nothing
            d.idanexoref = Nothing
            d.glosa = Nothing
            d.tipocomprobante = Nothing
            d.usuariocrea = Nothing
            d.usuariomod = Nothing
            d.fechacrea = Nothing
            d.fechamod = Nothing
            d.registromovimiento = Nothing
            d.tarifaprecio = Nothing
        End If
        Return d
    End Function
#End Region
End Class
