Imports CapaDatos
Public Class NDetalleMovimientoEspecifico
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idseriearticulo As Integer
    Public Property idagencia As String
    Public Property idalmacen As String
    Public Property tipodocumento As String
    Public Property numerodocumento As String
    Public Property item As String
    Public Property idarticulo As String
    Public Property nroserie As String
    Public Property cantidad As Decimal
    Public Property tipodocref As String
    Public Property nrodocref As String
    Public Property fecharef As System.DateTime
    Public Property idseriearticuloref As Integer
    Public Property saldo As Decimal

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NDetalleMovimientoEspecifico)

        Dim parametros() As Object = {"@idagencia", "@idalmacen", "@tipodocumento", "@numerodocumento", "@item", "@idarticulo", "@nroserie", "@cantidad", "@tipodocref", "@nrodocref", "@fecharef", "@idseriearticuloref", "@saldo"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Int, SqlDbType.Decimal}
        Dim valores() As Object = {d.idagencia, d.idalmacen, d.tipodocumento, d.numerodocumento, d.item, d.idarticulo, d.nroserie, d.cantidad, d.tipodocref, d.nrodocref, d.fecharef, d.idseriearticuloref, d.saldo}
        sql.EjecutarProcedure("Str_DetalleMovimientoEspecifico_I", parametros, valores, tipoParametro, 13)
    End Sub
    Public Sub Actualizar(d As NDetalleMovimientoEspecifico)
        Dim parametros() As Object = {"@idagencia", "@idalmacen", "@tipodocumento", "@numerodocumento", "@item", "@idarticulo", "@nroserie", "@cantidad", "@tipodocref", "@nrodocref", "@fecharef", "@idseriearticuloref", "@saldo"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Int, SqlDbType.Decimal}
        Dim valores() As Object = {d.idagencia, d.idalmacen, d.tipodocumento, d.numerodocumento, d.item, d.idarticulo, d.nroserie, d.cantidad, d.tipodocref, d.nrodocref, d.fecharef, d.idseriearticuloref, d.saldo}
        sql.EjecutarProcedure("Str_DetalleMovimientoEspecifico_U", parametros, valores, tipoParametro, 13)
    End Sub
    Public Function Agregar(d As NDetalleMovimientoEspecifico, Retornatable As Boolean) As NDetalleMovimientoEspecifico

        Dim parametros() As Object = {"@idagencia", "@idalmacen", "@tipodocumento", "@numerodocumento", "@item", "@idarticulo", "@nroserie", "@cantidad", "@tipodocref", "@nrodocref", "@fecharef", "@idseriearticuloref", "@saldo"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Int, SqlDbType.Decimal}
        Dim valores() As Object = {d.idagencia, d.idalmacen, d.tipodocumento, d.numerodocumento, d.item, d.idarticulo, d.nroserie, d.cantidad, d.tipodocref, d.nrodocref, d.fecharef, d.idseriearticuloref, d.saldo}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleMovimientoEspecifico_I_S", parametros, valores, tipoParametro, 13).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idseriearticulo = IIf(dt.Rows(0).Item("idseriearticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idseriearticulo"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.tipodocumento = IIf(dt.Rows(0).Item("tipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocumento"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.nroserie = IIf(dt.Rows(0).Item("nroserie") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroserie"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.tipodocref = IIf(dt.Rows(0).Item("tipodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocref"))
            d.nrodocref = IIf(dt.Rows(0).Item("nrodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocref"))
            d.fecharef = IIf(dt.Rows(0).Item("fecharef") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecharef"))
            d.idseriearticuloref = IIf(dt.Rows(0).Item("idseriearticuloref") Is DBNull.Value, Nothing, dt.Rows(0).Item("idseriearticuloref"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
        Else
            d.idseriearticulo = Nothing
            d.idagencia = Nothing
            d.idalmacen = Nothing
            d.tipodocumento = Nothing
            d.numerodocumento = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.nroserie = Nothing
            d.cantidad = Nothing
            d.tipodocref = Nothing
            d.nrodocref = Nothing
            d.fecharef = Nothing
            d.idseriearticuloref = Nothing
            d.saldo = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NDetalleMovimientoEspecifico, Retornatable As Boolean) As NDetalleMovimientoEspecifico

        Dim parametros() As Object = {"@idagencia", "@idalmacen", "@tipodocumento", "@numerodocumento", "@item", "@idarticulo", "@nroserie", "@cantidad", "@tipodocref", "@nrodocref", "@fecharef", "@idseriearticuloref", "@saldo"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Int, SqlDbType.Decimal}
        Dim valores() As Object = {d.saldo = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleMovimientoEspecifico_U_S", parametros, valores, tipoParametro, 41).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idseriearticulo = IIf(dt.Rows(0).Item("idseriearticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idseriearticulo"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.tipodocumento = IIf(dt.Rows(0).Item("tipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocumento"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.nroserie = IIf(dt.Rows(0).Item("nroserie") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroserie"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.tipodocref = IIf(dt.Rows(0).Item("tipodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocref"))
            d.nrodocref = IIf(dt.Rows(0).Item("nrodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocref"))
            d.fecharef = IIf(dt.Rows(0).Item("fecharef") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecharef"))
            d.idseriearticuloref = IIf(dt.Rows(0).Item("idseriearticuloref") Is DBNull.Value, Nothing, dt.Rows(0).Item("idseriearticuloref"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
        Else
            d.idseriearticulo = Nothing
            d.idagencia = Nothing
            d.idalmacen = Nothing
            d.tipodocumento = Nothing
            d.numerodocumento = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.nroserie = Nothing
            d.cantidad = Nothing
            d.tipodocref = Nothing
            d.nrodocref = Nothing
            d.fecharef = Nothing
            d.idseriearticuloref = Nothing
            d.saldo = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NDetalleMovimientoEspecifico)
        Dim parametros() As Object = {"@idseriearticulo"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.idseriearticulo}
        sql.EjecutarProcedure("Str_DetalleMovimientoEspecifico_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_DetalleMovimientoEspecifico(d As NDetalleMovimientoEspecifico)
        Dim parametros() As Object = {"@idseriearticulo"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.idseriearticulo}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_DetalleMovimientoEspecifico", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idseriearticulo"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleMovimientoEspecifico_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NDetalleMovimientoEspecifico) As DataTable
        Dim parametros() As Object = {"@idseriearticulo"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.idseriearticulo}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleMovimientoEspecifico_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDetalleMovimientoEspecifico) As NDetalleMovimientoEspecifico
        Dim parametros() As Object = {"@idseriearticulo"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.idseriearticulo}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleMovimientoEspecifico_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idseriearticulo = IIf(dt.Rows(0).Item("idseriearticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idseriearticulo"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.tipodocumento = IIf(dt.Rows(0).Item("tipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocumento"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.nroserie = IIf(dt.Rows(0).Item("nroserie") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroserie"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.tipodocref = IIf(dt.Rows(0).Item("tipodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocref"))
            d.nrodocref = IIf(dt.Rows(0).Item("nrodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocref"))
            d.fecharef = IIf(dt.Rows(0).Item("fecharef") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecharef"))
            d.idseriearticuloref = IIf(dt.Rows(0).Item("idseriearticuloref") Is DBNull.Value, Nothing, dt.Rows(0).Item("idseriearticuloref"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
        Else
            d.idseriearticulo = Nothing
            d.idagencia = Nothing
            d.idalmacen = Nothing
            d.tipodocumento = Nothing
            d.numerodocumento = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.nroserie = Nothing
            d.cantidad = Nothing
            d.tipodocref = Nothing
            d.nrodocref = Nothing
            d.fecharef = Nothing
            d.idseriearticuloref = Nothing
            d.saldo = Nothing
        End If
        Return d
    End Function

    Public Function Movimiento_Serie(i As DateTime, f As DateTime, ai As String, af As String, d As DataTable, arti As DataTable) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF", "@IdAlmacenI", "@IdAlmacenF", "@IdMov", "@arti"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Structured, SqlDbType.Structured}
        Dim valores() As Object = {i, f, ai, af, d, arti}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_MovimientoAlmacenSerie", parametros, valores, tipoParametro, 6).Tables(0)
        Return dt
    End Function

#End Region

End Class
