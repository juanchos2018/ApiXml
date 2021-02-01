Imports CapaDatos
Public Class NOrdenCompra_CotSap
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property cotizacion_sap As String
    Public Property item As String
    Public Property idarticulo As String
    Public Property articulo As String
    Public Property cantidad As Decimal
    Public Property saldo As Decimal
    Public Property preciounitario As Decimal
    Public Property total As Decimal
    Public Property numeroproceso As String
    Public Property estado As Boolean
    Public Property fechacrea As System.DateTime
    Public Property usuariocrea As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region


#Region "Metodos"
    Public Sub Agregar(d As NOrdenCompra_CotSap)

        Dim parametros() As Object = {"@cotizacion_sap", "@item", "@idarticulo", "@articulo", "@cantidad", "@saldo", "@preciounitario", "@total", "@numeroproceso", "@estado", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.cotizacion_sap, d.item, d.idarticulo, d.articulo, d.cantidad, d.saldo, d.preciounitario, d.total, d.numeroproceso, d.estado, d.fechacrea, d.usuariocrea}
        sql.EjecutarProcedure("Str_OrdenCompra_CotSap_I", parametros, valores, tipoParametro, 11)
    End Sub
    Public Sub Actualizar(d As NOrdenCompra_CotSap)
        Dim parametros() As Object = {"@cotizacion_sap", "@item", "@idarticulo", "@articulo", "@cantidad", "@saldo", "@preciounitario", "@total", "@numeroproceso", "@estado", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.cotizacion_sap, d.item, d.idarticulo, d.articulo, d.cantidad, d.saldo, d.preciounitario, d.total, d.numeroproceso, d.estado, d.fechacrea, d.usuariocrea}
        sql.EjecutarProcedure("Str_OrdenCompra_CotSap_U", parametros, valores, tipoParametro, 11)
    End Sub
    Public Function Agregar(d As NOrdenCompra_CotSap, Retornatable As Boolean) As NOrdenCompra_CotSap

        Dim parametros() As Object = {"@cotizacion_sap", "@item", "@idarticulo", "@articulo", "@cantidad", "@saldo", "@preciounitario", "@total", "@numeroproceso", "@estado", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.cotizacion_sap, d.item, d.idarticulo, d.articulo, d.cantidad, d.saldo, d.preciounitario, d.total, d.numeroproceso, d.estado, d.fechacrea, d.usuariocrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_OrdenCompra_CotSap_I_S", parametros, valores, tipoParametro, 11).Tables(0)
        If dt.Rows.Count > 0 Then
            d.cotizacion_sap = IIf(dt.Rows(0).Item("cotizacion_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("cotizacion_sap"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.articulo = IIf(dt.Rows(0).Item("articulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("articulo"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
            d.total = IIf(dt.Rows(0).Item("total") Is DBNull.Value, Nothing, dt.Rows(0).Item("total"))
            d.numeroproceso = IIf(dt.Rows(0).Item("numeroproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroproceso"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else

            d.cotizacion_sap = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.articulo = Nothing
            d.cantidad = Nothing
            d.saldo = Nothing
            d.preciounitario = Nothing
            d.total = Nothing
            d.numeroproceso = Nothing
            d.estado = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NOrdenCompra_CotSap, Retornatable As Boolean) As NOrdenCompra_CotSap

        Dim parametros() As Object = {"@cotizacion_sap", "@item", "@idarticulo", "@articulo", "@cantidad", "@saldo", "@preciounitario", "@total", "@numeroproceso", "@estado", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.usuariocrea = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_OrdenCompra_CotSap_U_S", parametros, valores, tipoParametro, 11).Tables(0)
        If dt.Rows.Count > 0 Then

            d.cotizacion_sap = IIf(dt.Rows(0).Item("cotizacion_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("cotizacion_sap"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.articulo = IIf(dt.Rows(0).Item("articulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("articulo"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
            d.total = IIf(dt.Rows(0).Item("total") Is DBNull.Value, Nothing, dt.Rows(0).Item("total"))
            d.numeroproceso = IIf(dt.Rows(0).Item("numeroproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroproceso"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else

            d.cotizacion_sap = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.articulo = Nothing
            d.cantidad = Nothing
            d.saldo = Nothing
            d.preciounitario = Nothing
            d.total = Nothing
            d.numeroproceso = Nothing
            d.estado = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NOrdenCompra_CotSap)
        Dim parametros() As Object = {"@cotizacion_sap"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.cotizacion_sap}
        sql.EjecutarProcedure("Str_OrdenCompra_CotSap_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_OrdenCompra_CotSap(d As NOrdenCompra_CotSap) As Boolean
        Dim parametros() As Object = {"@cotizacion_sap"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.cotizacion_sap}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_OrdenCompra_CotSap", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@cotizacion_sap"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_OrdenCompra_CotSap_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NOrdenCompra_CotSap) As DataTable
        Dim parametros() As Object = {"@cotizacion_sap"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.cotizacion_sap}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_OrdenCompra_CotSap_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NOrdenCompra_CotSap) As NOrdenCompra_CotSap
        Dim parametros() As Object = {"@cotizacion_sap"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.cotizacion_sap}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_OrdenCompra_CotSap_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.cotizacion_sap = IIf(dt.Rows(0).Item("cotizacion_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("cotizacion_sap"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.articulo = IIf(dt.Rows(0).Item("articulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("articulo"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
            d.total = IIf(dt.Rows(0).Item("total") Is DBNull.Value, Nothing, dt.Rows(0).Item("total"))
            d.numeroproceso = IIf(dt.Rows(0).Item("numeroproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroproceso"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.cotizacion_sap = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.articulo = Nothing
            d.cantidad = Nothing
            d.saldo = Nothing
            d.preciounitario = Nothing
            d.total = Nothing
            d.numeroproceso = Nothing
            d.estado = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
#End Region

End Class
