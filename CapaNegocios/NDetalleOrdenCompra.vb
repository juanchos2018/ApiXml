Imports CapaDatos
Public Class NDetalleOrdenCompra
    Dim sql As New ClsConexion

#Region "Declarations"

    Public Property iddetoc As Integer
    Public Property item As String
    Public Property idarticulo As String
    Public Property descripcion As String
    Public Property cantidad As Decimal
    Public Property cantidadatendida As Decimal
    Public Property unidad As String
    Public Property igv As Decimal
    Public Property preciounitario As Decimal
    Public Property tipocambio As Decimal
    Public Property valorventa As Decimal
    Public Property importeigv As Decimal
    Public Property total As Decimal
    Public Property estado As String
    Public Property idoc As String
    Public Property idmoneda As String
    Public Property valorventamn As Decimal
    Public Property valorventaus As Decimal
    Public Property importeigvmn As Decimal
    Public Property importeigvus As Decimal
    Public Property totalmn As Decimal
    Public Property totalus As Decimal
    Public Property nroproceso As String
    Public Property cotizacion_sap As String
    Public Property pedido_sap As String
    Public Property idmodalidad As String
    Public Property lugar_entrega As String
    Public Property Observaciondet As String
    Public Property SerieCompra As String
    Public Property NroDocCompra As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NDetalleOrdenCompra)

        Dim parametros() As Object = {"@item", "@idarticulo", "@descripcion", "@cantidad", "@cantidadatendida", "@unidad", "@igv", "@preciounitario", "@tipocambio", "@valorventa", "@importeigv", "@total", "@estado", "@idoc", "@idmoneda", "@nroproceso", "@cotizacion_sap", "@pedido_sap", "@idmodalidad", "@lugar_entrega", "@Observaciondet"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.NVarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.item, d.idarticulo, d.descripcion, d.cantidad, d.cantidadatendida, d.unidad, d.igv, d.preciounitario, d.tipocambio, d.valorventa, d.importeigv, d.total, d.estado, d.idoc, d.idmoneda, d.nroproceso, d.cotizacion_sap, d.pedido_sap, d.idmodalidad, d.lugar_entrega, d.Observaciondet}
        sql.EjecutarProcedure("Str_DetalleOrdenCompra_I", parametros, valores, tipoParametro, 21)
    End Sub
    Public Sub Actualizar(d As NDetalleOrdenCompra)
        Dim parametros() As Object = {"@IdDetOC", "@item", "@idarticulo", "@descripcion", "@cantidad", "@cantidadatendida", "@unidad", "@igv", "@preciounitario", "@tipocambio", "@valorventa", "@importeigv", "@total", "@estado", "@idoc", "@idmoneda", "@nroproceso", "@cotizacion_sap", "@pedido_sap", "@idmodalidad", "@lugar_entrega", "@Observaciondet", "@SerieCompra", "@NroDocCompra"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.NVarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.iddetoc, d.item, d.idarticulo, d.descripcion, d.cantidad, d.cantidadatendida, d.unidad, d.igv, d.preciounitario, d.tipocambio, d.valorventa, d.importeigv, d.total, d.estado, d.idoc, d.idmoneda, d.nroproceso, d.cotizacion_sap, d.pedido_sap, d.idmodalidad, d.lugar_entrega, d.Observaciondet, d.SerieCompra, d.NroDocCompra}
        sql.EjecutarProcedure("Str_DetalleOrdenCompra_U", parametros, valores, tipoParametro, 24)
    End Sub
    Public Function Agregar(d As NDetalleOrdenCompra, Retornatable As Boolean) As NDetalleOrdenCompra

        Dim parametros() As Object = {"@item", "@idarticulo", "@descripcion", "@cantidad", "@cantidadatendida", "@unidad", "@igv", "@preciounitario", "@tipocambio", "@valorventa", "@importeigv", "@total", "@estado", "@idoc", "@idmoneda", "@nroproceso", "@cotizacion_sap", "@pedido_sap"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.NVarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.item, d.idarticulo, d.descripcion, d.cantidad, d.cantidadatendida, d.unidad, d.igv, d.preciounitario, d.tipocambio, d.valorventa, d.importeigv, d.total, d.estado, d.idoc, d.idmoneda, d.nroproceso, d.cotizacion_sap, d.pedido_sap}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_DetalleOrdenCompra_I_S", parametros, valores, tipoParametro, 18).Tables(0)
        If dt.Rows.Count > 0 Then
            d.iddetoc = IIf(dt.Rows(0).Item("iddetoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetoc"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.cantidadatendida = IIf(dt.Rows(0).Item("cantidadatendida") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidadatendida"))
            d.unidad = IIf(dt.Rows(0).Item("unidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidad"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.valorventa = IIf(dt.Rows(0).Item("valorventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventa"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.total = IIf(dt.Rows(0).Item("total") Is DBNull.Value, Nothing, dt.Rows(0).Item("total"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.idoc = IIf(dt.Rows(0).Item("idoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("idoc"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.valorventamn = IIf(dt.Rows(0).Item("valorventamn") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventamn"))
            d.valorventaus = IIf(dt.Rows(0).Item("valorventaus") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventaus"))
            d.importeigvmn = IIf(dt.Rows(0).Item("importeigvmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvmn"))
            d.importeigvus = IIf(dt.Rows(0).Item("importeigvus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvus"))
            d.totalmn = IIf(dt.Rows(0).Item("totalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("totalmn"))
            d.totalus = IIf(dt.Rows(0).Item("totalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("totalus"))
            d.nroproceso = IIf(dt.Rows(0).Item("nroproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroproceso"))
            d.cotizacion_sap = IIf(dt.Rows(0).Item("cotizacion_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("cotizacion_sap"))
            d.pedido_sap = IIf(dt.Rows(0).Item("pedido_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("pedido_sap"))
        Else
            d.iddetoc = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
            d.cantidad = Nothing
            d.cantidadatendida = Nothing
            d.unidad = Nothing
            d.igv = Nothing
            d.preciounitario = Nothing
            d.tipocambio = Nothing
            d.valorventa = Nothing
            d.importeigv = Nothing
            d.total = Nothing
            d.estado = Nothing
            d.idoc = Nothing
            d.idmoneda = Nothing
            d.valorventamn = Nothing
            d.valorventaus = Nothing
            d.importeigvmn = Nothing
            d.importeigvus = Nothing
            d.totalmn = Nothing
            d.totalus = Nothing
            d.nroproceso = Nothing
            d.cotizacion_sap = Nothing
            d.pedido_sap = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NDetalleOrdenCompra, Retornatable As Boolean) As NDetalleOrdenCompra

        Dim parametros() As Object = {"@item", "@idarticulo", "@descripcion", "@cantidad", "@cantidadatendida", "@unidad", "@igv", "@preciounitario", "@tipocambio", "@valorventa", "@importeigv", "@total", "@estado", "@idoc", "@idmoneda", "@nroproceso", "@cotizacion_sap", "@pedido_sap"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.NVarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.pedido_sap = Nothing}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_DetalleOrdenCompra_U_S", parametros, valores, tipoParametro, 68).Tables(0)
        If dt.Rows.Count > 0 Then
            d.iddetoc = IIf(dt.Rows(0).Item("iddetoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetoc"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.cantidadatendida = IIf(dt.Rows(0).Item("cantidadatendida") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidadatendida"))
            d.unidad = IIf(dt.Rows(0).Item("unidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidad"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.valorventa = IIf(dt.Rows(0).Item("valorventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventa"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.total = IIf(dt.Rows(0).Item("total") Is DBNull.Value, Nothing, dt.Rows(0).Item("total"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.idoc = IIf(dt.Rows(0).Item("idoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("idoc"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.valorventamn = IIf(dt.Rows(0).Item("valorventamn") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventamn"))
            d.valorventaus = IIf(dt.Rows(0).Item("valorventaus") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventaus"))
            d.importeigvmn = IIf(dt.Rows(0).Item("importeigvmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvmn"))
            d.importeigvus = IIf(dt.Rows(0).Item("importeigvus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvus"))
            d.totalmn = IIf(dt.Rows(0).Item("totalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("totalmn"))
            d.totalus = IIf(dt.Rows(0).Item("totalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("totalus"))
            d.nroproceso = IIf(dt.Rows(0).Item("nroproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroproceso"))
            d.cotizacion_sap = IIf(dt.Rows(0).Item("cotizacion_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("cotizacion_sap"))
            d.pedido_sap = IIf(dt.Rows(0).Item("pedido_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("pedido_sap"))
        Else
            d.iddetoc = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
            d.cantidad = Nothing
            d.cantidadatendida = Nothing
            d.unidad = Nothing
            d.igv = Nothing
            d.preciounitario = Nothing
            d.tipocambio = Nothing
            d.valorventa = Nothing
            d.importeigv = Nothing
            d.total = Nothing
            d.estado = Nothing
            d.idoc = Nothing
            d.idmoneda = Nothing
            d.valorventamn = Nothing
            d.valorventaus = Nothing
            d.importeigvmn = Nothing
            d.importeigvus = Nothing
            d.totalmn = Nothing
            d.totalus = Nothing
            d.nroproceso = Nothing
            d.cotizacion_sap = Nothing
            d.pedido_sap = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NDetalleOrdenCompra)
        Dim parametros() As Object = {"@IdOC"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idoc}
        sql.EjecutarProcedure("Str_DetalleOrdenCompra_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_DetalleOrdenCompra(d As NDetalleOrdenCompra) As Boolean
        Dim parametros() As Object = {"@iddetoc"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.iddetoc}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = Sql.procedimiento_escalar("Existe_DetalleOrdenCompra", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@iddetoc"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_DetalleOrdenCompra_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NDetalleOrdenCompra) As DataTable
        Dim parametros() As Object = {"@iddetoc", "@idoc"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar}
        Dim valores() As Object = {d.iddetoc, d.idoc}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleOrdenCompra_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDetalleOrdenCompra) As NDetalleOrdenCompra
        Dim parametros() As Object = {"@iddetoc", "@IdOC"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar}
        Dim valores() As Object = {d.iddetoc, d.idoc}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleOrdenCompra_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.iddetoc = IIf(dt.Rows(0).Item("iddetoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetoc"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.cantidadatendida = IIf(dt.Rows(0).Item("cantidadatendida") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidadatendida"))
            d.unidad = IIf(dt.Rows(0).Item("unidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidad"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.valorventa = IIf(dt.Rows(0).Item("valorventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventa"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.total = IIf(dt.Rows(0).Item("total") Is DBNull.Value, Nothing, dt.Rows(0).Item("total"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.idoc = IIf(dt.Rows(0).Item("idoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("idoc"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.valorventamn = IIf(dt.Rows(0).Item("valorventamn") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventamn"))
            d.valorventaus = IIf(dt.Rows(0).Item("valorventaus") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventaus"))
            d.importeigvmn = IIf(dt.Rows(0).Item("importeigvmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvmn"))
            d.importeigvus = IIf(dt.Rows(0).Item("importeigvus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvus"))
            d.totalmn = IIf(dt.Rows(0).Item("totalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("totalmn"))
            d.totalus = IIf(dt.Rows(0).Item("totalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("totalus"))
            d.nroproceso = IIf(dt.Rows(0).Item("nroproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroproceso"))
            d.cotizacion_sap = IIf(dt.Rows(0).Item("cotizacion_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("cotizacion_sap"))
            d.pedido_sap = IIf(dt.Rows(0).Item("pedido_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("pedido_sap"))
            d.idmodalidad = IIf(dt.Rows(0).Item("idmodalidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmodalidad"))
            d.lugar_entrega = IIf(dt.Rows(0).Item("lugar_entrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugar_entrega"))
            d.Observaciondet = IIf(dt.Rows(0).Item("Observaciondet") Is DBNull.Value, Nothing, dt.Rows(0).Item("Observaciondet"))
            d.SerieCompra = IIf(dt.Rows(0).Item("SerieCompra") Is DBNull.Value, Nothing, dt.Rows(0).Item("SerieCompra"))
            d.NroDocCompra = IIf(dt.Rows(0).Item("NroDocCompra") Is DBNull.Value, Nothing, dt.Rows(0).Item("NroDocCompra"))
        Else
            d.iddetoc = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
            d.cantidad = Nothing
            d.cantidadatendida = Nothing
            d.unidad = Nothing
            d.igv = Nothing
            d.preciounitario = Nothing
            d.tipocambio = Nothing
            d.valorventa = Nothing
            d.importeigv = Nothing
            d.total = Nothing
            d.estado = Nothing
            d.idoc = Nothing
            d.idmoneda = Nothing
            d.valorventamn = Nothing
            d.valorventaus = Nothing
            d.importeigvmn = Nothing
            d.importeigvus = Nothing
            d.totalmn = Nothing
            d.totalus = Nothing
            d.nroproceso = Nothing
            d.cotizacion_sap = Nothing
            d.pedido_sap = Nothing
            d.idmodalidad = Nothing
            d.lugar_entrega = Nothing
            d.Observaciondet = Nothing
            d.SerieCompra = Nothing
            d.NroDocCompra = Nothing
        End If
        Return d
    End Function
    Public Function Lista_detalle_ordenes(fechaI As DateTime, fechaF As DateTime, Optional op As Boolean = False, Optional dtv As DataTable = Nothing) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF", "@OptionYura", "@IdItem"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Bit, SqlDbType.Structured}
        Dim valores() As Object = {fechaI, fechaF, op, dtv}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Lista_Proceso_detalle", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function


#End Region

End Class
