Imports CapaDatos

Public Class Ntbl_Articulo_UnidadVenta
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property item As Long
    Public Property idarticulo As String
    Public Property idventa As Long
    Public Property preciounitario As Decimal
#End Region

#Region "Constructors"
    Public Sub New()

    End Sub
#End Region
#Region "Metodos"
    Public Sub Agregar(d As Ntbl_Articulo_UnidadVenta)

        Dim parametros() As Object = {"@idarticulo", "@idventa", "@preciounitario"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.BigInt, SqlDbType.Decimal}
        Dim valores() As Object = {d.idarticulo, d.idventa, d.preciounitario}
        sql.EjecutarProcedure("Str_tbl_Articulo_UnidadVenta_I", parametros, valores, tipoParametro, 3)
    End Sub
    Public Sub Actualizar(d As Ntbl_Articulo_UnidadVenta)
        Dim parametros() As Object = {"@idarticulo", "@idventa", "@preciounitario"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.BigInt, SqlDbType.Decimal}
        Dim valores() As Object = {d.idarticulo, d.idventa, d.preciounitario}
        sql.EjecutarProcedure("Str_tbl_Articulo_UnidadVenta_U", parametros, valores, tipoParametro, 3)
    End Sub
    Public Function Agregar(d As Ntbl_Articulo_UnidadVenta, Retornatable As Boolean) As Ntbl_Articulo_UnidadVenta

        Dim parametros() As Object = {"@idarticulo", "@idventa", "@preciounitario"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.BigInt, SqlDbType.Decimal}
        Dim valores() As Object = {d.idarticulo, d.idventa, d.preciounitario}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Articulo_UnidadVenta_I_S", parametros, valores, tipoParametro, 3).Tables(0)
        If dt.Rows.Count > 0 Then
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.idventa = IIf(dt.Rows(0).Item("idventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idventa"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
        Else
            d.item = Nothing
            d.idarticulo = Nothing
            d.idventa = Nothing
            d.preciounitario = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_Articulo_UnidadVenta, Retornatable As Boolean) As Ntbl_Articulo_UnidadVenta

        Dim parametros() As Object = {"@idarticulo", "@idventa", "@preciounitario"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.BigInt, SqlDbType.Decimal}
        Dim valores() As Object = {d.preciounitario = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Articulo_UnidadVenta_U_S", parametros, valores, tipoParametro, 11).Tables(0)
        If dt.Rows.Count > 0 Then
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.idventa = IIf(dt.Rows(0).Item("idventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idventa"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
        Else
            d.item = Nothing
            d.idarticulo = Nothing
            d.idventa = Nothing
            d.preciounitario = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_Articulo_UnidadVenta)
        Dim parametros() As Object = {"@idarticulo", "@idventa"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.BigInt}
        Dim valores() As Object = {d.idarticulo, d.idventa}
        sql.EjecutarProcedure("Str_tbl_Articulo_UnidadVenta_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Existe_tbl_Articulo_UnidadVenta(d As Ntbl_Articulo_UnidadVenta) As Boolean
        Dim parametros() As Object = {"@idarticulo", "@idventa"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.BigInt}
        Dim valores() As Object = {d.idarticulo, d.idventa}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_tbl_Articulo_UnidadVenta", parametros, valores, tipoParametro, 2)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idarticulo", "@idventa"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Articulo_UnidadVenta_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_Articulo_UnidadVenta) As DataTable
        Dim parametros() As Object = {"@idarticulo", "@idventa"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.BigInt}
        Dim valores() As Object = {d.idarticulo, d.idventa}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Articulo_UnidadVenta_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Ntbl_Articulo_UnidadVenta) As Ntbl_Articulo_UnidadVenta
        Dim parametros() As Object = {"@idarticulo", "@idventa"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.BigInt}
        Dim valores() As Object = {d.idarticulo, d.idventa}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Articulo_UnidadVenta_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.idventa = IIf(dt.Rows(0).Item("idventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idventa"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
        Else
            d.item = Nothing
            d.idarticulo = Nothing
            d.idventa = Nothing
            d.preciounitario = Nothing
        End If
        Return d
    End Function
#End Region


End Class
