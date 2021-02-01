Imports CapaDatos
Public Class Ntbl_DetalleDepachosYura
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idticket As String
    Public Property item As String
    Public Property idarticulo As String
    Public Property cantidad As Decimal
    Public Property pesototal As Decimal
    Public Property pedido_sap As String
    Public Property iddetoc As Long
    Public Property estado As Boolean

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As Ntbl_DetalleDepachosYura)

        Dim parametros() As Object = {"@idticket", "@item", "@idarticulo", "@cantidad", "@pesototal", "@pedido_sap", "@iddetoc", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.BigInt, SqlDbType.Bit}
        Dim valores() As Object = {d.idticket, d.item, d.idarticulo, d.cantidad, d.pesototal, d.pedido_sap, d.iddetoc, d.estado}
        sql.EjecutarProcedure("Str_tbl_DetalleDepachosYura_I", parametros, valores, tipoParametro, 8)
    End Sub
    Public Sub Actualizar(d As Ntbl_DetalleDepachosYura)
        Dim parametros() As Object = {"@idticket", "@item", "@idarticulo", "@cantidad", "@pesototal", "@pedido_sap", "@iddetoc", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.BigInt, SqlDbType.Bit}
        Dim valores() As Object = {d.idticket, d.item, d.idarticulo, d.cantidad, d.pesototal, d.pedido_sap, d.iddetoc, d.estado}
        sql.EjecutarProcedure("Str_tbl_DetalleDepachosYura_U", parametros, valores, tipoParametro, 8)
    End Sub
    Public Function Agregar(d As Ntbl_DetalleDepachosYura, Retornatable As Boolean) As Ntbl_DetalleDepachosYura

        Dim parametros() As Object = {"@idticket", "@item", "@idarticulo", "@cantidad", "@pesototal", "@pedido_sap", "@iddetoc", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.BigInt, SqlDbType.Bit}
        Dim valores() As Object = {d.idticket, d.item, d.idarticulo, d.cantidad, d.pesototal, d.pedido_sap, d.iddetoc, d.estado}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DetalleDepachosYura_I_S", parametros, valores, tipoParametro, 8).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idticket = IIf(dt.Rows(0).Item("idticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("idticket"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.pesototal = IIf(dt.Rows(0).Item("pesototal") Is DBNull.Value, Nothing, dt.Rows(0).Item("pesototal"))
            d.pedido_sap = IIf(dt.Rows(0).Item("pedido_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("pedido_sap"))
            d.iddetoc = IIf(dt.Rows(0).Item("iddetoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetoc"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
        Else
            d.idticket = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.cantidad = Nothing
            d.pesototal = Nothing
            d.pedido_sap = Nothing
            d.iddetoc = Nothing
            d.estado = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_DetalleDepachosYura, Retornatable As Boolean) As Ntbl_DetalleDepachosYura

        Dim parametros() As Object = {"@idticket", "@item", "@idarticulo", "@cantidad", "@pesototal", "@pedido_sap", "@iddetoc", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.BigInt, SqlDbType.Bit}
        Dim valores() As Object = {d.estado = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DetalleDepachosYura_U_S", parametros, valores, tipoParametro, 24).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idticket = IIf(dt.Rows(0).Item("idticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("idticket"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.pesototal = IIf(dt.Rows(0).Item("pesototal") Is DBNull.Value, Nothing, dt.Rows(0).Item("pesototal"))
            d.pedido_sap = IIf(dt.Rows(0).Item("pedido_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("pedido_sap"))
            d.iddetoc = IIf(dt.Rows(0).Item("iddetoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetoc"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
        Else
            d.idticket = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.cantidad = Nothing
            d.pesototal = Nothing
            d.pedido_sap = Nothing
            d.iddetoc = Nothing
            d.estado = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_DetalleDepachosYura)
        Dim parametros() As Object = {"@idticket", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idticket, d.item}
        sql.EjecutarProcedure("Str_tbl_DetalleDepachosYura_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Existe_tbl_DetalleDepachosYura(d As Ntbl_DetalleDepachosYura) As Boolean
        Dim parametros() As Object = {"@idticket", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idticket, d.item}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_tbl_DetalleDepachosYura", parametros, valores, tipoParametro, 2)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idticket", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DetalleDepachosYura_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_DetalleDepachosYura) As DataTable
        Dim parametros() As Object = {"@idticket", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idticket, d.item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DetalleDepachosYura_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Ntbl_DetalleDepachosYura) As Ntbl_DetalleDepachosYura
        Dim parametros() As Object = {"@idticket", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idticket, d.item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DetalleDepachosYura_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idticket = IIf(dt.Rows(0).Item("idticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("idticket"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.pesototal = IIf(dt.Rows(0).Item("pesototal") Is DBNull.Value, Nothing, dt.Rows(0).Item("pesototal"))
            d.pedido_sap = IIf(dt.Rows(0).Item("pedido_sap") Is DBNull.Value, Nothing, dt.Rows(0).Item("pedido_sap"))
            d.iddetoc = IIf(dt.Rows(0).Item("iddetoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetoc"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
        Else
            d.idticket = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.cantidad = Nothing
            d.pesototal = Nothing
            d.pedido_sap = Nothing
            d.iddetoc = Nothing
            d.estado = Nothing
        End If
        Return d
    End Function
#End Region

End Class
