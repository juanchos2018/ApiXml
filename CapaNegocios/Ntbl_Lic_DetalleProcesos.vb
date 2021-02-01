Imports CapaDatos
Public Class Ntbl_Lic_DetalleProcesos
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property iddetproceso As Long
    Public Property item As Integer
    Public Property idarticulo As String
    Public Property descripcion As String
    Public Property cantidad As Decimal
    Public Property preciounitario As Decimal
    Public Property total As Decimal
    Public Property idproceso As Long

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As Ntbl_Lic_DetalleProcesos)
        Dim parametros() As Object = {"@item", "@idarticulo", "@descripcion", "@cantidad", "@preciounitario", "@total", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Char, SqlDbType.NVarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.BigInt}
        Dim valores() As Object = {d.item, d.idarticulo, d.descripcion, d.cantidad, d.preciounitario, d.total, d.idproceso}
        sql.EjecutarProcedure("Str_tbl_Lic_DetalleProcesos_I", parametros, valores, tipoParametro, 7)
    End Sub
    Public Sub Actualizar(d As Ntbl_Lic_DetalleProcesos)
        Dim parametros() As Object = {"@iddetproceso", "@item", "@idarticulo", "@descripcion", "@cantidad", "@preciounitario", "@total", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.Int, SqlDbType.Char, SqlDbType.NVarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.BigInt}
        Dim valores() As Object = {d.iddetproceso, d.item, d.idarticulo, d.descripcion, d.cantidad, d.preciounitario, d.total, d.idproceso}
        sql.EjecutarProcedure("Str_tbl_Lic_DetalleProcesos_U", parametros, valores, tipoParametro, 8)
    End Sub
    Public Function Agregar(d As Ntbl_Lic_DetalleProcesos, Retornatable As Boolean) As Ntbl_Lic_DetalleProcesos
        Dim parametros() As Object = {"@item", "@idarticulo", "@descripcion", "@cantidad", "@preciounitario", "@total", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Char, SqlDbType.NVarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.BigInt}
        Dim valores() As Object = {d.item, d.idarticulo, d.descripcion, d.cantidad, d.preciounitario, d.total, d.idproceso}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_DetalleProcesos_I_S", parametros, valores, tipoParametro, 7).Tables(0)
        If dt.Rows.Count > 0 Then
            d.iddetproceso = IIf(dt.Rows(0).Item("iddetproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetproceso"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
            d.total = IIf(dt.Rows(0).Item("total") Is DBNull.Value, Nothing, dt.Rows(0).Item("total"))
            d.idproceso = IIf(dt.Rows(0).Item("idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproceso"))
        Else
            d.iddetproceso = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
            d.cantidad = Nothing
            d.preciounitario = Nothing
            d.total = Nothing
            d.idproceso = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_Lic_DetalleProcesos, Retornatable As Boolean) As Ntbl_Lic_DetalleProcesos
        Dim parametros() As Object = {"@iddetproceso", "@item", "@idarticulo", "@descripcion", "@cantidad", "@preciounitario", "@total", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.Int, SqlDbType.Char, SqlDbType.NVarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.BigInt}
        Dim valores() As Object = {d.iddetproceso, d.item, d.idarticulo, d.descripcion, d.cantidad, d.preciounitario, d.total, d.idproceso}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_DetalleProcesos_U_S", parametros, valores, tipoParametro, 8).Tables(0)
        If dt.Rows.Count > 0 Then
            d.iddetproceso = IIf(dt.Rows(0).Item("iddetproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetproceso"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
            d.total = IIf(dt.Rows(0).Item("total") Is DBNull.Value, Nothing, dt.Rows(0).Item("total"))
            d.idproceso = IIf(dt.Rows(0).Item("idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproceso"))
        Else
            d.iddetproceso = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
            d.cantidad = Nothing
            d.preciounitario = Nothing
            d.total = Nothing
            d.idproceso = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_Lic_DetalleProcesos)
        Dim parametros() As Object = {"@iddetproceso", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.BigInt}
        Dim valores() As Object = {d.iddetproceso, d.idproceso}
        sql.EjecutarProcedure("Str_tbl_Lic_DetalleProcesos_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Existe_tbl_Lic_DetalleProcesos(d As Ntbl_Lic_DetalleProcesos)
        Dim parametros() As Object = {"@iddetproceso", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.BigInt}
        Dim valores() As Object = {d.iddetproceso, d.idproceso}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_tbl_Lic_DetalleProcesos", parametros, valores, tipoParametro, 2)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@iddetproceso", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_DetalleProcesos_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_Lic_DetalleProcesos) As DataTable
        Dim parametros() As Object = {"@iddetproceso", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.BigInt}
        Dim valores() As Object = {d.iddetproceso, d.idproceso}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_DetalleProcesos_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Ntbl_Lic_DetalleProcesos) As Ntbl_Lic_DetalleProcesos
        Dim parametros() As Object = {"@iddetproceso", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.BigInt}
        Dim valores() As Object = {d.iddetproceso, d.idproceso}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_DetalleProcesos_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.iddetproceso = IIf(dt.Rows(0).Item("iddetproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetproceso"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.preciounitario = IIf(dt.Rows(0).Item("preciounitario") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitario"))
            d.total = IIf(dt.Rows(0).Item("total") Is DBNull.Value, Nothing, dt.Rows(0).Item("total"))
            d.idproceso = IIf(dt.Rows(0).Item("idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproceso"))
        Else
            d.iddetproceso = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
            d.cantidad = Nothing
            d.preciounitario = Nothing
            d.total = Nothing
            d.idproceso = Nothing
        End If
        Return d
    End Function
    Public Function Lista_Detalle_Procesos(idproceso As Long) As DataTable
        Dim sentenciaSQL As String = " select IdArticulo,Descripcion from tbl_Lic_DetalleProcesos where IdProceso=" & idproceso & ""
        Return Me.sql.EjecutarConsulta("s", sentenciaSQL).Tables(0)
    End Function

#End Region


End Class
