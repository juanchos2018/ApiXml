Imports CapaDatos
Public Class Narticulo_equivalente
    Dim sql As New ClsConexion

#Region "Declarations"

    Public Property idarticuloe As String
    Public Property idarticulo As String
    Public Property descripcion As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As Narticulo_equivalente)

        Dim parametros() As Object = {"@idarticuloe", "@idarticulo", "@descripcion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idarticuloe, d.idarticulo, d.descripcion}
        sql.EjecutarProcedure("Str_articulo_equivalente_I", parametros, valores, tipoParametro, 3)
    End Sub
    Public Sub Actualizar(d As Narticulo_equivalente)
        Dim parametros() As Object = {"@idarticuloe", "@idarticulo", "@descripcion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idarticuloe, d.idarticulo, d.descripcion}
        sql.EjecutarProcedure("Str_articulo_equivalente_U", parametros, valores, tipoParametro, 3)
    End Sub
    Public Function Agregar(d As Narticulo_equivalente, Retornatable As Boolean) As Narticulo_equivalente

        Dim parametros() As Object = {"@idarticuloe", "@idarticulo", "@descripcion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idarticuloe, d.idarticulo, d.descripcion}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_articulo_equivalente_I_S", parametros, valores, tipoParametro, 3).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idarticuloe = IIf(dt.Rows(0).Item("idarticuloe") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticuloe"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
        Else
            d.idarticuloe = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Narticulo_equivalente, Retornatable As Boolean) As Narticulo_equivalente

        Dim parametros() As Object = {"@idarticuloe", "@idarticulo", "@descripcion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.descripcion = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_articulo_equivalente_U_S", parametros, valores, tipoParametro, 9).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idarticuloe = IIf(dt.Rows(0).Item("idarticuloe") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticuloe"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
        Else
            d.idarticuloe = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As Narticulo_equivalente)
        Dim parametros() As Object = {"@idarticuloe"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idarticuloe}
        sql.EjecutarProcedure("Str_articulo_equivalente_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_articulo_equivalente(d As Narticulo_equivalente) As Boolean
        Dim parametros() As Object = {"@idarticuloe"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idarticuloe}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_articulo_equivalente", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idarticuloe"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_articulo_equivalente_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Narticulo_equivalente) As DataTable
        Dim parametros() As Object = {"@idarticuloe"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idarticuloe}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_articulo_equivalente_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function ListaPreciosEquivalente() As DataTable
        Dim parametros() As Object = {"@idarticuloe"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.Proc_DataReader("Str_PrecioArticulo", parametros, valores, tipoParametro, 0)
        Return dt
    End Function

    Public Function Registro(d As Narticulo_equivalente) As Narticulo_equivalente
        Dim parametros() As Object = {"@idarticuloe"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idarticuloe}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_articulo_equivalente_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idarticuloe = IIf(dt.Rows(0).Item("idarticuloe") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticuloe"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
        Else
            d.idarticuloe = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
        End If
        Return d
    End Function
#End Region


End Class
