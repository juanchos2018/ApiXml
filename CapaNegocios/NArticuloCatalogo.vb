Imports CapaDatos

Public Class NArticuloCatalogo
    Dim sql As New ClsConexion

#Region "Declarations"

    Public Property idarticulo As String
    Public Property moneda As String
    Public Property importe As Decimal
    Public Property importemn As Decimal
    Public Property importeus As Decimal
    Public Property item As Integer
    Public Property descripcion As String
    Public Property idtipoprecio As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NArticuloCatalogo)

        Dim parametros() As Object = {"@idarticulo", "@moneda", "@importe", "@importemn", "@importeus", "@descripcion", "@idtipoprecio"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idarticulo, d.moneda, d.importe, d.importemn, d.importeus, d.descripcion, d.idtipoprecio}
        sql.EjecutarProcedure("Str_ArticuloCatalogo_I", parametros, valores, tipoParametro, 7)
    End Sub
    Public Sub Actualizar(d As NArticuloCatalogo)
        Dim parametros() As Object = {"@item", "@idarticulo", "@moneda", "@importe", "@importemn", "@importeus", "@descripcion", "@idtipoprecio"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Char, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.item, d.idarticulo, d.moneda, d.importe, d.importemn, d.importeus, d.descripcion, d.idtipoprecio}
        sql.EjecutarProcedure("Str_ArticuloCatalogo_U", parametros, valores, tipoParametro, 8)
    End Sub
    Public Function Agregar(d As NArticuloCatalogo, Retornatable As Boolean) As NArticuloCatalogo

        Dim parametros() As Object = {"@idarticulo", "@moneda", "@importe", "@importemn", "@importeus", "@descripcion", "@idtipoprecio"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idarticulo, d.moneda, d.importe, d.importemn, d.importeus, d.descripcion, d.idtipoprecio}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ArticuloCatalogo_I_S", parametros, valores, tipoParametro, 7).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.moneda = IIf(dt.Rows(0).Item("moneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("moneda"))
            d.importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.idtipoprecio = IIf(dt.Rows(0).Item("idtipoprecio") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipoprecio"))
        Else
            d.idarticulo = Nothing
            d.moneda = Nothing
            d.importe = Nothing
            d.importemn = Nothing
            d.importeus = Nothing
            d.item = Nothing
            d.descripcion = Nothing
            d.idtipoprecio = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NArticuloCatalogo, Retornatable As Boolean) As NArticuloCatalogo
        Dim parametros() As Object = {"@item", "@idarticulo", "@moneda", "@importe", "@importemn", "@importeus", "@descripcion", "@idtipoprecio"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Char, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.item, d.idarticulo, d.moneda, d.importe, d.importemn, d.importeus, d.descripcion, d.idtipoprecio}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ArticuloCatalogo_U_S", parametros, valores, tipoParametro, 8).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.moneda = IIf(dt.Rows(0).Item("moneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("moneda"))
            d.importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.idtipoprecio = IIf(dt.Rows(0).Item("idtipoprecio") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipoprecio"))
        Else
            d.idarticulo = Nothing
            d.moneda = Nothing
            d.importe = Nothing
            d.importemn = Nothing
            d.importeus = Nothing
            d.item = Nothing
            d.descripcion = Nothing
            d.idtipoprecio = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NArticuloCatalogo)
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.item}
        sql.EjecutarProcedure("Str_ArticuloCatalogo_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ArticuloCatalogo_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NArticuloCatalogo) As DataTable
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ArticuloCatalogo_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NArticuloCatalogo) As NArticuloCatalogo
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ArticuloCatalogo_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.moneda = IIf(dt.Rows(0).Item("moneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("moneda"))
            d.importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.idtipoprecio = IIf(dt.Rows(0).Item("idtipoprecio") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipoprecio"))
        Else
            d.idarticulo = Nothing
            d.moneda = Nothing
            d.importe = Nothing
            d.importemn = Nothing
            d.importeus = Nothing
            d.item = Nothing
            d.descripcion = Nothing
            d.idtipoprecio = Nothing
        End If
        Return d
    End Function

    Public Function ListaArticulo(d As NArticuloCatalogo) As DataTable
        Dim txt$
        txt = " SELECT Item,descripcion,Moneda,Importe,IdTipoPrecio FROM ArticuloCatalogo "
        txt += " where IdArticulo='" & d.IdArticulo & "'"
        Dim dtx As New DataTable
        dtx = sql.EjecutarConsulta("se", txt).Tables(0)
        Return dtx
    End Function

#End Region
End Class
