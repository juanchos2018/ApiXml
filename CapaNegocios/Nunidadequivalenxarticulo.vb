Imports CapaDatos
Public Class Nunidadequivalenxarticulo
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property id As Long
    Public Property idarticulo As String
    Public Property unidadorigen As String
    Public Property unidaddestino As String
    Public Property factorconversion As Decimal
    Public Property medida As String
#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As Nunidadequivalenxarticulo)
        Dim parametros() As Object = {"@idarticulo", "@unidadorigen", "@unidaddestino", "@factorconversion", "@medida"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.idarticulo, d.unidadorigen, d.unidaddestino, d.factorconversion, d.medida}
        sql.EjecutarProcedure("Str_Unidadequivalenxarticulo_I", parametros, valores, tipoParametro, 5)
    End Sub
    Public Sub Actualizar(d As Nunidadequivalenxarticulo)
        Dim parametros() As Object = {"@id", "@idarticulo", "@unidadorigen", "@unidaddestino", "@factorconversion", "@medida"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.id, d.idarticulo, d.unidadorigen, d.unidaddestino, d.factorconversion, d.medida}
        sql.EjecutarProcedure("Str_Unidadequivalenxarticulo_U", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Eliminar(d As Nunidadequivalenxarticulo)
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        sql.EjecutarProcedure("Str_Unidadequivalenxarticulo_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Unidadequivalenxarticulo_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Nunidadequivalenxarticulo) As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Unidadequivalenxarticulo_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Nunidadequivalenxarticulo) As Nunidadequivalenxarticulo
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Unidadequivalenxarticulo_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.unidadorigen = IIf(dt.Rows(0).Item("unidadorigen") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidadorigen"))
            d.unidaddestino = IIf(dt.Rows(0).Item("unidaddestino") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidaddestino"))
            d.factorconversion = IIf(dt.Rows(0).Item("factorconversion") Is DBNull.Value, Nothing, dt.Rows(0).Item("factorconversion"))
            d.medida = IIf(dt.Rows(0).Item("medida") Is DBNull.Value, Nothing, dt.Rows(0).Item("medida"))
        Else
            d.id = Nothing
            d.idarticulo = Nothing
            d.unidadorigen = Nothing
            d.unidaddestino = Nothing
            d.factorconversion = Nothing
            d.medida = Nothing
        End If
        Return d
    End Function
    Public Function lista_und_vta() As DataTable
        Dim s As String = " select cast(0 as bit) as Ok,id,unidadorigen,unidaddestino,factorconversion,Medida,0.00 as Precio from  unidadequivalenxarticulo "
        Return sql.EjecutarConsulta("xf", s).Tables(0)
    End Function

#End Region



End Class
