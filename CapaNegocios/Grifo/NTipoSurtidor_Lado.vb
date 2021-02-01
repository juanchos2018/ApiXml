Imports CapaDatos
Public Class NTipoSurtidor_Lado
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property item As Integer
    Public Property lado As String
    Public Property transact As String
    Public Property ruta_dbf As String
    Public Property descripcion As String
    Public Property fechacrea As System.DateTime
    Public Property usuariocrea As String


#End Region

#Region "Constructors"
    Public Sub New()

    End Sub
#End Region
#Region "Metodos"
    Public Sub Agregar(d As NTipoSurtidor_Lado)

        Dim parametros() As Object = {"@item", "@lado", "@transact", "@ruta_dbf", "@descripcion", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.item, d.lado, d.transact, d.ruta_dbf, d.descripcion, d.fechacrea, d.usuariocrea}
        sql.EjecutarProcedure("Str_TipoSurtidor_Lado_I", parametros, valores, tipoParametro, 7)
    End Sub
    Public Function Agregar(d As NTipoSurtidor_Lado, Retornatable As Boolean) As NTipoSurtidor_Lado

        Dim parametros() As Object = {"@item", "@lado", "@transact", "@ruta_dbf", "@descripcion", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.item, d.lado, d.transact, d.ruta_dbf, d.descripcion, d.fechacrea, d.usuariocrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TipoSurtidor_Lado_I_S", parametros, valores, tipoParametro, 7).Tables(0)
        If dt.Rows.Count > 0 Then
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.lado = IIf(dt.Rows(0).Item("lado") Is DBNull.Value, Nothing, dt.Rows(0).Item("lado"))
            d.transact = IIf(dt.Rows(0).Item("transact") Is DBNull.Value, Nothing, dt.Rows(0).Item("transact"))
            d.ruta_dbf = IIf(dt.Rows(0).Item("ruta_dbf") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruta_dbf"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.item = Nothing
            d.lado = Nothing
            d.transact = Nothing
            d.ruta_dbf = Nothing
            d.descripcion = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Sub Actualizar(d As NTipoSurtidor_Lado)
        Dim parametros() As Object = {"@item", "@lado", "@transact", "@ruta_dbf", "@descripcion", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.usuariocrea = Nothing}
        sql.EjecutarProcedure("Str_TipoSurtidor_Lado_U", parametros, valores, tipoParametro, 7)
    End Sub
    Public Function Actualizar(d As NTipoSurtidor_Lado, Retornatable As Boolean) As NTipoSurtidor_Lado

        Dim parametros() As Object = {"@item", "@lado", "@transact", "@ruta_dbf", "@descripcion", "@fechacrea", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.usuariocrea = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TipoSurtidor_Lado_U_S", parametros, valores, tipoParametro, 7).Tables(0)
        If dt.Rows.Count > 0 Then
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.lado = IIf(dt.Rows(0).Item("lado") Is DBNull.Value, Nothing, dt.Rows(0).Item("lado"))
            d.transact = IIf(dt.Rows(0).Item("transact") Is DBNull.Value, Nothing, dt.Rows(0).Item("transact"))
            d.ruta_dbf = IIf(dt.Rows(0).Item("ruta_dbf") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruta_dbf"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.item = Nothing
            d.lado = Nothing
            d.transact = Nothing
            d.ruta_dbf = Nothing
            d.descripcion = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NTipoSurtidor_Lado)
        Dim parametros() As Object = {"@item", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar}
        Dim valores() As Object = {d.item, d.lado}
        sql.EjecutarProcedure("Str_TipoSurtidor_Lado_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@item", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TipoSurtidor_Lado_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTipoSurtidor_Lado) As DataTable
        Dim parametros() As Object = {"@item", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar}
        Dim valores() As Object = {d.item, d.lado}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TipoSurtidor_Lado_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTipoSurtidor_Lado) As NTipoSurtidor_Lado
        Dim parametros() As Object = {"@item", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar}
        Dim valores() As Object = {d.item, d.lado}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_TipoSurtidor_Lado_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.lado = IIf(dt.Rows(0).Item("lado") Is DBNull.Value, Nothing, dt.Rows(0).Item("lado"))
            d.transact = IIf(dt.Rows(0).Item("transact") Is DBNull.Value, Nothing, dt.Rows(0).Item("transact"))
            d.ruta_dbf = IIf(dt.Rows(0).Item("ruta_dbf") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruta_dbf"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.item = Nothing
            d.lado = Nothing
            d.transact = Nothing
            d.ruta_dbf = Nothing
            d.descripcion = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
#End Region

End Class
