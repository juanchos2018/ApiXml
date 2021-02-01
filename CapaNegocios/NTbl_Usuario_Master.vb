Imports CapaDatos
Public Class NTbl_Usuario_Master
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idusuariomenu As String
    Public Property idmenu As String
    Public Property descripciomenu As String
    Public Property idusuariocrea As String
    Public Property fechacrea As System.DateTime
    Public Property aplicacion As String
#End Region
#Region "Constructors"
    Public Sub New()
    End Sub
#End Region
#Region "Metodos"
    Public Sub Agregar(d As NTbl_Usuario_Master)

        Dim parametros() As Object = {"@idusuariomenu", "@idmenu", "@descripciomenu", "@idusuariocrea", "@fechacrea", "@aplicacion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idusuariomenu, d.idmenu, d.descripciomenu, d.idusuariocrea, d.fechacrea, d.aplicacion}
        Sql.EjecutarProcedure("Str_Tbl_Usuario_Master_I", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Actualizar(d As NTbl_Usuario_Master)
        Dim parametros() As Object = {"@idusuariomenu", "@idmenu", "@descripciomenu", "@idusuariocrea", "@fechacrea", "@aplicacion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idusuariomenu, d.idmenu, d.descripciomenu, d.idusuariocrea, d.fechacrea, d.aplicacion}
        Sql.EjecutarProcedure("Str_Tbl_Usuario_Master_U", parametros, valores, tipoParametro, 6)
    End Sub
    Public Function Agregar(d As NTbl_Usuario_Master, Retornatable As Boolean) As NTbl_Usuario_Master

        Dim parametros() As Object = {"@idusuariomenu", "@idmenu", "@descripciomenu", "@idusuariocrea", "@fechacrea", "@aplicacion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idusuariomenu, d.idmenu, d.descripciomenu, d.idusuariocrea, d.fechacrea, d.aplicacion}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_Tbl_Usuario_Master_I_S", parametros, valores, tipoParametro, 6).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idusuariomenu = IIf(dt.Rows(0).Item("idusuariomenu") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariomenu"))
            d.idmenu = IIf(dt.Rows(0).Item("idmenu") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmenu"))
            d.descripciomenu = IIf(dt.Rows(0).Item("descripciomenu") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripciomenu"))
            d.idusuariocrea = IIf(dt.Rows(0).Item("idusuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.aplicacion = IIf(dt.Rows(0).Item("aplicacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("aplicacion"))
        Else
            d.idusuariomenu = Nothing
            d.idmenu = Nothing
            d.descripciomenu = Nothing
            d.idusuariocrea = Nothing
            d.fechacrea = Nothing
            d.aplicacion = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NTbl_Usuario_Master, Retornatable As Boolean) As NTbl_Usuario_Master

        Dim parametros() As Object = {"@idusuariomenu", "@idmenu", "@descripciomenu", "@idusuariocrea", "@fechacrea", "@aplicacion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.aplicacion = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Usuario_Master_U_S", parametros, valores, tipoParametro, 6).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idusuariomenu = IIf(dt.Rows(0).Item("idusuariomenu") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariomenu"))
            d.idmenu = IIf(dt.Rows(0).Item("idmenu") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmenu"))
            d.descripciomenu = IIf(dt.Rows(0).Item("descripciomenu") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripciomenu"))
            d.idusuariocrea = IIf(dt.Rows(0).Item("idusuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.aplicacion = IIf(dt.Rows(0).Item("aplicacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("aplicacion"))
        Else
            d.idusuariomenu = Nothing
            d.idmenu = Nothing
            d.descripciomenu = Nothing
            d.idusuariocrea = Nothing
            d.fechacrea = Nothing
            d.aplicacion = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NTbl_Usuario_Master)
        Dim parametros() As Object = {"@idusuariomenu", "@idmenu"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idusuariomenu, d.idmenu}
        Sql.EjecutarProcedure("Str_Tbl_Usuario_Master_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idusuariomenu", "@idmenu"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_Tbl_Usuario_Master_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTbl_Usuario_Master) As DataTable
        Dim parametros() As Object = {"@idusuariomenu", "@idmenu"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idusuariomenu, d.idmenu}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_Tbl_Usuario_Master_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTbl_Usuario_Master) As NTbl_Usuario_Master
        Dim parametros() As Object = {"@idusuariomenu", "@idmenu"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idusuariomenu, d.idmenu}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_Tbl_Usuario_Master_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idusuariomenu = IIf(dt.Rows(0).Item("idusuariomenu") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariomenu"))
            d.idmenu = IIf(dt.Rows(0).Item("idmenu") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmenu"))
            d.descripciomenu = IIf(dt.Rows(0).Item("descripciomenu") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripciomenu"))
            d.idusuariocrea = IIf(dt.Rows(0).Item("idusuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.aplicacion = IIf(dt.Rows(0).Item("aplicacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("aplicacion"))
        Else
            d.idusuariomenu = Nothing
            d.idmenu = Nothing
            d.descripciomenu = Nothing
            d.idusuariocrea = Nothing
            d.fechacrea = Nothing
            d.aplicacion = Nothing
        End If
        Return d
    End Function
#End Region

End Class
