Imports CapaDatos
Public Class NUsuario
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idusuario As String
    Public Property idnivelusuario As String
    Public Property [alias] As String
    Public Property password As String
    Public Property idestado As String
    Public Property swcaja As Boolean
    Public Property swturno As Boolean
    Public Property fullname As String
    Public Property swalmacen As Boolean
    Public Property islogin_sql As Boolean

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NUsuario)

        Dim parametros() As Object = {"@idusuario", "@idnivelusuario", "@alias", "@password", "@idestado", "@swcaja", "@swturno", "@fullname", "@swalmacen", "@islogin_sql"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.idusuario, d.idnivelusuario, d.alias, d.password, d.idestado, d.swcaja, d.swturno, d.fullname, d.swalmacen, d.islogin_sql}
        sql.EjecutarProcedure("Str_ptusuario_I", parametros, valores, tipoParametro, 10)
    End Sub
    Public Sub Actualizar(d As NUsuario)
        Dim parametros() As Object = {"@idusuario", "@idnivelusuario", "@alias", "@password", "@idestado", "@swcaja", "@swturno", "@fullname", "@swalmacen", "@islogin_sql"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.idusuario, d.idnivelusuario, d.alias, d.password, d.idestado, d.swcaja, d.swturno, d.fullname, d.swalmacen, d.islogin_sql}
        sql.EjecutarProcedure("Str_ptusuario_U", parametros, valores, tipoParametro, 10)
    End Sub
    Public Function Agregar(d As NUsuario, Retornatable As Boolean) As NUsuario

        Dim parametros() As Object = {"@idusuario", "@idnivelusuario", "@alias", "@password", "@idestado", "@swcaja", "@swturno", "@fullname", "@swalmacen", "@islogin_sql"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.idusuario, d.idnivelusuario, d.alias, d.password, d.idestado, d.swcaja, d.swturno, d.fullname, d.swalmacen, d.islogin_sql}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ptusuario_I_S", parametros, valores, tipoParametro, 10).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idusuario = IIf(dt.Rows(0).Item("idusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuario"))
            d.idnivelusuario = IIf(dt.Rows(0).Item("idnivelusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idnivelusuario"))
            d.alias = IIf(dt.Rows(0).Item("alias") Is DBNull.Value, Nothing, dt.Rows(0).Item("alias"))
            d.password = IIf(dt.Rows(0).Item("password") Is DBNull.Value, Nothing, dt.Rows(0).Item("password"))
            d.idestado = IIf(dt.Rows(0).Item("idestado") Is DBNull.Value, Nothing, dt.Rows(0).Item("idestado"))
            d.swcaja = IIf(dt.Rows(0).Item("swcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("swcaja"))
            d.swturno = IIf(dt.Rows(0).Item("swturno") Is DBNull.Value, Nothing, dt.Rows(0).Item("swturno"))
            d.fullname = IIf(dt.Rows(0).Item("fullname") Is DBNull.Value, Nothing, dt.Rows(0).Item("fullname"))
            d.swalmacen = IIf(dt.Rows(0).Item("swalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("swalmacen"))
            d.islogin_sql = IIf(dt.Rows(0).Item("islogin_sql") Is DBNull.Value, Nothing, dt.Rows(0).Item("islogin_sql"))
        Else
            d.idusuario = Nothing
            d.idnivelusuario = Nothing
            d.alias = Nothing
            d.password = Nothing
            d.idestado = Nothing
            d.swcaja = Nothing
            d.swturno = Nothing
            d.fullname = Nothing
            d.swalmacen = Nothing
            d.islogin_sql = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NUsuario, Retornatable As Boolean) As NUsuario

        Dim parametros() As Object = {"@idusuario", "@idnivelusuario", "@alias", "@password", "@idestado", "@swcaja", "@swturno", "@fullname", "@swalmacen", "@islogin_sql"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.idusuario, d.idnivelusuario, d.alias, d.password, d.idestado, d.swcaja, d.swturno, d.fullname, d.swalmacen, d.islogin_sql}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ptusuario_U_S", parametros, valores, tipoParametro, 10).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idusuario = IIf(dt.Rows(0).Item("idusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuario"))
            d.idnivelusuario = IIf(dt.Rows(0).Item("idnivelusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idnivelusuario"))
            d.alias = IIf(dt.Rows(0).Item("alias") Is DBNull.Value, Nothing, dt.Rows(0).Item("alias"))
            d.password = IIf(dt.Rows(0).Item("password") Is DBNull.Value, Nothing, dt.Rows(0).Item("password"))
            d.idestado = IIf(dt.Rows(0).Item("idestado") Is DBNull.Value, Nothing, dt.Rows(0).Item("idestado"))
            d.swcaja = IIf(dt.Rows(0).Item("swcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("swcaja"))
            d.swturno = IIf(dt.Rows(0).Item("swturno") Is DBNull.Value, Nothing, dt.Rows(0).Item("swturno"))
            d.fullname = IIf(dt.Rows(0).Item("fullname") Is DBNull.Value, Nothing, dt.Rows(0).Item("fullname"))
            d.swalmacen = IIf(dt.Rows(0).Item("swalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("swalmacen"))
            d.islogin_sql = IIf(dt.Rows(0).Item("islogin_sql") Is DBNull.Value, Nothing, dt.Rows(0).Item("islogin_sql"))
        Else
            d.idusuario = Nothing
            d.idnivelusuario = Nothing
            d.alias = Nothing
            d.password = Nothing
            d.idestado = Nothing
            d.swcaja = Nothing
            d.swturno = Nothing
            d.fullname = Nothing
            d.swalmacen = Nothing
            d.islogin_sql = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NUsuario)
        Dim parametros() As Object = {"@idusuario"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.idusuario}
        sql.EjecutarProcedure("Str_ptusuario_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idusuario"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ptusuario_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NUsuario) As DataTable
        Dim parametros() As Object = {"@idusuario"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.idusuario}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ptusuario_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NUsuario) As NUsuario
        Dim parametros() As Object = {"@idusuario"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.idusuario}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ptusuario_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idusuario = IIf(dt.Rows(0).Item("idusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuario"))
            d.idnivelusuario = IIf(dt.Rows(0).Item("idnivelusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idnivelusuario"))
            d.alias = IIf(dt.Rows(0).Item("alias") Is DBNull.Value, Nothing, dt.Rows(0).Item("alias"))
            d.password = IIf(dt.Rows(0).Item("password") Is DBNull.Value, Nothing, dt.Rows(0).Item("password"))
            d.idestado = IIf(dt.Rows(0).Item("idestado") Is DBNull.Value, Nothing, dt.Rows(0).Item("idestado"))
            d.swcaja = IIf(dt.Rows(0).Item("swcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("swcaja"))
            d.swturno = IIf(dt.Rows(0).Item("swturno") Is DBNull.Value, Nothing, dt.Rows(0).Item("swturno"))
            d.fullname = IIf(dt.Rows(0).Item("fullname") Is DBNull.Value, Nothing, dt.Rows(0).Item("fullname"))
            d.swalmacen = IIf(dt.Rows(0).Item("swalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("swalmacen"))
            d.islogin_sql = IIf(dt.Rows(0).Item("islogin_sql") Is DBNull.Value, Nothing, dt.Rows(0).Item("islogin_sql"))
        Else
            d.idusuario = Nothing
            d.idnivelusuario = Nothing
            d.alias = Nothing
            d.password = Nothing
            d.idestado = Nothing
            d.swcaja = Nothing
            d.swturno = Nothing
            d.fullname = Nothing
            d.swalmacen = Nothing
            d.islogin_sql = Nothing
        End If
        Return d
    End Function
#End Region

End Class
