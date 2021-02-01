Imports CapaDatos
Public Class Ntbl_Conductores
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idconductor As String
    Public Property nombreconductor As String
    Public Property dni As String
    Public Property estado As Boolean

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region


#Region "Metodos"
    Public Sub Agregar(d As Ntbl_Conductores)

        Dim parametros() As Object = {"@idconductor", "@nombreconductor", "@dni", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.idconductor, d.nombreconductor, d.dni, d.estado}
        sql.EjecutarProcedure("Str_tbl_Conductores_I", parametros, valores, tipoParametro, 4)
    End Sub
    Public Sub Actualizar(d As Ntbl_Conductores)
        Dim parametros() As Object = {"@idconductor", "@nombreconductor", "@dni", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.idconductor, d.nombreconductor, d.dni, d.estado}
        sql.EjecutarProcedure("Str_tbl_Conductores_U", parametros, valores, tipoParametro, 4)
    End Sub
    Public Function Agregar(d As Ntbl_Conductores, Retornatable As Boolean) As Ntbl_Conductores

        Dim parametros() As Object = {"@idconductor", "@nombreconductor", "@dni", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.idconductor, d.nombreconductor, d.dni, d.estado}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Conductores_I_S", parametros, valores, tipoParametro, 4).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idconductor = IIf(dt.Rows(0).Item("idconductor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idconductor"))
            d.nombreconductor = IIf(dt.Rows(0).Item("nombreconductor") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreconductor"))
            d.dni = IIf(dt.Rows(0).Item("dni") Is DBNull.Value, Nothing, dt.Rows(0).Item("dni"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
        Else
            d.idconductor = Nothing
            d.nombreconductor = Nothing
            d.dni = Nothing
            d.estado = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_Conductores, Retornatable As Boolean) As Ntbl_Conductores

        Dim parametros() As Object = {"@idconductor", "@nombreconductor", "@dni", "@estado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.estado = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Conductores_U_S", parametros, valores, tipoParametro, 12).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idconductor = IIf(dt.Rows(0).Item("idconductor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idconductor"))
            d.nombreconductor = IIf(dt.Rows(0).Item("nombreconductor") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreconductor"))
            d.dni = IIf(dt.Rows(0).Item("dni") Is DBNull.Value, Nothing, dt.Rows(0).Item("dni"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
        Else
            d.idconductor = Nothing
            d.nombreconductor = Nothing
            d.dni = Nothing
            d.estado = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_Conductores)
        Dim parametros() As Object = {"@idconductor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idconductor}
        sql.EjecutarProcedure("Str_tbl_Conductores_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_tbl_Conductores(d As Ntbl_Conductores) As Boolean
        Dim parametros() As Object = {"@idconductor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idconductor}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_tbl_Conductores", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idconductor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Conductores_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_Conductores) As DataTable
        Dim parametros() As Object = {"@idconductor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idconductor}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Conductores_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Ntbl_Conductores) As Ntbl_Conductores
        Dim parametros() As Object = {"@idconductor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idconductor}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Conductores_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idconductor = IIf(dt.Rows(0).Item("idconductor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idconductor"))
            d.nombreconductor = IIf(dt.Rows(0).Item("nombreconductor") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreconductor"))
            d.dni = IIf(dt.Rows(0).Item("dni") Is DBNull.Value, Nothing, dt.Rows(0).Item("dni"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
        Else
            d.idconductor = Nothing
            d.nombreconductor = Nothing
            d.dni = Nothing
            d.estado = Nothing
        End If
        Return d
    End Function
#End Region

End Class
