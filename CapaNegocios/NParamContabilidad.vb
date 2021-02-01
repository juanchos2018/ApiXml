Imports CapaDatos
Public Class NParamContabilidad
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idrucsiscorp As String
    Public Property nombresiscorp As String
    Public Property flagcli As Boolean
    Public Property flagventas As Boolean
    Public Property flagcompras As Boolean
    Public Property flagprov As Boolean
    Public Property periodo As String
    Public Property flagventasdetalle As Boolean
    Public Property flagcomprasdetalle As Boolean
    Public Property flagpagos As Boolean
    Public Property flagcobros As Boolean
    Public Property flagcobconta As Boolean

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NParamContabilidad)

        Dim parametros() As Object = {"@idrucsiscorp", "@nombresiscorp", "@flagcli", "@flagventas", "@flagcompras", "@flagprov", "@periodo", "@flagventasdetalle", "@flagcomprasdetalle", "@flagpagos", "@flagcobros", "@flagcobconta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.idrucsiscorp, d.nombresiscorp, d.flagcli, d.flagventas, d.flagcompras, d.flagprov, d.periodo, d.flagventasdetalle, d.flagcomprasdetalle, d.flagpagos, d.flagcobros, d.flagcobconta}
        sql.EjecutarProcedure("Str_ParamContabilidad_I", parametros, valores, tipoParametro, 12)
    End Sub
    Public Sub Actualizar(d As NParamContabilidad)
        Dim parametros() As Object = {"@idrucsiscorp", "@nombresiscorp", "@flagcli", "@flagventas", "@flagcompras", "@flagprov", "@periodo", "@flagventasdetalle", "@flagcomprasdetalle", "@flagpagos", "@flagcobros", "@flagcobconta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.idrucsiscorp, d.nombresiscorp, d.flagcli, d.flagventas, d.flagcompras, d.flagprov, d.periodo, d.flagventasdetalle, d.flagcomprasdetalle, d.flagpagos, d.flagcobros, d.flagcobconta}
        sql.EjecutarProcedure("Str_ParamContabilidad_U", parametros, valores, tipoParametro, 12)
    End Sub
    Public Function Agregar(d As NParamContabilidad, Retornatable As Boolean) As NParamContabilidad

        Dim parametros() As Object = {"@idrucsiscorp", "@nombresiscorp", "@flagcli", "@flagventas", "@flagcompras", "@flagprov", "@periodo", "@flagventasdetalle", "@flagcomprasdetalle", "@flagpagos", "@flagcobros", "@flagcobconta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.idrucsiscorp, d.nombresiscorp, d.flagcli, d.flagventas, d.flagcompras, d.flagprov, d.periodo, d.flagventasdetalle, d.flagcomprasdetalle, d.flagpagos, d.flagcobros, d.flagcobconta}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ParamContabilidad_I_S", parametros, valores, tipoParametro, 12).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idrucsiscorp = IIf(dt.Rows(0).Item("idrucsiscorp") Is DBNull.Value, Nothing, dt.Rows(0).Item("idrucsiscorp"))
            d.nombresiscorp = IIf(dt.Rows(0).Item("nombresiscorp") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombresiscorp"))
            d.flagcli = IIf(dt.Rows(0).Item("flagcli") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcli"))
            d.flagventas = IIf(dt.Rows(0).Item("flagventas") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagventas"))
            d.flagcompras = IIf(dt.Rows(0).Item("flagcompras") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcompras"))
            d.flagprov = IIf(dt.Rows(0).Item("flagprov") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagprov"))
            d.periodo = IIf(dt.Rows(0).Item("periodo") Is DBNull.Value, Nothing, dt.Rows(0).Item("periodo"))
            d.flagventasdetalle = IIf(dt.Rows(0).Item("flagventasdetalle") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagventasdetalle"))
            d.flagcomprasdetalle = IIf(dt.Rows(0).Item("flagcomprasdetalle") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcomprasdetalle"))
            d.flagpagos = IIf(dt.Rows(0).Item("flagpagos") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagpagos"))
            d.flagcobros = IIf(dt.Rows(0).Item("flagcobros") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcobros"))
            d.flagcobconta = IIf(dt.Rows(0).Item("flagcobconta") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcobconta"))
        Else
            d.idrucsiscorp = Nothing
            d.nombresiscorp = Nothing
            d.flagcli = Nothing
            d.flagventas = Nothing
            d.flagcompras = Nothing
            d.flagprov = Nothing
            d.periodo = Nothing
            d.flagventasdetalle = Nothing
            d.flagcomprasdetalle = Nothing
            d.flagpagos = Nothing
            d.flagcobros = Nothing
            d.flagcobconta = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NParamContabilidad, Retornatable As Boolean) As NParamContabilidad

        Dim parametros() As Object = {"@idrucsiscorp", "@nombresiscorp", "@flagcli", "@flagventas", "@flagcompras", "@flagprov", "@periodo", "@flagventasdetalle", "@flagcomprasdetalle", "@flagpagos", "@flagcobros", "@flagcobconta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit}
        Dim valores() As Object = {d.flagcobconta = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ParamContabilidad_U_S", parametros, valores, tipoParametro, 36).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idrucsiscorp = IIf(dt.Rows(0).Item("idrucsiscorp") Is DBNull.Value, Nothing, dt.Rows(0).Item("idrucsiscorp"))
            d.nombresiscorp = IIf(dt.Rows(0).Item("nombresiscorp") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombresiscorp"))
            d.flagcli = IIf(dt.Rows(0).Item("flagcli") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcli"))
            d.flagventas = IIf(dt.Rows(0).Item("flagventas") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagventas"))
            d.flagcompras = IIf(dt.Rows(0).Item("flagcompras") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcompras"))
            d.flagprov = IIf(dt.Rows(0).Item("flagprov") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagprov"))
            d.periodo = IIf(dt.Rows(0).Item("periodo") Is DBNull.Value, Nothing, dt.Rows(0).Item("periodo"))
            d.flagventasdetalle = IIf(dt.Rows(0).Item("flagventasdetalle") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagventasdetalle"))
            d.flagcomprasdetalle = IIf(dt.Rows(0).Item("flagcomprasdetalle") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcomprasdetalle"))
            d.flagpagos = IIf(dt.Rows(0).Item("flagpagos") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagpagos"))
            d.flagcobros = IIf(dt.Rows(0).Item("flagcobros") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcobros"))
            d.flagcobconta = IIf(dt.Rows(0).Item("flagcobconta") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcobconta"))
        Else
            d.idrucsiscorp = Nothing
            d.nombresiscorp = Nothing
            d.flagcli = Nothing
            d.flagventas = Nothing
            d.flagcompras = Nothing
            d.flagprov = Nothing
            d.periodo = Nothing
            d.flagventasdetalle = Nothing
            d.flagcomprasdetalle = Nothing
            d.flagpagos = Nothing
            d.flagcobros = Nothing
            d.flagcobconta = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NParamContabilidad)
        Dim parametros() As Object = {"@idrucsiscorp"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idrucsiscorp}
        sql.EjecutarProcedure("Str_ParamContabilidad_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_ParamContabilidad(d As NParamContabilidad) As Boolean
        Dim parametros() As Object = {"@idrucsiscorp"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idrucsiscorp}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_ParamContabilidad", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idrucsiscorp"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ParamContabilidad_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NParamContabilidad) As DataTable
        Dim parametros() As Object = {"@idrucsiscorp"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idrucsiscorp}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ParamContabilidad_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NParamContabilidad) As NParamContabilidad
        Dim parametros() As Object = {"@idrucsiscorp"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idrucsiscorp}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ParamContabilidad_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idrucsiscorp = IIf(dt.Rows(0).Item("idrucsiscorp") Is DBNull.Value, Nothing, dt.Rows(0).Item("idrucsiscorp"))
            d.nombresiscorp = IIf(dt.Rows(0).Item("nombresiscorp") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombresiscorp"))
            d.flagcli = IIf(dt.Rows(0).Item("flagcli") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcli"))
            d.flagventas = IIf(dt.Rows(0).Item("flagventas") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagventas"))
            d.flagcompras = IIf(dt.Rows(0).Item("flagcompras") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcompras"))
            d.flagprov = IIf(dt.Rows(0).Item("flagprov") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagprov"))
            d.periodo = IIf(dt.Rows(0).Item("periodo") Is DBNull.Value, Nothing, dt.Rows(0).Item("periodo"))
            d.flagventasdetalle = IIf(dt.Rows(0).Item("flagventasdetalle") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagventasdetalle"))
            d.flagcomprasdetalle = IIf(dt.Rows(0).Item("flagcomprasdetalle") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcomprasdetalle"))
            d.flagpagos = IIf(dt.Rows(0).Item("flagpagos") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagpagos"))
            d.flagcobros = IIf(dt.Rows(0).Item("flagcobros") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcobros"))
            d.flagcobconta = IIf(dt.Rows(0).Item("flagcobconta") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagcobconta"))
        Else
            d.idrucsiscorp = Nothing
            d.nombresiscorp = Nothing
            d.flagcli = Nothing
            d.flagventas = Nothing
            d.flagcompras = Nothing
            d.flagprov = Nothing
            d.periodo = Nothing
            d.flagventasdetalle = Nothing
            d.flagcomprasdetalle = Nothing
            d.flagpagos = Nothing
            d.flagcobros = Nothing
            d.flagcobconta = Nothing
        End If
        Return d
    End Function
#End Region


End Class
