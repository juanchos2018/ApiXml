Imports CapaDatos

Public Class NPtUsuario_Formulario
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idformulario As Long
    Public Property formulario_name As String
    Public Property idusuario As String
    Public Property habilitar As Boolean
    Public Property mostrar As Boolean
    Public Property control_nombre As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NPtUsuario_Formulario)

        Dim parametros() As Object = {"@formulario_name", "@idusuario", "@habilitar", "@mostrar", "@control_nombre"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar}
        Dim valores() As Object = {d.formulario_name, d.idusuario, d.habilitar, d.mostrar, d.control_nombre}
        sql.EjecutarProcedure("Str_PtUsuario_Formulario_I", parametros, valores, tipoParametro, 5)
    End Sub
    Public Sub Actualizar(d As NPtUsuario_Formulario)
        Dim parametros() As Object = {"@idformulario", "@formulario_name", "@idusuario", "@habilitar", "@mostrar", "@control_nombre"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar}
        Dim valores() As Object = {d.idformulario, d.formulario_name, d.idusuario, d.habilitar, d.mostrar, d.control_nombre}
        sql.EjecutarProcedure("Str_PtUsuario_Formulario_U", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Eliminar(d As NPtUsuario_Formulario)
        Dim parametros() As Object = {"@idformulario"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idformulario}
        sql.EjecutarProcedure("Str_PtUsuario_Formulario_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idformulario"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_PtUsuario_Formulario_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NPtUsuario_Formulario) As DataTable
        Dim parametros() As Object = {"@idformulario"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idformulario}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_PtUsuario_Formulario_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NPtUsuario_Formulario) As NPtUsuario_Formulario
        Dim parametros() As Object = {"@idformulario"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idformulario}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_PtUsuario_Formulario_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idformulario = IIf(dt.Rows(0).Item("idformulario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idformulario"))
            d.formulario_name = IIf(dt.Rows(0).Item("formulario_name") Is DBNull.Value, Nothing, dt.Rows(0).Item("formulario_name"))
            d.idusuario = IIf(dt.Rows(0).Item("idusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuario"))
            d.habilitar = IIf(dt.Rows(0).Item("habilitar") Is DBNull.Value, Nothing, dt.Rows(0).Item("habilitar"))
            d.mostrar = IIf(dt.Rows(0).Item("mostrar") Is DBNull.Value, Nothing, dt.Rows(0).Item("mostrar"))
            d.control_nombre = IIf(dt.Rows(0).Item("control_nombre") Is DBNull.Value, Nothing, dt.Rows(0).Item("control_nombre"))
        Else
            d.idformulario = Nothing
            d.formulario_name = Nothing
            d.idusuario = Nothing
            d.habilitar = Nothing
            d.mostrar = Nothing
            d.control_nombre = Nothing
        End If
        Return d
    End Function
    Public Function registro_usuario(id As String, formName As String) As DataTable
        Dim c As String = " select IdFormulario,Formulario_Name,IdUsuario,Habilitar,MOstrar,Control_Nombre from ptusuario_formulario where IdUsuario='" & id & "' and Formulario_Name='" & formName & "'"
        Return sql.EjecutarConsulta("d", c).Tables(0)
    End Function
#End Region

End Class
