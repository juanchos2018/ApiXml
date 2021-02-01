Imports CapaDatos
Public Class NptBotones_acceso
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idusuario As String
    Public Property idbotonname As String
    Public Property caption_boton As String
    Public Property name_form As String
    Public Property text_form As String
    Public Property dimension_imagen As String
    Public Property icono_png As Byte()

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NptBotones_acceso)

        Dim parametros() As Object = {"@idusuario", "@idbotonname", "@caption_boton", "@name_form", "@text_form", "@dimension_imagen", "@icono_png"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary}
        Dim valores() As Object = {d.idusuario, d.idbotonname, d.caption_boton, d.name_form, d.text_form, d.dimension_imagen, d.icono_png}
        sql.EjecutarProcedure("Str_ptBotones_acceso_I", parametros, valores, tipoParametro, 7)
    End Sub
    Public Sub Actualizar(d As NptBotones_acceso)
        Dim parametros() As Object = {"@idusuario", "@idbotonname", "@caption_boton", "@name_form", "@text_form", "@dimension_imagen", "@icono_png"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary}
        Dim valores() As Object = {d.idusuario, d.idbotonname, d.caption_boton, d.name_form, d.text_form, d.dimension_imagen, d.icono_png}
        sql.EjecutarProcedure("Str_ptBotones_acceso_U", parametros, valores, tipoParametro, 7)
    End Sub
    Public Function Agregar(d As NptBotones_acceso, Retornatable As Boolean) As NptBotones_acceso

        Dim parametros() As Object = {"@idusuario", "@idbotonname", "@caption_boton", "@name_form", "@text_form", "@dimension_imagen", "@icono_png"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary}
        Dim valores() As Object = {d.idusuario, d.idbotonname, d.caption_boton, d.name_form, d.text_form, d.dimension_imagen, d.icono_png}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ptBotones_acceso_I_S", parametros, valores, tipoParametro, 7).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idusuario = IIf(dt.Rows(0).Item("idusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuario"))
            d.idbotonname = IIf(dt.Rows(0).Item("idbotonname") Is DBNull.Value, Nothing, dt.Rows(0).Item("idbotonname"))
            d.caption_boton = IIf(dt.Rows(0).Item("caption_boton") Is DBNull.Value, Nothing, dt.Rows(0).Item("caption_boton"))
            d.name_form = IIf(dt.Rows(0).Item("name_form") Is DBNull.Value, Nothing, dt.Rows(0).Item("name_form"))
            d.text_form = IIf(dt.Rows(0).Item("text_form") Is DBNull.Value, Nothing, dt.Rows(0).Item("text_form"))
            d.dimension_imagen = IIf(dt.Rows(0).Item("dimension_imagen") Is DBNull.Value, Nothing, dt.Rows(0).Item("dimension_imagen"))
            d.icono_png = IIf(dt.Rows(0).Item("icono_png") Is DBNull.Value, Nothing, dt.Rows(0).Item("icono_png"))
        Else
            d.idusuario = Nothing
            d.idbotonname = Nothing
            d.caption_boton = Nothing
            d.name_form = Nothing
            d.text_form = Nothing
            d.dimension_imagen = Nothing
            d.icono_png = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NptBotones_acceso, Retornatable As Boolean) As NptBotones_acceso

        Dim parametros() As Object = {"@idusuario", "@idbotonname", "@caption_boton", "@name_form", "@text_form", "@dimension_imagen", "@icono_png"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary}
        Dim valores() As Object = {d.idusuario, d.idbotonname, d.caption_boton, d.name_form, d.text_form, d.dimension_imagen, d.icono_png}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ptBotones_acceso_U_S", parametros, valores, tipoParametro, 7).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idusuario = IIf(dt.Rows(0).Item("idusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuario"))
            d.idbotonname = IIf(dt.Rows(0).Item("idbotonname") Is DBNull.Value, Nothing, dt.Rows(0).Item("idbotonname"))
            d.caption_boton = IIf(dt.Rows(0).Item("caption_boton") Is DBNull.Value, Nothing, dt.Rows(0).Item("caption_boton"))
            d.name_form = IIf(dt.Rows(0).Item("name_form") Is DBNull.Value, Nothing, dt.Rows(0).Item("name_form"))
            d.text_form = IIf(dt.Rows(0).Item("text_form") Is DBNull.Value, Nothing, dt.Rows(0).Item("text_form"))
            d.dimension_imagen = IIf(dt.Rows(0).Item("dimension_imagen") Is DBNull.Value, Nothing, dt.Rows(0).Item("dimension_imagen"))
            d.icono_png = IIf(dt.Rows(0).Item("icono_png") Is DBNull.Value, Nothing, dt.Rows(0).Item("icono_png"))
        Else
            d.idusuario = Nothing
            d.idbotonname = Nothing
            d.caption_boton = Nothing
            d.name_form = Nothing
            d.text_form = Nothing
            d.dimension_imagen = Nothing
            d.icono_png = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NptBotones_acceso)
        Dim parametros() As Object = {"@idusuario", "@idbotonname"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idusuario, d.idbotonname}
        sql.EjecutarProcedure("Str_ptBotones_acceso_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idusuario", "@idbotonname"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ptBotones_acceso_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NptBotones_acceso) As DataTable
        Dim parametros() As Object = {"@idusuario", "@idbotonname"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idusuario, d.idbotonname}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ptBotones_acceso_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NptBotones_acceso) As NptBotones_acceso
        Dim parametros() As Object = {"@idusuario", "@idbotonname"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idusuario, d.idbotonname}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ptBotones_acceso_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idusuario = IIf(dt.Rows(0).Item("idusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuario"))
            d.idbotonname = IIf(dt.Rows(0).Item("idbotonname") Is DBNull.Value, Nothing, dt.Rows(0).Item("idbotonname"))
            d.caption_boton = IIf(dt.Rows(0).Item("caption_boton") Is DBNull.Value, Nothing, dt.Rows(0).Item("caption_boton"))
            d.name_form = IIf(dt.Rows(0).Item("name_form") Is DBNull.Value, Nothing, dt.Rows(0).Item("name_form"))
            d.text_form = IIf(dt.Rows(0).Item("text_form") Is DBNull.Value, Nothing, dt.Rows(0).Item("text_form"))
            d.dimension_imagen = IIf(dt.Rows(0).Item("dimension_imagen") Is DBNull.Value, Nothing, dt.Rows(0).Item("dimension_imagen"))
            d.icono_png = IIf(dt.Rows(0).Item("icono_png") Is DBNull.Value, Nothing, dt.Rows(0).Item("icono_png"))
        Else
            d.idusuario = Nothing
            d.idbotonname = Nothing
            d.caption_boton = Nothing
            d.name_form = Nothing
            d.text_form = Nothing
            d.dimension_imagen = Nothing
            d.icono_png = Nothing
        End If
        Return d
    End Function
#End Region

End Class
