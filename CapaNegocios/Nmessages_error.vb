Imports CapaDatos
Public Class Nmessages_error
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property message_id As Integer
    Public Property language_id As Short
    Public Property severity As Byte
    Public Property is_event_logged As Boolean
    Public Property text As String
    Public Property mensaje_espaniol As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As Nmessages_error)

        Dim parametros() As Object = {"@message_id", "@language_id", "@severity", "@is_event_logged", "@text", "@mensaje_espaniol"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Int, SqlDbType.Int, SqlDbType.Bit, SqlDbType.NVarChar, SqlDbType.NVarChar}
        Dim valores() As Object = {d.message_id, d.language_id, d.severity, d.is_event_logged, d.text, d.mensaje_espaniol}
        sql.EjecutarProcedure("Str_messages_error_I", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Actualizar(d As Nmessages_error)
        Dim parametros() As Object = {"@message_id", "@language_id", "@severity", "@is_event_logged", "@text", "@mensaje_espaniol"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Int, SqlDbType.Int, SqlDbType.Bit, SqlDbType.NVarChar, SqlDbType.NVarChar}
        Dim valores() As Object = {d.message_id, d.language_id, d.severity, d.is_event_logged, d.text, d.mensaje_espaniol}
        sql.EjecutarProcedure("Str_messages_error_U", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Eliminar(d As Nmessages_error)
        Dim parametros() As Object = {"@message_id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.message_id}
        sql.EjecutarProcedure("Str_messages_error_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@message_id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_messages_error_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Nmessages_error) As DataTable
        Dim parametros() As Object = {"@message_id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.message_id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_messages_error_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Nmessages_error) As Nmessages_error
        Dim parametros() As Object = {"@message_id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.message_id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_messages_error_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.message_id = IIf(dt.Rows(0).Item("message_id") Is DBNull.Value, Nothing, dt.Rows(0).Item("message_id"))
            d.language_id = IIf(dt.Rows(0).Item("language_id") Is DBNull.Value, Nothing, dt.Rows(0).Item("language_id"))
            d.severity = IIf(dt.Rows(0).Item("severity") Is DBNull.Value, Nothing, dt.Rows(0).Item("severity"))
            d.is_event_logged = IIf(dt.Rows(0).Item("is_event_logged") Is DBNull.Value, Nothing, dt.Rows(0).Item("is_event_logged"))
            d.text = IIf(dt.Rows(0).Item("text") Is DBNull.Value, Nothing, dt.Rows(0).Item("text"))
            d.mensaje_espaniol = IIf(dt.Rows(0).Item("mensaje_espaniol") Is DBNull.Value, Nothing, dt.Rows(0).Item("mensaje_espaniol"))
        Else
            d.message_id = Nothing
            d.language_id = Nothing
            d.severity = Nothing
            d.is_event_logged = Nothing
            d.text = Nothing
            d.mensaje_espaniol = Nothing
        End If
        Return d
    End Function
#End Region


End Class
