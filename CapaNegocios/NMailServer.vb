Imports CapaDatos
Public Class NMailServer
    Dim sql As New ClsConexion

#Region "Declarations"

    Public Property smtp As String
    Public Property puerto As String
    Public Property ssl As Boolean
    Public Property credencial As Boolean
    Public Property mastermail As String
    Public Property pws As String
    Public Property cc As Boolean
    Public Property ccopiamail As String
    Public Property asunto As String
    Public Property cuerpomail As String
    Public Property id As Integer
    Public Property ftp As String
    Public Property userftp As String
    Public Property pwsftp As String
    Public Property pasivemode As Boolean

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NMailServer)

        Dim parametros() As Object = {"@smtp", "@puerto", "@ssl", "@credencial", "@mastermail", "@pws", "@cc", "@ccopiamail", "@asunto", "@cuerpomail", "@ftp", "@userftp", "@pwsftp", "@pasivemode"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.smtp, d.puerto, d.ssl, d.credencial, d.mastermail, d.pws, d.cc, d.ccopiamail, d.asunto, d.cuerpomail, d.ftp, d.userftp, d.pwsftp, d.pasivemode}
        sql.EjecutarProcedure("Str_MailServer_I", parametros, valores, tipoParametro, 14)
    End Sub
    Public Sub Actualizar(d As NMailServer)
        Dim parametros() As Object = {"@id", "@smtp", "@puerto", "@ssl", "@credencial", "@mastermail", "@pws", "@cc", "@ccopiamail", "@asunto", "@cuerpomail", "@ftp", "@userftp", "@pwsftp", "@pasivemode"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.id, d.smtp, d.puerto, d.ssl, d.credencial, d.mastermail, d.pws, d.cc, d.ccopiamail, d.asunto, d.cuerpomail, d.ftp, d.userftp, d.pwsftp, d.pasivemode}
        sql.EjecutarProcedure("Str_MailServer_U", parametros, valores, tipoParametro, 15)
    End Sub
    Public Function Agregar(d As NMailServer, Retornatable As Boolean) As NMailServer

        Dim parametros() As Object = {"@smtp", "@puerto", "@ssl", "@credencial", "@mastermail", "@pws", "@cc", "@ccopiamail", "@asunto", "@cuerpomail", "@ftp", "@userftp", "@pwsftp", "@pasivemode"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.smtp, d.puerto, d.ssl, d.credencial, d.mastermail, d.pws, d.cc, d.ccopiamail, d.asunto, d.cuerpomail, d.ftp, d.userftp, d.pwsftp, d.pasivemode}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_MailServer_I_S", parametros, valores, tipoParametro, 14).Tables(0)
        If dt.Rows.Count > 0 Then
            d.smtp = IIf(dt.Rows(0).Item("smtp") Is DBNull.Value, Nothing, dt.Rows(0).Item("smtp"))
            d.puerto = IIf(dt.Rows(0).Item("puerto") Is DBNull.Value, Nothing, dt.Rows(0).Item("puerto"))
            d.ssl = IIf(dt.Rows(0).Item("ssl") Is DBNull.Value, Nothing, dt.Rows(0).Item("ssl"))
            d.credencial = IIf(dt.Rows(0).Item("credencial") Is DBNull.Value, Nothing, dt.Rows(0).Item("credencial"))
            d.mastermail = IIf(dt.Rows(0).Item("mastermail") Is DBNull.Value, Nothing, dt.Rows(0).Item("mastermail"))
            d.pws = IIf(dt.Rows(0).Item("pws") Is DBNull.Value, Nothing, dt.Rows(0).Item("pws"))
            d.cc = IIf(dt.Rows(0).Item("cc") Is DBNull.Value, Nothing, dt.Rows(0).Item("cc"))
            d.ccopiamail = IIf(dt.Rows(0).Item("ccopiamail") Is DBNull.Value, Nothing, dt.Rows(0).Item("ccopiamail"))
            d.asunto = IIf(dt.Rows(0).Item("asunto") Is DBNull.Value, Nothing, dt.Rows(0).Item("asunto"))
            d.cuerpomail = IIf(dt.Rows(0).Item("cuerpomail") Is DBNull.Value, Nothing, dt.Rows(0).Item("cuerpomail"))
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.ftp = IIf(dt.Rows(0).Item("ftp") Is DBNull.Value, Nothing, dt.Rows(0).Item("ftp"))
            d.userftp = IIf(dt.Rows(0).Item("userftp") Is DBNull.Value, Nothing, dt.Rows(0).Item("userftp"))
            d.pwsftp = IIf(dt.Rows(0).Item("pwsftp") Is DBNull.Value, Nothing, dt.Rows(0).Item("pwsftp"))
            d.pasivemode = IIf(dt.Rows(0).Item("pasivemode") Is DBNull.Value, Nothing, dt.Rows(0).Item("pasivemode"))
        Else
            d.smtp = Nothing
            d.puerto = Nothing
            d.ssl = Nothing
            d.credencial = Nothing
            d.mastermail = Nothing
            d.pws = Nothing
            d.cc = Nothing
            d.ccopiamail = Nothing
            d.asunto = Nothing
            d.cuerpomail = Nothing
            d.id = Nothing
            d.ftp = Nothing
            d.userftp = Nothing
            d.pwsftp = Nothing
            d.pasivemode = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NMailServer, Retornatable As Boolean) As NMailServer

        Dim parametros() As Object = {"@smtp", "@puerto", "@ssl", "@credencial", "@mastermail", "@pws", "@cc", "@ccopiamail", "@asunto", "@cuerpomail", "@ftp", "@userftp", "@pwsftp", "@pasivemode"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.pasivemode = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_MailServer_U_S", parametros, valores, tipoParametro, 44).Tables(0)
        If dt.Rows.Count > 0 Then
            d.smtp = IIf(dt.Rows(0).Item("smtp") Is DBNull.Value, Nothing, dt.Rows(0).Item("smtp"))
            d.puerto = IIf(dt.Rows(0).Item("puerto") Is DBNull.Value, Nothing, dt.Rows(0).Item("puerto"))
            d.ssl = IIf(dt.Rows(0).Item("ssl") Is DBNull.Value, Nothing, dt.Rows(0).Item("ssl"))
            d.credencial = IIf(dt.Rows(0).Item("credencial") Is DBNull.Value, Nothing, dt.Rows(0).Item("credencial"))
            d.mastermail = IIf(dt.Rows(0).Item("mastermail") Is DBNull.Value, Nothing, dt.Rows(0).Item("mastermail"))
            d.pws = IIf(dt.Rows(0).Item("pws") Is DBNull.Value, Nothing, dt.Rows(0).Item("pws"))
            d.cc = IIf(dt.Rows(0).Item("cc") Is DBNull.Value, Nothing, dt.Rows(0).Item("cc"))
            d.ccopiamail = IIf(dt.Rows(0).Item("ccopiamail") Is DBNull.Value, Nothing, dt.Rows(0).Item("ccopiamail"))
            d.asunto = IIf(dt.Rows(0).Item("asunto") Is DBNull.Value, Nothing, dt.Rows(0).Item("asunto"))
            d.cuerpomail = IIf(dt.Rows(0).Item("cuerpomail") Is DBNull.Value, Nothing, dt.Rows(0).Item("cuerpomail"))
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.ftp = IIf(dt.Rows(0).Item("ftp") Is DBNull.Value, Nothing, dt.Rows(0).Item("ftp"))
            d.userftp = IIf(dt.Rows(0).Item("userftp") Is DBNull.Value, Nothing, dt.Rows(0).Item("userftp"))
            d.pwsftp = IIf(dt.Rows(0).Item("pwsftp") Is DBNull.Value, Nothing, dt.Rows(0).Item("pwsftp"))
            d.pasivemode = IIf(dt.Rows(0).Item("pasivemode") Is DBNull.Value, Nothing, dt.Rows(0).Item("pasivemode"))
        Else
            d.smtp = Nothing
            d.puerto = Nothing
            d.ssl = Nothing
            d.credencial = Nothing
            d.mastermail = Nothing
            d.pws = Nothing
            d.cc = Nothing
            d.ccopiamail = Nothing
            d.asunto = Nothing
            d.cuerpomail = Nothing
            d.id = Nothing
            d.ftp = Nothing
            d.userftp = Nothing
            d.pwsftp = Nothing
            d.pasivemode = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NMailServer)
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.id}
        sql.EjecutarProcedure("Str_MailServer_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_MailServer(d As NMailServer) As Boolean
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.id}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_MailServer", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_MailServer_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NMailServer) As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_MailServer_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NMailServer) As NMailServer
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_MailServer_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.smtp = IIf(dt.Rows(0).Item("smtp") Is DBNull.Value, Nothing, dt.Rows(0).Item("smtp"))
            d.puerto = IIf(dt.Rows(0).Item("puerto") Is DBNull.Value, Nothing, dt.Rows(0).Item("puerto"))
            d.ssl = IIf(dt.Rows(0).Item("ssl") Is DBNull.Value, Nothing, dt.Rows(0).Item("ssl"))
            d.credencial = IIf(dt.Rows(0).Item("credencial") Is DBNull.Value, Nothing, dt.Rows(0).Item("credencial"))
            d.mastermail = IIf(dt.Rows(0).Item("mastermail") Is DBNull.Value, Nothing, dt.Rows(0).Item("mastermail"))
            d.pws = IIf(dt.Rows(0).Item("pws") Is DBNull.Value, Nothing, dt.Rows(0).Item("pws"))
            d.cc = IIf(dt.Rows(0).Item("cc") Is DBNull.Value, Nothing, dt.Rows(0).Item("cc"))
            d.ccopiamail = IIf(dt.Rows(0).Item("ccopiamail") Is DBNull.Value, Nothing, dt.Rows(0).Item("ccopiamail"))
            d.asunto = IIf(dt.Rows(0).Item("asunto") Is DBNull.Value, Nothing, dt.Rows(0).Item("asunto"))
            d.cuerpomail = IIf(dt.Rows(0).Item("cuerpomail") Is DBNull.Value, Nothing, dt.Rows(0).Item("cuerpomail"))
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.ftp = IIf(dt.Rows(0).Item("ftp") Is DBNull.Value, Nothing, dt.Rows(0).Item("ftp"))
            d.userftp = IIf(dt.Rows(0).Item("userftp") Is DBNull.Value, Nothing, dt.Rows(0).Item("userftp"))
            d.pwsftp = IIf(dt.Rows(0).Item("pwsftp") Is DBNull.Value, Nothing, dt.Rows(0).Item("pwsftp"))
            d.pasivemode = IIf(dt.Rows(0).Item("pasivemode") Is DBNull.Value, Nothing, dt.Rows(0).Item("pasivemode"))
        Else
            d.smtp = Nothing
            d.puerto = Nothing
            d.ssl = Nothing
            d.credencial = Nothing
            d.mastermail = Nothing
            d.pws = Nothing
            d.cc = Nothing
            d.ccopiamail = Nothing
            d.asunto = Nothing
            d.cuerpomail = Nothing
            d.id = Nothing
            d.ftp = Nothing
            d.userftp = Nothing
            d.pwsftp = Nothing
            d.pasivemode = Nothing
        End If
        Return d
    End Function
#End Region


End Class
