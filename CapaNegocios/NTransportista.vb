Imports CapaDatos
Public Class NTransportista
#Region "Declarations"
    Dim sql As New ClsConexion
    Public Property idTransportista As String
    Public Property nombre As String
    Public Property direccion As String
    Public Property telefono As String
    Public Property localidad As String
    Public Property ruc As String
    Public Property dGH As String
    Public Property estado As String
    Public Property usuarioCrea As String
    Public Property usuarioMod As String
    Public Property fechaCrea As System.DateTime
    Public Property fechaMod As System.DateTime
    Public Property autoMTC As String
    Public Property licencia As String
    Public Property marca As String
    Public Property placa As String
    Public Property chofer As String
    Public Property cubicaje As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NTransportista)

        Dim parametros() As Object = {"@autoMTC", "@chofer", "@cubicaje", "@dGH", "@direccion", "@estado", "@fechaCrea", "@fechaMod", "@idTransportista", "@licencia", "@localidad", "@marca", "@nombre", "@placa", "@ruc", "@telefono", "@usuarioCrea", "@usuarioMod"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.autoMTC, d.chofer, d.cubicaje, d.dGH, d.direccion, d.estado, d.fechaCrea, d.fechaMod, d.idTransportista, d.licencia, d.localidad, d.marca, d.nombre, d.placa, d.ruc, d.telefono, d.usuarioCrea, d.usuarioMod}
        sql.EjecutarProcedure("Str_Transportista_I", parametros, valores, tipoParametro, 18)
    End Sub
    Public Sub Actualizar(d As NTransportista)
        Dim parametros() As Object = {"@autoMTC", "@chofer", "@cubicaje", "@dGH", "@direccion", "@estado", "@fechaCrea", "@fechaMod", "@idTransportista", "@licencia", "@localidad", "@marca", "@nombre", "@placa", "@ruc", "@telefono", "@usuarioCrea", "@usuarioMod"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.autoMTC, d.chofer, d.cubicaje, d.dGH, d.direccion, d.estado, d.fechaCrea, d.fechaMod, d.idTransportista, d.licencia, d.localidad, d.marca, d.nombre, d.placa, d.ruc, d.telefono, d.usuarioCrea, d.usuarioMod}
        sql.EjecutarProcedure("Str_Transportista_U", parametros, valores, tipoParametro, 18)
    End Sub
    Public Sub Eliminar(d As NTransportista)
        Dim parametros() As Object = {"@idTransportista"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idTransportista}
        sql.EjecutarProcedure("Str_Transportista_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idTransportista"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Transportista_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTransportista) As DataTable
        Dim parametros() As Object = {"@idTransportista"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idTransportista}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Transportista_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTransportista) As NTransportista
        Dim parametros() As Object = {"@idTransportista"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idTransportista}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Transportista_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.autoMTC = IIf(dt.Rows(0).Item("autoMTC") Is DBNull.Value, Nothing, dt.Rows(0).Item("autoMTC"))
            d.chofer = IIf(dt.Rows(0).Item("chofer") Is DBNull.Value, Nothing, dt.Rows(0).Item("chofer"))
            d.cubicaje = IIf(dt.Rows(0).Item("cubicaje") Is DBNull.Value, Nothing, dt.Rows(0).Item("cubicaje"))
            d.dGH = IIf(dt.Rows(0).Item("dGH") Is DBNull.Value, Nothing, dt.Rows(0).Item("dGH"))
            d.direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.fechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.fechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.idTransportista = IIf(dt.Rows(0).Item("idTransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTransportista"))
            d.licencia = IIf(dt.Rows(0).Item("licencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("licencia"))
            d.localidad = IIf(dt.Rows(0).Item("localidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("localidad"))
            d.marca = IIf(dt.Rows(0).Item("marca") Is DBNull.Value, Nothing, dt.Rows(0).Item("marca"))
            d.nombre = IIf(dt.Rows(0).Item("nombre") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombre"))
            d.placa = IIf(dt.Rows(0).Item("placa") Is DBNull.Value, Nothing, dt.Rows(0).Item("placa"))
            d.ruc = IIf(dt.Rows(0).Item("ruc") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruc"))
            d.telefono = IIf(dt.Rows(0).Item("telefono") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono"))
            d.usuarioCrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.usuarioMod = IIf(dt.Rows(0).Item("usuarioMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioMod"))
        Else
            d.autoMTC = Nothing
            d.chofer = Nothing
            d.cubicaje = Nothing
            d.dGH = Nothing
            d.direccion = Nothing
            d.estado = Nothing
            d.fechaCrea = Nothing
            d.fechaMod = Nothing
            d.idTransportista = Nothing
            d.licencia = Nothing
            d.localidad = Nothing
            d.marca = Nothing
            d.nombre = Nothing
            d.placa = Nothing
            d.ruc = Nothing
            d.telefono = Nothing
            d.usuarioCrea = Nothing
            d.usuarioMod = Nothing
        End If
        Return d
    End Function
#End Region
End Class
