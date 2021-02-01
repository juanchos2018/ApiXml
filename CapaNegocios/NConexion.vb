Public Module NConexion
    Public Property aServidor As String
    Public Property aBaseDatos As String
    Public Property aUsuario As String
    Public Property aContrasena As String


    Public Sub Parametros(ByVal aServidor As String, ByVal aUsuario As String, ByVal aContrasena As String, ByVal aBaseDatos As String)
        CapaDatos.pc_Servidor = aServidor
        CapaDatos.pc_BaseDatos = aBaseDatos
        CapaDatos.pc_Usuario = aUsuario
        CapaDatos.pc_Contrasena = aContrasena
        aServidor = aServidor
        aBaseDatos = aBaseDatos
        aUsuario = aUsuario
        aContrasena = aContrasena
    End Sub
End Module




