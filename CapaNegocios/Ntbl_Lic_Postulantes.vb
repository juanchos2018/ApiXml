Imports CapaDatos
Public Class Ntbl_Lic_Postulantes
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idpostulantes As Long
    Public Property puesto As Integer
    Public Property idempresapostulante As String
    Public Property precioofertado As Decimal
    Public Property esganador As Boolean
    Public Property tbl_procesos_idproceso As Long
    Public Property comentario As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As Ntbl_Lic_Postulantes)
        Dim parametros() As Object = {"@puesto", "@idempresapostulante", "@precioofertado", "@esganador", "@tbl_procesos_idproceso", "@comentario"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Bit, SqlDbType.BigInt, SqlDbType.VarChar}
        Dim valores() As Object = {d.puesto, d.idempresapostulante, d.precioofertado, d.esganador, d.tbl_procesos_idproceso, d.comentario}
        sql.EjecutarProcedure("Str_tbl_Lic_Postulantes_I", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Actualizar(d As Ntbl_Lic_Postulantes)
        Dim parametros() As Object = {"@puesto", "@idempresapostulante", "@precioofertado", "@esganador", "@tbl_procesos_idproceso", "@comentario"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Bit, SqlDbType.BigInt, SqlDbType.VarChar}
        Dim valores() As Object = {d.puesto, d.idempresapostulante, d.precioofertado, d.esganador, d.tbl_procesos_idproceso, d.comentario}
        sql.EjecutarProcedure("Str_tbl_Lic_Postulantes_U", parametros, valores, tipoParametro, 6)
    End Sub
    Public Function Agregar(d As Ntbl_Lic_Postulantes, Retornatable As Boolean) As Ntbl_Lic_Postulantes
        Dim parametros() As Object = {"@puesto", "@idempresapostulante", "@precioofertado", "@esganador", "@tbl_procesos_idproceso", "@comentario"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Bit, SqlDbType.BigInt, SqlDbType.VarChar}
        Dim valores() As Object = {d.puesto, d.idempresapostulante, d.precioofertado, d.esganador, d.tbl_procesos_idproceso, d.comentario}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_Postulantes_I_S", parametros, valores, tipoParametro, 6).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idpostulantes = IIf(dt.Rows(0).Item("idpostulantes") Is DBNull.Value, Nothing, dt.Rows(0).Item("idpostulantes"))
            d.puesto = IIf(dt.Rows(0).Item("puesto") Is DBNull.Value, Nothing, dt.Rows(0).Item("puesto"))
            d.idempresapostulante = IIf(dt.Rows(0).Item("idempresapostulante") Is DBNull.Value, Nothing, dt.Rows(0).Item("idempresapostulante"))
            d.precioofertado = IIf(dt.Rows(0).Item("precioofertado") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioofertado"))
            d.esganador = IIf(dt.Rows(0).Item("esganador") Is DBNull.Value, Nothing, dt.Rows(0).Item("esganador"))
            d.tbl_procesos_idproceso = IIf(dt.Rows(0).Item("tbl_procesos_idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("tbl_procesos_idproceso"))
            d.comentario = IIf(dt.Rows(0).Item("comentario") Is DBNull.Value, Nothing, dt.Rows(0).Item("comentario"))
        Else
            d.idpostulantes = Nothing
            d.puesto = Nothing
            d.idempresapostulante = Nothing
            d.precioofertado = Nothing
            d.esganador = Nothing
            d.tbl_procesos_idproceso = Nothing
            d.comentario = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_Lic_Postulantes, Retornatable As Boolean) As Ntbl_Lic_Postulantes
        Dim parametros() As Object = {"@puesto", "@idempresapostulante", "@precioofertado", "@esganador", "@tbl_procesos_idproceso", "@comentario"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Bit, SqlDbType.BigInt, SqlDbType.VarChar}
        Dim valores() As Object = {d.comentario = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_Postulantes_U_S", parametros, valores, tipoParametro, 6).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idpostulantes = IIf(dt.Rows(0).Item("idpostulantes") Is DBNull.Value, Nothing, dt.Rows(0).Item("idpostulantes"))
            d.puesto = IIf(dt.Rows(0).Item("puesto") Is DBNull.Value, Nothing, dt.Rows(0).Item("puesto"))
            d.idempresapostulante = IIf(dt.Rows(0).Item("idempresapostulante") Is DBNull.Value, Nothing, dt.Rows(0).Item("idempresapostulante"))
            d.precioofertado = IIf(dt.Rows(0).Item("precioofertado") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioofertado"))
            d.esganador = IIf(dt.Rows(0).Item("esganador") Is DBNull.Value, Nothing, dt.Rows(0).Item("esganador"))
            d.tbl_procesos_idproceso = IIf(dt.Rows(0).Item("tbl_procesos_idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("tbl_procesos_idproceso"))
            d.comentario = IIf(dt.Rows(0).Item("comentario") Is DBNull.Value, Nothing, dt.Rows(0).Item("comentario"))
        Else
            d.idpostulantes = Nothing
            d.puesto = Nothing
            d.idempresapostulante = Nothing
            d.precioofertado = Nothing
            d.esganador = Nothing
            d.tbl_procesos_idproceso = Nothing
            d.comentario = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_Lic_Postulantes)
        Dim parametros() As Object = {"@idpostulantes", "@IdProceso"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.BigInt}
        Dim valores() As Object = {d.idpostulantes, d.tbl_procesos_idproceso}
        sql.EjecutarProcedure("Str_tbl_Lic_Postulantes_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Existe_tbl_Lic_Postulantes(d As Ntbl_Lic_Postulantes)
        Dim parametros() As Object = {"@idpostulantes"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idpostulantes}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_tbl_Lic_Postulantes", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idpostulantes"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_Postulantes_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_Lic_Postulantes) As DataTable
        Dim parametros() As Object = {"@idpostulantes"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idpostulantes}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_Postulantes_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function ListaFK(d As Ntbl_Lic_Postulantes) As DataTable
        Dim parametros() As Object = {"@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.tbl_procesos_idproceso}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_PostulantesPK_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Ntbl_Lic_Postulantes) As Ntbl_Lic_Postulantes
        Dim parametros() As Object = {"@idpostulantes"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idpostulantes}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_Postulantes_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idpostulantes = IIf(dt.Rows(0).Item("idpostulantes") Is DBNull.Value, Nothing, dt.Rows(0).Item("idpostulantes"))
            d.puesto = IIf(dt.Rows(0).Item("puesto") Is DBNull.Value, Nothing, dt.Rows(0).Item("puesto"))
            d.idempresapostulante = IIf(dt.Rows(0).Item("idempresapostulante") Is DBNull.Value, Nothing, dt.Rows(0).Item("idempresapostulante"))
            d.precioofertado = IIf(dt.Rows(0).Item("precioofertado") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioofertado"))
            d.esganador = IIf(dt.Rows(0).Item("esganador") Is DBNull.Value, Nothing, dt.Rows(0).Item("esganador"))
            d.tbl_procesos_idproceso = IIf(dt.Rows(0).Item("tbl_procesos_idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("tbl_procesos_idproceso"))
            d.comentario = IIf(dt.Rows(0).Item("comentario") Is DBNull.Value, Nothing, dt.Rows(0).Item("comentario"))
        Else
            d.idpostulantes = Nothing
            d.puesto = Nothing
            d.idempresapostulante = Nothing
            d.precioofertado = Nothing
            d.esganador = Nothing
            d.tbl_procesos_idproceso = Nothing
            d.comentario = Nothing
        End If
        Return d
    End Function
#End Region

End Class
