Imports CapaDatos

Public Class Ntbl_DepachosYura
    Dim sql As New ClsConexion

#Region "Declarations"

    Public Property idticket As String
    Public Property idconductor As String
    Public Property idtransportista As String
    Public Property placatrackto As String
    Public Property placacarreta As String
    Public Property nroguiacliente As String
    Public Property nroguiatransportista As String
    Public Property estado As Boolean
    Public Property fecharegistro As System.DateTime
    Public Property fechadocumento As System.DateTime
    Public Property usuariocrea As String


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As Ntbl_DepachosYura)

        Dim parametros() As Object = {"@idticket", "@idconductor", "@idtransportista", "@placatrackto", "@placacarreta", "@nroguiacliente", "@nroguiatransportista", "@estado", "@fecharegistro", "@fechadocumento", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idticket, d.idconductor, d.idtransportista, d.placatrackto, d.placacarreta, d.nroguiacliente, d.nroguiatransportista, d.estado, d.fecharegistro, d.fechadocumento, d.usuariocrea}
        sql.EjecutarProcedure("Str_tbl_DepachosYura_I", parametros, valores, tipoParametro, 11)
    End Sub
    Public Sub Actualizar(d As Ntbl_DepachosYura)
        Dim parametros() As Object = {"@idticket", "@idconductor", "@idtransportista", "@placatrackto", "@placacarreta", "@nroguiacliente", "@nroguiatransportista", "@estado", "@fecharegistro", "@fechadocumento", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idticket, d.idconductor, d.idtransportista, d.placatrackto, d.placacarreta, d.nroguiacliente, d.nroguiatransportista, d.estado, d.fecharegistro, d.fechadocumento, d.usuariocrea}
        sql.EjecutarProcedure("Str_tbl_DepachosYura_U", parametros, valores, tipoParametro, 11)
    End Sub
    Public Function Agregar(d As Ntbl_DepachosYura, Retornatable As Boolean) As Ntbl_DepachosYura

        Dim parametros() As Object = {"@idticket", "@idconductor", "@idtransportista", "@placatrackto", "@placacarreta", "@nroguiacliente", "@nroguiatransportista", "@estado", "@fecharegistro", "@fechadocumento", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idticket, d.idconductor, d.idtransportista, d.placatrackto, d.placacarreta, d.nroguiacliente, d.nroguiatransportista, d.estado, d.fecharegistro, d.fechadocumento, d.usuariocrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DepachosYura_I_S", parametros, valores, tipoParametro, 11).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idticket = IIf(dt.Rows(0).Item("idticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("idticket"))
            d.idconductor = IIf(dt.Rows(0).Item("idconductor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idconductor"))
            d.idtransportista = IIf(dt.Rows(0).Item("idtransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtransportista"))
            d.placatrackto = IIf(dt.Rows(0).Item("placatrackto") Is DBNull.Value, Nothing, dt.Rows(0).Item("placatrackto"))
            d.placacarreta = IIf(dt.Rows(0).Item("placacarreta") Is DBNull.Value, Nothing, dt.Rows(0).Item("placacarreta"))
            d.nroguiacliente = IIf(dt.Rows(0).Item("nroguiacliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroguiacliente"))
            d.nroguiatransportista = IIf(dt.Rows(0).Item("nroguiatransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroguiatransportista"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.fecharegistro = IIf(dt.Rows(0).Item("fecharegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecharegistro"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.idticket = Nothing
            d.idconductor = Nothing
            d.idtransportista = Nothing
            d.placatrackto = Nothing
            d.placacarreta = Nothing
            d.nroguiacliente = Nothing
            d.nroguiatransportista = Nothing
            d.estado = Nothing
            d.fecharegistro = Nothing
            d.fechadocumento = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_DepachosYura, Retornatable As Boolean) As Ntbl_DepachosYura

        Dim parametros() As Object = {"@idticket", "@idconductor", "@idtransportista", "@placatrackto", "@placacarreta", "@nroguiacliente", "@nroguiatransportista", "@estado", "@fecharegistro", "@fechadocumento", "@usuariocrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.usuariocrea = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DepachosYura_U_S", parametros, valores, tipoParametro, 33).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idticket = IIf(dt.Rows(0).Item("idticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("idticket"))
            d.idconductor = IIf(dt.Rows(0).Item("idconductor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idconductor"))
            d.idtransportista = IIf(dt.Rows(0).Item("idtransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtransportista"))
            d.placatrackto = IIf(dt.Rows(0).Item("placatrackto") Is DBNull.Value, Nothing, dt.Rows(0).Item("placatrackto"))
            d.placacarreta = IIf(dt.Rows(0).Item("placacarreta") Is DBNull.Value, Nothing, dt.Rows(0).Item("placacarreta"))
            d.nroguiacliente = IIf(dt.Rows(0).Item("nroguiacliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroguiacliente"))
            d.nroguiatransportista = IIf(dt.Rows(0).Item("nroguiatransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroguiatransportista"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.fecharegistro = IIf(dt.Rows(0).Item("fecharegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecharegistro"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.idticket = Nothing
            d.idconductor = Nothing
            d.idtransportista = Nothing
            d.placatrackto = Nothing
            d.placacarreta = Nothing
            d.nroguiacliente = Nothing
            d.nroguiatransportista = Nothing
            d.estado = Nothing
            d.fecharegistro = Nothing
            d.fechadocumento = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_DepachosYura)
        Dim parametros() As Object = {"@idticket"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idticket}
        sql.EjecutarProcedure("Str_tbl_DepachosYura_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_tbl_DepachosYura(d As Ntbl_DepachosYura) As Boolean
        Dim parametros() As Object = {"@idticket"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idticket}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_tbl_DepachosYura", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idticket"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DepachosYura_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_DepachosYura) As DataTable
        Dim parametros() As Object = {"@idticket"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idticket}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DepachosYura_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Ntbl_DepachosYura) As Ntbl_DepachosYura
        Dim parametros() As Object = {"@idticket"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idticket}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_DepachosYura_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idticket = IIf(dt.Rows(0).Item("idticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("idticket"))
            d.idconductor = IIf(dt.Rows(0).Item("idconductor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idconductor"))
            d.idtransportista = IIf(dt.Rows(0).Item("idtransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtransportista"))
            d.placatrackto = IIf(dt.Rows(0).Item("placatrackto") Is DBNull.Value, Nothing, dt.Rows(0).Item("placatrackto"))
            d.placacarreta = IIf(dt.Rows(0).Item("placacarreta") Is DBNull.Value, Nothing, dt.Rows(0).Item("placacarreta"))
            d.nroguiacliente = IIf(dt.Rows(0).Item("nroguiacliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroguiacliente"))
            d.nroguiatransportista = IIf(dt.Rows(0).Item("nroguiatransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroguiatransportista"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.fecharegistro = IIf(dt.Rows(0).Item("fecharegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecharegistro"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
        Else
            d.idticket = Nothing
            d.idconductor = Nothing
            d.idtransportista = Nothing
            d.placatrackto = Nothing
            d.placacarreta = Nothing
            d.nroguiacliente = Nothing
            d.nroguiatransportista = Nothing
            d.estado = Nothing
            d.fecharegistro = Nothing
            d.fechadocumento = Nothing
            d.usuariocrea = Nothing
        End If
        Return d
    End Function
#End Region


End Class
