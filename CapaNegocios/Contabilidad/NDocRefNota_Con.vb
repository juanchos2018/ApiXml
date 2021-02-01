Imports CapaDatos
Public Class NDocRefNota_Con
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idSubdiario As String
    Public Property nroComprobante As String
    Public Property secComprobante As String
    Public Property secuencia As String
    Public Property idTipoDoc As String
    Public Property serieDoc As String
    Public Property nroDoc As String
    Public Property fechaDoc As System.DateTime
    Public Property baseIMN As Decimal
    Public Property baseIUS As Decimal
    Public Property iGVMN As Decimal
    Public Property iGVUS As Decimal
    Public Property Bd As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NDocRefNota_Con)
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@secComprobante", "@secuencia", "@idTipoDoc", "@serieDoc", "@nroDoc", "@fechaDoc", "@baseIMN", "@baseIUS", "@iGVMN", "@iGVUS"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal}
        Dim valores() As Object = {d.idSubdiario, d.nroComprobante, d.secComprobante, d.secuencia, d.idTipoDoc, d.serieDoc, d.nroDoc, d.fechaDoc, d.baseIMN, d.baseIUS, d.iGVMN, d.iGVUS}
        sql.EjecutarProcedure(d.Bd & ".dbo.Str_DocRefNota_I", parametros, valores, tipoParametro, 12)
    End Sub
    Public Sub Actualizar(d As NDocRefNota_Con)
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@secComprobante", "@secuencia", "@idTipoDoc", "@serieDoc", "@nroDoc", "@fechaDoc", "@baseIMN", "@baseIUS", "@iGVMN", "@iGVUS"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal}
        Dim valores() As Object = {d.idSubdiario, d.nroComprobante, d.secComprobante, d.secuencia, d.idTipoDoc, d.serieDoc, d.nroDoc, d.fechaDoc, d.baseIMN, d.baseIUS, d.iGVMN, d.iGVUS}
        sql.EjecutarProcedure(d.Bd & ".dbo.Str_DocRefNota_U", parametros, valores, tipoParametro, 12)
    End Sub
    Public Sub Eliminar(d As NDocRefNota_Con)
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@secComprobante", "@secuencia"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char}
        Dim valores() As Object = {d.idSubdiario, d.nroComprobante, d.secComprobante, d.secuencia}
        sql.EjecutarProcedure(d.Bd & ".dbo.Str_DocRefNota_D", parametros, valores, tipoParametro, 4)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@secComprobante", "@secuencia"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DocRefNota_S", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NDocRefNota_Con) As DataTable
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@secComprobante", "@secuencia"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char}
        Dim valores() As Object = {d.idSubdiario, d.nroComprobante, d.secComprobante, d.secuencia}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(d.Bd & ".dbo.Str_DocRefNota_S", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDocRefNota_Con) As NDocRefNota_Con
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@secComprobante", "@secuencia"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char}
        Dim valores() As Object = {d.idSubdiario, d.nroComprobante, d.secComprobante, d.secuencia}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(d.Bd & ".dbo.Str_DocRefNota_S", parametros, valores, tipoParametro, 4).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idSubdiario = IIf(dt.Rows(0).Item("idSubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idSubdiario"))
            d.nroComprobante = IIf(dt.Rows(0).Item("nroComprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroComprobante"))
            d.secComprobante = IIf(dt.Rows(0).Item("secComprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("secComprobante"))
            d.secuencia = IIf(dt.Rows(0).Item("secuencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("secuencia"))
            d.idTipoDoc = IIf(dt.Rows(0).Item("idTipoDoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoDoc"))
            d.serieDoc = IIf(dt.Rows(0).Item("serieDoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("serieDoc"))
            d.nroDoc = IIf(dt.Rows(0).Item("nroDoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroDoc"))
            d.fechaDoc = IIf(dt.Rows(0).Item("fechaDoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaDoc"))
            d.baseIMN = IIf(dt.Rows(0).Item("baseIMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("baseIMN"))
            d.baseIUS = IIf(dt.Rows(0).Item("baseIUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("baseIUS"))
            d.iGVMN = IIf(dt.Rows(0).Item("iGVMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("iGVMN"))
            d.iGVUS = IIf(dt.Rows(0).Item("iGVUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("iGVUS"))
        Else
            d.idSubdiario = Nothing
            d.nroComprobante = Nothing
            d.secComprobante = Nothing
            d.secuencia = Nothing
            d.idTipoDoc = Nothing
            d.serieDoc = Nothing
            d.nroDoc = Nothing
            d.fechaDoc = Nothing
            d.baseIMN = Nothing
            d.baseIUS = Nothing
            d.iGVMN = Nothing
            d.iGVUS = Nothing
        End If
        Return d
    End Function
#End Region

End Class
