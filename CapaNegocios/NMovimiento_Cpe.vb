Imports CapaDatos
Public Class NMovimiento_Cpe
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property tipodocumento As String
    Public Property numerodocumento As String
    Public Property xml_zip As Byte()
    Public Property cdr_zip As Byte()
    Public Property pdf_pdf As Byte()
    Public Property codigobarras As Byte()
    Public Property codigohash As String
    Public Property signaturevalue As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NMovimiento_Cpe)

        Dim parametros() As Object = {"@tipodocumento", "@numerodocumento", "@xml_zip", "@cdr_zip", "@pdf_pdf", "@codigobarras", "@codigohash", "@signaturevalue"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipodocumento, d.numerodocumento, d.xml_zip, d.cdr_zip, d.pdf_pdf, d.codigobarras, d.codigohash, d.signaturevalue}
        sql.EjecutarProcedure("Str_Movimiento_Cpe_I", parametros, valores, tipoParametro, 8)
    End Sub
    Public Sub Actualizar(d As NMovimiento_Cpe)
        Dim parametros() As Object = {"@tipodocumento", "@numerodocumento", "@xml_zip", "@cdr_zip", "@pdf_pdf", "@codigobarras", "@codigohash", "@signaturevalue"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipodocumento, d.numerodocumento, d.xml_zip, d.cdr_zip, d.pdf_pdf, d.codigobarras, d.codigohash, d.signaturevalue}
        sql.EjecutarProcedure("Str_Movimiento_Cpe_U", parametros, valores, tipoParametro, 8)
    End Sub
    Public Function Agregar(d As NMovimiento_Cpe, Retornatable As Boolean) As NMovimiento_Cpe

        Dim parametros() As Object = {"@tipodocumento", "@numerodocumento", "@xml_zip", "@cdr_zip", "@pdf_pdf", "@codigobarras", "@codigohash", "@signaturevalue"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipodocumento, d.numerodocumento, d.xml_zip, d.cdr_zip, d.pdf_pdf, d.codigobarras, d.codigohash, d.signaturevalue}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Movimiento_Cpe_I_S", parametros, valores, tipoParametro, 8).Tables(0)
        If dt.Rows.Count > 0 Then
            d.tipodocumento = IIf(dt.Rows(0).Item("tipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocumento"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.xml_zip = IIf(dt.Rows(0).Item("xml_zip") Is DBNull.Value, Nothing, dt.Rows(0).Item("xml_zip"))
            d.cdr_zip = IIf(dt.Rows(0).Item("cdr_zip") Is DBNull.Value, Nothing, dt.Rows(0).Item("cdr_zip"))
            d.pdf_pdf = IIf(dt.Rows(0).Item("pdf_pdf") Is DBNull.Value, Nothing, dt.Rows(0).Item("pdf_pdf"))
            d.codigobarras = IIf(dt.Rows(0).Item("codigobarras") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigobarras"))
            d.codigohash = IIf(dt.Rows(0).Item("codigohash") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigohash"))
            d.signaturevalue = IIf(dt.Rows(0).Item("signaturevalue") Is DBNull.Value, Nothing, dt.Rows(0).Item("signaturevalue"))
        Else
            d.tipodocumento = Nothing
            d.numerodocumento = Nothing
            d.xml_zip = Nothing
            d.cdr_zip = Nothing
            d.pdf_pdf = Nothing
            d.codigobarras = Nothing
            d.codigohash = Nothing
            d.signaturevalue = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NMovimiento_Cpe, Retornatable As Boolean) As NMovimiento_Cpe
        Dim parametros() As Object = {"@tipodocumento", "@numerodocumento", "@xml_zip", "@cdr_zip", "@pdf_pdf", "@codigobarras", "@codigohash", "@signaturevalue"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarBinary, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.signaturevalue = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Movimiento_Cpe_U_S", parametros, valores, tipoParametro, 8).Tables(0)
        If dt.Rows.Count > 0 Then
            d.tipodocumento = IIf(dt.Rows(0).Item("tipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocumento"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.xml_zip = IIf(dt.Rows(0).Item("xml_zip") Is DBNull.Value, Nothing, dt.Rows(0).Item("xml_zip"))
            d.cdr_zip = IIf(dt.Rows(0).Item("cdr_zip") Is DBNull.Value, Nothing, dt.Rows(0).Item("cdr_zip"))
            d.pdf_pdf = IIf(dt.Rows(0).Item("pdf_pdf") Is DBNull.Value, Nothing, dt.Rows(0).Item("pdf_pdf"))
            d.codigobarras = IIf(dt.Rows(0).Item("codigobarras") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigobarras"))
            d.codigohash = IIf(dt.Rows(0).Item("codigohash") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigohash"))
            d.signaturevalue = IIf(dt.Rows(0).Item("signaturevalue") Is DBNull.Value, Nothing, dt.Rows(0).Item("signaturevalue"))
        Else
            d.tipodocumento = Nothing
            d.numerodocumento = Nothing
            d.xml_zip = Nothing
            d.cdr_zip = Nothing
            d.pdf_pdf = Nothing
            d.codigobarras = Nothing
            d.codigohash = Nothing
            d.signaturevalue = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NMovimiento_Cpe)
        Dim parametros() As Object = {"@tipodocumento", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipodocumento, d.numerodocumento}
        sql.EjecutarProcedure("Str_Movimiento_Cpe_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Existe_Movimiento_Cpe(d As NMovimiento_Cpe) As Boolean
        Dim parametros() As Object = {"@tipodocumento", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipodocumento, d.numerodocumento}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_Movimiento_Cpe", parametros, valores, tipoParametro, 2)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@tipodocumento", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Movimiento_Cpe_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NMovimiento_Cpe) As DataTable
        Dim parametros() As Object = {"@tipodocumento", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipodocumento, d.numerodocumento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Movimiento_Cpe_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NMovimiento_Cpe) As NMovimiento_Cpe
        Dim parametros() As Object = {"@tipodocumento", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipodocumento, d.numerodocumento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Movimiento_Cpe_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.tipodocumento = IIf(dt.Rows(0).Item("tipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocumento"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.xml_zip = IIf(dt.Rows(0).Item("xml_zip") Is DBNull.Value, Nothing, dt.Rows(0).Item("xml_zip"))
            d.cdr_zip = IIf(dt.Rows(0).Item("cdr_zip") Is DBNull.Value, Nothing, dt.Rows(0).Item("cdr_zip"))
            d.pdf_pdf = IIf(dt.Rows(0).Item("pdf_pdf") Is DBNull.Value, Nothing, dt.Rows(0).Item("pdf_pdf"))
            d.codigobarras = IIf(dt.Rows(0).Item("codigobarras") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigobarras"))
            d.codigohash = IIf(dt.Rows(0).Item("codigohash") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigohash"))
            d.signaturevalue = IIf(dt.Rows(0).Item("signaturevalue") Is DBNull.Value, Nothing, dt.Rows(0).Item("signaturevalue"))
        Else
            d.tipodocumento = Nothing
            d.numerodocumento = Nothing
            d.xml_zip = Nothing
            d.cdr_zip = Nothing
            d.pdf_pdf = Nothing
            d.codigobarras = Nothing
            d.codigohash = Nothing
            d.signaturevalue = Nothing
        End If
        Return d
    End Function
#End Region

End Class
