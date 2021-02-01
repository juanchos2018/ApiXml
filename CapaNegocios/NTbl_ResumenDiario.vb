Imports CapaDatos

Public Class NTbl_ResumenDiario
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property id As Integer
    Public Property nrodocumento As String
    Public Property nroticket As String
    Public Property fechaenvio As System.DateTime
    Public Property estado As String
    Public Property observaciones As String
    Public Property diaresumen As String
    Public Property tipodoc As String
    Public Property nroidsunat As String
    Public Property fecharecepcion As System.DateTime
    Public Property horarecepcion As String
    Public Property horacdr As String
    Public Property nota As String
    Public Property nrodocenviado As String
    Public Property descripcionerror As String
    Public Property nrodocfirmado As String
    Public Property idaquiriente As String
    Public Property codrecepcion As String
    Public Property fechacdr As System.DateTime
    Public Property xml_zip As Byte()
    Public Property cdr_zip As Byte()

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NTbl_ResumenDiario)

        Dim parametros() As Object = {"@nrodocumento", "@nroticket", "@fechaenvio", "@estado", "@observaciones", "@diaresumen", "@tipodoc", "@nroidsunat", "@fecharecepcion", "@horarecepcion", "@horacdr", "@nota", "@nrodocenviado", "@descripcionerror", "@nrodocfirmado", "@idaquiriente", "@codrecepcion", "@fechacdr", "@xml_zip", "@cdr_zip"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarBinary, SqlDbType.VarBinary}
        Dim valores() As Object = {d.nrodocumento, d.nroticket, d.fechaenvio, d.estado, d.observaciones, d.diaresumen, d.tipodoc, d.nroidsunat, d.fecharecepcion, d.horarecepcion, d.horacdr, d.nota, d.nrodocenviado, d.descripcionerror, d.nrodocfirmado, d.idaquiriente, d.codrecepcion, d.fechacdr, d.xml_zip, d.cdr_zip}
        sql.EjecutarProcedure("Str_Tbl_ResumenDiario_I", parametros, valores, tipoParametro, 20)
    End Sub
    Public Sub Actualizar(d As NTbl_ResumenDiario)
        Dim parametros() As Object = {"@nrodocumento", "@nroticket", "@fechaenvio", "@estado", "@observaciones", "@diaresumen", "@tipodoc", "@nroidsunat", "@fecharecepcion", "@horarecepcion", "@horacdr", "@nota", "@nrodocenviado", "@descripcionerror", "@nrodocfirmado", "@idaquiriente", "@codrecepcion", "@fechacdr", "@xml_zip", "@cdr_zip"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarBinary, SqlDbType.VarBinary}
        Dim valores() As Object = {d.nrodocumento, d.nroticket, d.fechaenvio, d.estado, d.observaciones, d.diaresumen, d.tipodoc, d.nroidsunat, d.fecharecepcion, d.horarecepcion, d.horacdr, d.nota, d.nrodocenviado, d.descripcionerror, d.nrodocfirmado, d.idaquiriente, d.codrecepcion, d.fechacdr, d.xml_zip, d.cdr_zip}
        sql.EjecutarProcedure("Str_Tbl_ResumenDiario_U", parametros, valores, tipoParametro, 20)
    End Sub
    Public Sub Eliminar(d As NTbl_ResumenDiario)
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.id}
        sql.EjecutarProcedure("Str_Tbl_ResumenDiario_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_Tbl_ResumenDiario(d As NTbl_ResumenDiario) As Boolean
        Dim parametros() As Object = {"@id", "@nrodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar}
        Dim valores() As Object = {d.id, d.nrodocumento}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_Tbl_ResumenDiario", parametros, valores, tipoParametro, 2)
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
        dt = sql.ProcedureSQL("Str_Tbl_ResumenDiario_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTbl_ResumenDiario) As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_ResumenDiario_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTbl_ResumenDiario) As NTbl_ResumenDiario
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_ResumenDiario_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.nrodocumento = IIf(dt.Rows(0).Item("nrodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocumento"))
            d.nroticket = IIf(dt.Rows(0).Item("nroticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroticket"))
            d.fechaenvio = IIf(dt.Rows(0).Item("fechaenvio") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaenvio"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.observaciones = IIf(dt.Rows(0).Item("observaciones") Is DBNull.Value, Nothing, dt.Rows(0).Item("observaciones"))
            d.diaresumen = IIf(dt.Rows(0).Item("diaresumen") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaresumen"))
            d.tipodoc = IIf(dt.Rows(0).Item("tipodoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodoc"))
            d.nroidsunat = IIf(dt.Rows(0).Item("nroidsunat") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroidsunat"))
            d.fecharecepcion = IIf(dt.Rows(0).Item("fecharecepcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecharecepcion"))
            d.horarecepcion = IIf(dt.Rows(0).Item("horarecepcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("horarecepcion"))
            d.horacdr = IIf(dt.Rows(0).Item("horacdr") Is DBNull.Value, Nothing, dt.Rows(0).Item("horacdr"))
            d.nota = IIf(dt.Rows(0).Item("nota") Is DBNull.Value, Nothing, dt.Rows(0).Item("nota"))
            d.nrodocenviado = IIf(dt.Rows(0).Item("nrodocenviado") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocenviado"))
            d.descripcionerror = IIf(dt.Rows(0).Item("descripcionerror") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcionerror"))
            d.nrodocfirmado = IIf(dt.Rows(0).Item("nrodocfirmado") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocfirmado"))
            d.idaquiriente = IIf(dt.Rows(0).Item("idaquiriente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idaquiriente"))
            d.codrecepcion = IIf(dt.Rows(0).Item("codrecepcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("codrecepcion"))
            d.fechacdr = IIf(dt.Rows(0).Item("fechacdr") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacdr"))
            d.xml_zip = IIf(dt.Rows(0).Item("xml_zip") Is DBNull.Value, Nothing, dt.Rows(0).Item("xml_zip"))
            d.cdr_zip = IIf(dt.Rows(0).Item("cdr_zip") Is DBNull.Value, Nothing, dt.Rows(0).Item("cdr_zip"))
        Else
            d.id = Nothing
            d.nrodocumento = Nothing
            d.nroticket = Nothing
            d.fechaenvio = Nothing
            d.estado = Nothing
            d.observaciones = Nothing
            d.diaresumen = Nothing
            d.tipodoc = Nothing
            d.nroidsunat = Nothing
            d.fecharecepcion = Nothing
            d.horarecepcion = Nothing
            d.horacdr = Nothing
            d.nota = Nothing
            d.nrodocenviado = Nothing
            d.descripcionerror = Nothing
            d.nrodocfirmado = Nothing
            d.idaquiriente = Nothing
            d.codrecepcion = Nothing
            d.fechacdr = Nothing
            d.xml_zip = Nothing
            d.cdr_zip = Nothing
        End If
        Return d
    End Function
#End Region


End Class
