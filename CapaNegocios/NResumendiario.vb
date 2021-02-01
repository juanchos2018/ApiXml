Imports CapaDatos
Public Class NResumendiario
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _id As Integer
    Private _nroDocumento As String
    Private _nroTicket As String
    Private _fechaEnvio As System.DateTime
    Private _estado As String
    Private _observaciones As String
    Private _diaResumen As String
    Private _tipoDoc As String
    Private _nroIdSunat As String
    Private _fechaRecepcion As System.DateTime
    Private _horarecepcion As String
    Private _horaCDR As String
    Private _nota As String
    Private _nroDocEnviado As String
    Private _descripcionerror As String
    Private _nroDocFirmado As String
    Private _idAquiriente As String
    Private _codRecepcion As String
    Private _fechaCDR As System.DateTime
    Private _xml_zip As Byte()
    Private _cdr_zip As Byte()

#End Region

#Region "Properties"

    Public Property Id As Integer
        Get
            Return _id
        End Get
        Set
            _id = Value
        End Set
    End Property

    Public Property NroDocumento As String
        Get
            Return _nroDocumento
        End Get
        Set
            _nroDocumento = Value
        End Set
    End Property

    Public Property NroTicket As String
        Get
            Return _nroTicket
        End Get
        Set
            _nroTicket = Value
        End Set
    End Property

    Public Property FechaEnvio As System.DateTime
        Get
            Return _fechaEnvio
        End Get
        Set
            _fechaEnvio = Value
        End Set
    End Property

    Public Property Estado As String
        Get
            Return _estado
        End Get
        Set
            _estado = Value
        End Set
    End Property

    Public Property Observaciones As String
        Get
            Return _observaciones
        End Get
        Set
            _observaciones = Value
        End Set
    End Property

    Public Property DiaResumen As String
        Get
            Return _diaResumen
        End Get
        Set
            _diaResumen = Value
        End Set
    End Property

    Public Property TipoDoc As String
        Get
            Return _tipoDoc
        End Get
        Set
            _tipoDoc = Value
        End Set
    End Property

    Public Property NroIdSunat As String
        Get
            Return _nroIdSunat
        End Get
        Set
            _nroIdSunat = Value
        End Set
    End Property

    Public Property FechaRecepcion As System.DateTime
        Get
            Return _fechaRecepcion
        End Get
        Set
            _fechaRecepcion = Value
        End Set
    End Property

    Public Property Horarecepcion As String
        Get
            Return _horarecepcion
        End Get
        Set
            _horarecepcion = Value
        End Set
    End Property

    Public Property HoraCDR As String
        Get
            Return _horaCDR
        End Get
        Set
            _horaCDR = Value
        End Set
    End Property

    Public Property Nota As String
        Get
            Return _nota
        End Get
        Set
            _nota = Value
        End Set
    End Property

    Public Property NroDocEnviado As String
        Get
            Return _nroDocEnviado
        End Get
        Set
            _nroDocEnviado = Value
        End Set
    End Property

    Public Property Descripcionerror As String
        Get
            Return _descripcionerror
        End Get
        Set
            _descripcionerror = Value
        End Set
    End Property

    Public Property NroDocFirmado As String
        Get
            Return _nroDocFirmado
        End Get
        Set
            _nroDocFirmado = Value
        End Set
    End Property

    Public Property IdAquiriente As String
        Get
            Return _idAquiriente
        End Get
        Set
            _idAquiriente = Value
        End Set
    End Property

    Public Property CodRecepcion As String
        Get
            Return _codRecepcion
        End Get
        Set
            _codRecepcion = Value
        End Set
    End Property

    Public Property FechaCDR As System.DateTime
        Get
            Return _fechaCDR
        End Get
        Set
            _fechaCDR = Value
        End Set
    End Property

    Public Property xml_zip As Byte()
        Get
            Return _xml_zip
        End Get
        Set
            _xml_zip = Value
        End Set
    End Property

    Public Property cdr_zip As Byte()
        Get
            Return _cdr_zip
        End Get
        Set
            _cdr_zip = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    'Public Sub New()

    'End Sub

    'Public Sub New(ByVal id As Integer, ByVal nroDocumento As String, ByVal nroTicket As String, ByVal fechaEnvio As System.DateTime, ByVal estado As String, ByVal observaciones As String, ByVal diaResumen As String, ByVal tipoDoc As String, ByVal nroIdSunat As String, ByVal fechaRecepcion As System.DateTime, ByVal horarecepcion As String, ByVal horaCDR As String, ByVal nota As String, ByVal nroDocEnviado As String, ByVal descripcionerror As String, ByVal nroDocFirmado As String, ByVal idAquiriente As String, ByVal codRecepcion As String, ByVal fechaCDR As System.DateTime, ByVal xml_zip As Byte(), ByVal cdr_zip As Byte())
    '    Me.New()
    'End Sub

#End Region
#Region "Metodos"
    Public Sub agregar(d As NResumendiario)
        Dim parametros() As Object = {"@nroDocumento", "@nroTicket", "@fechaEnvio", "@estado", "@observaciones", "@diaResumen", "@tipoDoc", "@nroIdSunat", "@fechaRecepcion", "@horarecepcion", "@horaCDR", "@nota", "@nroDocEnviado", "@descripcionerror", "@nroDocFirmado", "@idAquiriente", "@codRecepcion", "@fechaCDR", "@xml_zip", "@cdr_zip"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarBinary, SqlDbType.VarBinary}
        Dim valores() As Object = {d.NroDocumento, d.NroTicket, d.FechaEnvio, d.Estado, d.Observaciones, d.DiaResumen, d.TipoDoc, d.NroIdSunat, d.FechaRecepcion, d.Horarecepcion, d.HoraCDR, d.Nota, d.NroDocEnviado, d.Descripcionerror, d.NroDocFirmado, d.IdAquiriente, d.CodRecepcion, d.FechaCDR, d.xml_zip, d.cdr_zip}
        sql.EjecutarProcedure("Str_Agrega_Resumen", parametros, valores, tipoParametro, 20)
    End Sub
    Public Sub actualizar(d As NResumendiario)
        Dim parametros() As Object = {"@Id", "@nroTicket", "@fechaEnvio", "@estado", "@observaciones", "@diaResumen", "@tipoDoc", "@nroIdSunat", "@fechaRecepcion", "@horarecepcion", "@horaCDR", "@nota", "@nroDocEnviado", "@descripcionerror", "@nroDocFirmado", "@idAquiriente", "@codRecepcion", "@fechaCDR", "@cdr_zip"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarBinary}
        Dim valores() As Object = {d.Id, d.NroTicket, d.FechaEnvio, d.Estado, d.Observaciones, d.DiaResumen, d.TipoDoc, d.NroIdSunat, d.FechaRecepcion, d.Horarecepcion, d.HoraCDR, d.Nota, d.NroDocEnviado, d.Descripcionerror, d.NroDocFirmado, d.IdAquiriente, d.CodRecepcion, d.FechaCDR, d.cdr_zip}
        sql.EjecutarProcedure("Str_UptResumenCDR", parametros, valores, tipoParametro, 19)
    End Sub

    Public Function xmlzip(d As NResumendiario) As Byte()
        Dim cadena As String = " select xml_zip from tbl_resumendiario "
        cadena += " where id=" & d.Id & ""
        Dim dt As New DataTable
        dt = sql.EjecutarConsulta("d", cadena).Tables(0)
        If dt.Rows.Count > 0 Then
            d.xml_zip = dt.Rows(0).Item(0)
        End If
        Return d.xml_zip
    End Function

    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@Id"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_ResumenDiario_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Sub actualizaCDR(d As NResumendiario)
        Dim valoresx As String = "NroIdSunat='" & d.NroIdSunat & "',FechaRecepcion='" & d.FechaRecepcion & "',HoraRecepcion='" & d.Horarecepcion.ToString & "',"
        valoresx += " FechaCDR='" & d.FechaCDR & "',HoraCDR='" & d.HoraCDR.ToString & "',Nota='" & d.Nota & "',NroDocEnviado='" & d.NroDocEnviado & "',"
        valoresx += " CodRecepcion='" & d.CodRecepcion & "',Descripcionerror='" & d.Descripcionerror & "',NroDocFirmado='" & d.NroDocFirmado & "',IdAquiriente='" & d.IdAquiriente & "'"
        sql.Editar("Tbl_ResumenDiario", valoresx, "Id='" & d.Id & "'")
    End Sub

    Public Function Registro(d As NResumendiario) As NResumendiario

        Dim parametros() As Object = {"@Id"}


        Dim tipoParametro() As Object = {SqlDbType.VarChar}

        Dim valores() As Object = {d.Id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_ResumenDiario_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.Id = dt.Rows(0).Item("id")
            d.NroDocumento = dt.Rows(0).Item("nroDocumento")
            d.NroTicket = dt.Rows(0).Item("nroTicket")
            d.FechaEnvio = dt.Rows(0).Item("fechaEnvio")
            d.Estado = dt.Rows(0).Item("estado")
            d.Observaciones = dt.Rows(0).Item("observaciones").ToString 
            d.DiaResumen = dt.Rows(0).Item("diaResumen")
            d.TipoDoc = dt.Rows(0).Item("tipoDoc")
            d.NroIdSunat = dt.Rows(0).Item("nroIdSunat").ToString
            d.FechaRecepcion = dt.Rows(0).Item("fechaRecepcion")
            d.Horarecepcion = dt.Rows(0).Item("horarecepcion").ToString
            d.HoraCDR = dt.Rows(0).Item("horaCDR").ToString
            d.Nota = dt.Rows(0).Item("nota").ToString
            d.NroDocEnviado = dt.Rows(0).Item("nroDocEnviado").ToString
            d.Descripcionerror = dt.Rows(0).Item("descripcionerror").ToString
            d.NroDocFirmado = dt.Rows(0).Item("nroDocFirmado").ToString
            d.IdAquiriente = dt.Rows(0).Item("idAquiriente").ToString
            d.CodRecepcion = dt.Rows(0).Item("codRecepcion").ToString 
            d.FechaCDR = dt.Rows(0).Item("fechaCDR")
            d.xml_zip = IIf(dt.Rows(0).Item("xml_zip") Is DBNull.Value, Nothing, dt.Rows(0).Item("xml_zip"))

        Else
            d.NroDocumento = ""
            d.NroTicket = ""

            d.Estado = ""
            d.Observaciones = ""
            d.DiaResumen = ""
            d.TipoDoc = ""
            d.NroIdSunat = ""

            d.Horarecepcion = ""
            d.HoraCDR = ""
            d.Nota = ""
            d.NroDocEnviado = ""
            d.Descripcionerror = ""
            d.NroDocFirmado = ""
            d.IdAquiriente = ""
            d.CodRecepcion = ""

        End If
        Return d
    End Function

    Public Function ListaResumenBajas(fi As String, ff As String, td As String) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF", "@Td"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {fi, ff, td}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ListaResumenBajas", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function

#End Region


End Class
