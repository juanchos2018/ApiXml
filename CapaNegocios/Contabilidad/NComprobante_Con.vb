Imports CapaDatos
Public Class NComprobante_Con
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _idSubdiario As String
    Private _nroComprobante As String
    Private _fechaComprobante As String
    Private _idMoneda As String
    Private _situacion As String
    Private _tipoCambio As Decimal
    Private _glosa As String
    Private _total As Decimal
    Private _idCambioMoneda As String
    Private _conConversion As String
    Private _fechaRegistro As System.DateTime
    Private _horaRegistro As String
    Private _usuario As String
    Private _fechaCambioMoneda As String
    Private _corig As String
    Private _cform As String
    Private _ctipcom As String
    Private _cextor As String
    Private _fechaComprobante2 As System.DateTime
    Private _fechaCambioMoneda2 As System.DateTime
    Private _flagCom As Boolean
    Private _tipoBien As String
    Private _c_IdAnexo As String
    Private _c_TipoDocumento As String
    Private _c_NroDocumento As String
    Private _c_FechaDocumento As System.DateTime
    Private _c_FechaVencimiento As System.DateTime
    Private _c_Importe As Decimal
    Private _c_Igv As Decimal
    Private _c_Total As Decimal
    Private _c_Inafecto As Decimal
    Private _c_TasaIGV As Decimal
    Private _c_ImporteISC As Decimal
    Private _c_Detraccion As Boolean
    Private _c_NroDetraccion As String
    Private _c_FechaDetraccion As System.DateTime
    Private _c_ImporteDetraccion As Decimal
    Private _c_IdTipoAnexo As String
    Private _c_ImporteMN As Decimal
    Private _c_ImporteUS As Decimal
    Private _c_IGVMN As Decimal
    Private _c_IGVUS As Decimal
    Private _c_TotalMN As Decimal
    Private _c_TotalUS As Decimal
    Private _c_ImporteISCMN As Decimal
    Private _c_ImporteISCUS As Decimal
    Private _c_Agencia_Aduana As String
    Private _c_ImporteRta4ta As Decimal
    Private _c_IdTipoAnexo1 As String
    Private _c_IdAnexo1 As String
    Private _c_IdVinculo_Residencia As String
    Private _c_Tipo_Renta As String
    Private _c_Mod_Serv_NoDomiciliado As String
    Private _c_Ex_Op_Nodomic As String
    Private _c_Tdref_import As String
    Private _c_Serie_import As String
    Private _c_Nroref_Import As String
    Private _c_Fecha_Aduana As System.DateTime
    Private _Bd As String

#End Region

#Region "Properties"

    Public Property idSubdiario As String
        Get
            Return _idSubdiario
        End Get
        Set
            _idSubdiario = Value
        End Set
    End Property

    Public Property NroComprobante As String
        Get
            Return _nroComprobante
        End Get
        Set
            _nroComprobante = Value
        End Set
    End Property

    Public Property FechaComprobante As String
        Get
            Return _fechaComprobante
        End Get
        Set
            _fechaComprobante = Value
        End Set
    End Property

    Public Property idMoneda As String
        Get
            Return _idMoneda
        End Get
        Set
            _idMoneda = Value
        End Set
    End Property

    Public Property Situacion As String
        Get
            Return _situacion
        End Get
        Set
            _situacion = Value
        End Set
    End Property

    Public Property TipoCambio As Decimal
        Get
            Return _tipoCambio
        End Get
        Set
            _tipoCambio = Value
        End Set
    End Property

    Public Property Glosa As String
        Get
            Return _glosa
        End Get
        Set
            _glosa = Value
        End Set
    End Property

    Public Property Total As Decimal
        Get
            Return _total
        End Get
        Set
            _total = Value
        End Set
    End Property

    Public Property idCambioMoneda As String
        Get
            Return _idCambioMoneda
        End Get
        Set
            _idCambioMoneda = Value
        End Set
    End Property

    Public Property ConConversion As String
        Get
            Return _conConversion
        End Get
        Set
            _conConversion = Value
        End Set
    End Property

    Public Property FechaRegistro As System.DateTime
        Get
            Return _fechaRegistro
        End Get
        Set
            _fechaRegistro = Value
        End Set
    End Property

    Public Property HoraRegistro As String
        Get
            Return _horaRegistro
        End Get
        Set
            _horaRegistro = Value
        End Set
    End Property

    Public Property Usuario As String
        Get
            Return _usuario
        End Get
        Set
            _usuario = Value
        End Set
    End Property

    Public Property FechaCambioMoneda As String
        Get
            Return _fechaCambioMoneda
        End Get
        Set
            _fechaCambioMoneda = Value
        End Set
    End Property

    Public Property corig As String
        Get
            Return _corig
        End Get
        Set
            _corig = Value
        End Set
    End Property

    Public Property cform As String
        Get
            Return _cform
        End Get
        Set
            _cform = Value
        End Set
    End Property

    Public Property ctipcom As String
        Get
            Return _ctipcom
        End Get
        Set
            _ctipcom = Value
        End Set
    End Property

    Public Property cextor As String
        Get
            Return _cextor
        End Get
        Set
            _cextor = Value
        End Set
    End Property

    Public Property FechaComprobante2 As System.DateTime
        Get
            Return _fechaComprobante2
        End Get
        Set
            _fechaComprobante2 = Value
        End Set
    End Property

    Public Property FechaCambioMoneda2 As System.DateTime
        Get
            Return _fechaCambioMoneda2
        End Get
        Set
            _fechaCambioMoneda2 = Value
        End Set
    End Property

    Public Property FlagCom As Boolean
        Get
            Return _flagCom
        End Get
        Set
            _flagCom = Value
        End Set
    End Property

    Public Property TipoBien As String
        Get
            Return _tipoBien
        End Get
        Set
            _tipoBien = Value
        End Set
    End Property

    Public Property C_IdAnexo As String
        Get
            Return _c_IdAnexo
        End Get
        Set
            _c_IdAnexo = Value
        End Set
    End Property

    Public Property C_TipoDocumento As String
        Get
            Return _c_TipoDocumento
        End Get
        Set
            _c_TipoDocumento = Value
        End Set
    End Property

    Public Property C_NroDocumento As String
        Get
            Return _c_NroDocumento
        End Get
        Set
            _c_NroDocumento = Value
        End Set
    End Property

    Public Property C_FechaDocumento As System.DateTime
        Get
            Return _c_FechaDocumento
        End Get
        Set
            _c_FechaDocumento = Value
        End Set
    End Property

    Public Property C_FechaVencimiento As System.DateTime
        Get
            Return _c_FechaVencimiento
        End Get
        Set
            _c_FechaVencimiento = Value
        End Set
    End Property

    Public Property C_Importe As Decimal
        Get
            Return _c_Importe
        End Get
        Set
            _c_Importe = Value
        End Set
    End Property

    Public Property C_Igv As Decimal
        Get
            Return _c_Igv
        End Get
        Set
            _c_Igv = Value
        End Set
    End Property

    Public Property C_Total As Decimal
        Get
            Return _c_Total
        End Get
        Set
            _c_Total = Value
        End Set
    End Property

    Public Property C_Inafecto As Decimal
        Get
            Return _c_Inafecto
        End Get
        Set
            _c_Inafecto = Value
        End Set
    End Property

    Public Property C_TasaIGV As Decimal
        Get
            Return _c_TasaIGV
        End Get
        Set
            _c_TasaIGV = Value
        End Set
    End Property

    Public Property C_ImporteISC As Decimal
        Get
            Return _c_ImporteISC
        End Get
        Set
            _c_ImporteISC = Value
        End Set
    End Property

    Public Property C_Detraccion As Boolean
        Get
            Return _c_Detraccion
        End Get
        Set
            _c_Detraccion = Value
        End Set
    End Property

    Public Property C_NroDetraccion As String
        Get
            Return _c_NroDetraccion
        End Get
        Set
            _c_NroDetraccion = Value
        End Set
    End Property

    Public Property C_FechaDetraccion As System.DateTime
        Get
            Return _c_FechaDetraccion
        End Get
        Set
            _c_FechaDetraccion = Value
        End Set
    End Property

    Public Property C_ImporteDetraccion As Decimal
        Get
            Return _c_ImporteDetraccion
        End Get
        Set
            _c_ImporteDetraccion = Value
        End Set
    End Property

    Public Property C_IdTipoAnexo As String
        Get
            Return _c_IdTipoAnexo
        End Get
        Set
            _c_IdTipoAnexo = Value
        End Set
    End Property

    Public Property C_ImporteMN As Decimal
        Get
            Return _c_ImporteMN
        End Get
        Set
            _c_ImporteMN = Value
        End Set
    End Property

    Public Property C_ImporteUS As Decimal
        Get
            Return _c_ImporteUS
        End Get
        Set
            _c_ImporteUS = Value
        End Set
    End Property

    Public Property C_IGVMN As Decimal
        Get
            Return _c_IGVMN
        End Get
        Set
            _c_IGVMN = Value
        End Set
    End Property

    Public Property C_IGVUS As Decimal
        Get
            Return _c_IGVUS
        End Get
        Set
            _c_IGVUS = Value
        End Set
    End Property

    Public Property C_TotalMN As Decimal
        Get
            Return _c_TotalMN
        End Get
        Set
            _c_TotalMN = Value
        End Set
    End Property

    Public Property C_TotalUS As Decimal
        Get
            Return _c_TotalUS
        End Get
        Set
            _c_TotalUS = Value
        End Set
    End Property

    Public Property C_ImporteISCMN As Decimal
        Get
            Return _c_ImporteISCMN
        End Get
        Set
            _c_ImporteISCMN = Value
        End Set
    End Property

    Public Property C_ImporteISCUS As Decimal
        Get
            Return _c_ImporteISCUS
        End Get
        Set
            _c_ImporteISCUS = Value
        End Set
    End Property

    Public Property C_Agencia_Aduana As String
        Get
            Return _c_Agencia_Aduana
        End Get
        Set
            _c_Agencia_Aduana = Value
        End Set
    End Property

    Public Property C_ImporteRta4ta As Decimal
        Get
            Return _c_ImporteRta4ta
        End Get
        Set
            _c_ImporteRta4ta = Value
        End Set
    End Property

    Public Property C_IdTipoAnexo1 As String
        Get
            Return _c_IdTipoAnexo1
        End Get
        Set
            _c_IdTipoAnexo1 = Value
        End Set
    End Property

    Public Property C_IdAnexo1 As String
        Get
            Return _c_IdAnexo1
        End Get
        Set
            _c_IdAnexo1 = Value
        End Set
    End Property

    Public Property C_IdVinculo_Residencia As String
        Get
            Return _c_IdVinculo_Residencia
        End Get
        Set
            _c_IdVinculo_Residencia = Value
        End Set
    End Property

    Public Property C_Tipo_Renta As String
        Get
            Return _c_Tipo_Renta
        End Get
        Set
            _c_Tipo_Renta = Value
        End Set
    End Property

    Public Property C_Mod_Serv_NoDomiciliado As String
        Get
            Return _c_Mod_Serv_NoDomiciliado
        End Get
        Set
            _c_Mod_Serv_NoDomiciliado = Value
        End Set
    End Property

    Public Property C_Ex_Op_Nodomic As String
        Get
            Return _c_Ex_Op_Nodomic
        End Get
        Set
            _c_Ex_Op_Nodomic = Value
        End Set
    End Property

    Public Property C_Tdref_import As String
        Get
            Return _c_Tdref_import
        End Get
        Set
            _c_Tdref_import = Value
        End Set
    End Property

    Public Property C_Serie_import As String
        Get
            Return _c_Serie_import
        End Get
        Set
            _c_Serie_import = Value
        End Set
    End Property

    Public Property C_Nroref_Import As String
        Get
            Return _c_Nroref_Import
        End Get
        Set
            _c_Nroref_Import = Value
        End Set
    End Property

    Public Property C_Fecha_Aduana As System.DateTime
        Get
            Return _c_Fecha_Aduana
        End Get
        Set
            _c_Fecha_Aduana = Value
        End Set
    End Property

    Public Property Bd As String
        Get
            Return _Bd
        End Get
        Set(value As String)
            _Bd = value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idSubdiario As String, ByVal nroComprobante As String, ByVal fechaComprobante As String, ByVal idMoneda As String, ByVal situacion As String, ByVal tipoCambio As Decimal, ByVal glosa As String, ByVal total As Decimal, ByVal idCambioMoneda As String, ByVal conConversion As String, ByVal fechaRegistro As System.DateTime, ByVal horaRegistro As String, ByVal usuario As String, ByVal fechaCambioMoneda As String, ByVal corig As String, ByVal cform As String, ByVal ctipcom As String, ByVal cextor As String, ByVal fechaComprobante2 As System.DateTime, ByVal fechaCambioMoneda2 As System.DateTime, ByVal flagCom As Boolean, ByVal tipoBien As String, ByVal c_IdAnexo As String, ByVal c_TipoDocumento As String, ByVal c_NroDocumento As String, ByVal c_FechaDocumento As System.DateTime, ByVal c_FechaVencimiento As System.DateTime, ByVal c_Importe As Decimal, ByVal c_Igv As Decimal, ByVal c_Total As Decimal, ByVal c_Inafecto As Decimal, ByVal c_TasaIGV As Decimal, ByVal c_ImporteISC As Decimal, ByVal c_Detraccion As Boolean, ByVal c_NroDetraccion As String, ByVal c_FechaDetraccion As System.DateTime, ByVal c_ImporteDetraccion As Decimal, ByVal c_IdTipoAnexo As String, ByVal c_ImporteMN As Decimal, ByVal c_ImporteUS As Decimal, ByVal c_IGVMN As Decimal, ByVal c_IGVUS As Decimal, ByVal c_TotalMN As Decimal, ByVal c_TotalUS As Decimal, ByVal c_ImporteISCMN As Decimal, ByVal c_ImporteISCUS As Decimal, ByVal c_Agencia_Aduana As String, ByVal c_ImporteRta4ta As Decimal, ByVal c_IdTipoAnexo1 As String, ByVal c_IdAnexo1 As String, ByVal c_IdVinculo_Residencia As String, ByVal c_Tipo_Renta As String, ByVal c_Mod_Serv_NoDomiciliado As String, ByVal c_Ex_Op_Nodomic As String, ByVal c_Tdref_import As String, ByVal c_Serie_import As String, ByVal c_Nroref_Import As String, ByVal c_Fecha_Aduana As System.DateTime)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NComprobante_Con)
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@fechaComprobante", "@idMoneda", "@situacion", "@tipoCambio", "@glosa", "@total", "@idCambioMoneda", "@conConversion", "@fechaRegistro", "@horaRegistro", "@usuario", "@fechaCambioMoneda", "@corig", "@cform", "@ctipcom", "@cextor", "@fechaComprobante2", "@fechaCambioMoneda2", "@flagCom", "@tipoBien", "@c_IdAnexo", "@c_TipoDocumento", "@c_NroDocumento", "@c_FechaDocumento", "@c_FechaVencimiento", "@c_Importe", "@c_Igv", "@c_Total", "@c_Inafecto", "@c_TasaIGV", "@c_ImporteISC", "@c_Detraccion", "@c_NroDetraccion", "@c_FechaDetraccion", "@c_ImporteDetraccion", "@c_IdTipoAnexo", "@c_Agencia_Aduana", "@c_ImporteRta4ta", "@c_IdTipoAnexo1", "@c_IdAnexo1", "@c_IdVinculo_Residencia", "@c_Tipo_Renta", "@c_Mod_Serv_NoDomiciliado", "@c_Ex_Op_Nodomic", "@c_Tdref_import", "@c_Serie_import", "@c_Nroref_Import", "@c_Fecha_Aduana"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.idSubdiario, d.NroComprobante, d.FechaComprobante, d.idMoneda, d.Situacion, d.TipoCambio, d.Glosa, d.Total, d.idCambioMoneda, d.ConConversion, d.FechaRegistro, d.HoraRegistro, d.Usuario, d.FechaCambioMoneda, d.corig, d.cform, d.ctipcom, d.cextor, d.FechaComprobante2, d.FechaCambioMoneda2, d.FlagCom, d.TipoBien, d.C_IdAnexo, d.C_TipoDocumento, d.C_NroDocumento, d.C_FechaDocumento, d.C_FechaVencimiento, d.C_Importe, d.C_Igv, d.C_Total, d.C_Inafecto, d.C_TasaIGV, d.C_ImporteISC, d.C_Detraccion, d.C_NroDetraccion, d.C_FechaDetraccion, d.C_ImporteDetraccion, d.C_IdTipoAnexo, d.C_Agencia_Aduana, d.C_ImporteRta4ta, d.C_IdTipoAnexo1, d.C_IdAnexo1, d.C_IdVinculo_Residencia, d.C_Tipo_Renta, d.C_Mod_Serv_NoDomiciliado, d.C_Ex_Op_Nodomic, d.C_Tdref_import, d.C_Serie_import, d.C_Nroref_Import, d.C_Fecha_Aduana}
        sql.EjecutarProcedure(d.Bd & ".dbo.Str_Comprobante_I", parametros, valores, tipoParametro, 50)
    End Sub
    Public Sub Actualizar(d As NComprobante_Con)
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@fechaComprobante", "@idMoneda", "@situacion", "@tipoCambio", "@glosa", "@total", "@idCambioMoneda", "@conConversion", "@fechaRegistro", "@horaRegistro", "@usuario", "@fechaCambioMoneda", "@corig", "@cform", "@ctipcom", "@cextor", "@fechaComprobante2", "@fechaCambioMoneda2", "@flagCom", "@tipoBien", "@c_IdAnexo", "@c_TipoDocumento", "@c_NroDocumento", "@c_FechaDocumento", "@c_FechaVencimiento", "@c_Importe", "@c_Igv", "@c_Total", "@c_Inafecto", "@c_TasaIGV", "@c_ImporteISC", "@c_Detraccion", "@c_NroDetraccion", "@c_FechaDetraccion", "@c_ImporteDetraccion", "@c_IdTipoAnexo", "@c_ImporteMN", "@c_ImporteUS", "@c_IGVMN", "@c_IGVUS", "@c_TotalMN", "@c_TotalUS", "@c_ImporteISCMN", "@c_ImporteISCUS", "@c_Agencia_Aduana", "@c_ImporteRta4ta", "@c_IdTipoAnexo1", "@c_IdAnexo1", "@c_IdVinculo_Residencia", "@c_Tipo_Renta", "@c_Mod_Serv_NoDomiciliado", "@c_Ex_Op_Nodomic", "@c_Tdref_import", "@c_Serie_import", "@c_Nroref_Import", "@c_Fecha_Aduana"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.idSubdiario, d.NroComprobante, d.FechaComprobante, d.idMoneda, d.Situacion, d.TipoCambio, d.Glosa, d.Total, d.idCambioMoneda, d.ConConversion, d.FechaRegistro, d.HoraRegistro, d.Usuario, d.FechaCambioMoneda, d.corig, d.cform, d.ctipcom, d.cextor, d.FechaComprobante2, d.FechaCambioMoneda2, d.FlagCom, d.TipoBien, d.C_IdAnexo, d.C_TipoDocumento, d.C_NroDocumento, d.C_FechaDocumento, d.C_FechaVencimiento, d.C_Importe, d.C_Igv, d.C_Total, d.C_Inafecto, d.C_TasaIGV, d.C_ImporteISC, d.C_Detraccion, d.C_NroDetraccion, d.C_FechaDetraccion, d.C_ImporteDetraccion, d.C_IdTipoAnexo, d.C_ImporteMN, d.C_ImporteUS, d.C_IGVMN, d.C_IGVUS, d.C_TotalMN, d.C_TotalUS, d.C_ImporteISCMN, d.C_ImporteISCUS, d.C_Agencia_Aduana, d.C_ImporteRta4ta, d.C_IdTipoAnexo1, d.C_IdAnexo1, d.C_IdVinculo_Residencia, d.C_Tipo_Renta, d.C_Mod_Serv_NoDomiciliado, d.C_Ex_Op_Nodomic, d.C_Tdref_import, d.C_Serie_import, d.C_Nroref_Import, d.C_Fecha_Aduana}
        sql.EjecutarProcedure(d.Bd & ".dbo.Str_Comprobante_U", parametros, valores, tipoParametro, 58)
    End Sub
    Public Sub Eliminar(d As NComprobante_Con)
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idSubdiario, d.NroComprobante}
        sql.EjecutarProcedure(d.Bd & ".dbo.Str_Comprobante_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Lista(d As NComprobante_Con) As DataTable
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(d.Bd & ".dbo.Str_Comprobante_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NComprobante_Con) As NComprobante_Con
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idSubdiario, d.NroComprobante}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(d.Bd & ".dbo.Str_Comprobante_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idSubdiario = IIf(dt.Rows(0).Item("idSubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idSubdiario"))
            d.NroComprobante = IIf(dt.Rows(0).Item("nroComprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroComprobante"))
            d.FechaComprobante = IIf(dt.Rows(0).Item("fechaComprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaComprobante"))
            d.idMoneda = IIf(dt.Rows(0).Item("idMoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMoneda"))
            d.Situacion = IIf(dt.Rows(0).Item("situacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("situacion"))
            d.TipoCambio = IIf(dt.Rows(0).Item("tipoCambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoCambio"))
            d.Glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.Total = IIf(dt.Rows(0).Item("total") Is DBNull.Value, Nothing, dt.Rows(0).Item("total"))
            d.idCambioMoneda = IIf(dt.Rows(0).Item("idCambioMoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCambioMoneda"))
            d.ConConversion = IIf(dt.Rows(0).Item("conConversion") Is DBNull.Value, Nothing, dt.Rows(0).Item("conConversion"))
            d.FechaRegistro = IIf(dt.Rows(0).Item("fechaRegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaRegistro"))
            d.HoraRegistro = IIf(dt.Rows(0).Item("horaRegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("horaRegistro"))
            d.Usuario = IIf(dt.Rows(0).Item("usuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuario"))
            d.FechaCambioMoneda = IIf(dt.Rows(0).Item("fechaCambioMoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCambioMoneda"))
            d.corig = IIf(dt.Rows(0).Item("corig") Is DBNull.Value, Nothing, dt.Rows(0).Item("corig"))
            d.cform = IIf(dt.Rows(0).Item("cform") Is DBNull.Value, Nothing, dt.Rows(0).Item("cform"))
            d.ctipcom = IIf(dt.Rows(0).Item("ctipcom") Is DBNull.Value, Nothing, dt.Rows(0).Item("ctipcom"))
            d.cextor = IIf(dt.Rows(0).Item("cextor") Is DBNull.Value, Nothing, dt.Rows(0).Item("cextor"))
            d.FechaComprobante2 = IIf(dt.Rows(0).Item("fechaComprobante2") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaComprobante2"))
            d.FechaCambioMoneda2 = IIf(dt.Rows(0).Item("fechaCambioMoneda2") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCambioMoneda2"))
            d.FlagCom = IIf(dt.Rows(0).Item("flagCom") Is DBNull.Value, Nothing, dt.Rows(0).Item("flagCom"))
            d.TipoBien = IIf(dt.Rows(0).Item("tipoBien") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoBien"))
            d.C_IdAnexo = IIf(dt.Rows(0).Item("c_IdAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_IdAnexo"))
            d.C_TipoDocumento = IIf(dt.Rows(0).Item("c_TipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_TipoDocumento"))
            d.C_NroDocumento = IIf(dt.Rows(0).Item("c_NroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_NroDocumento"))
            d.C_FechaDocumento = IIf(dt.Rows(0).Item("c_FechaDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_FechaDocumento"))
            d.C_FechaVencimiento = IIf(dt.Rows(0).Item("c_FechaVencimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_FechaVencimiento"))
            d.C_Importe = IIf(dt.Rows(0).Item("c_Importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Importe"))
            d.C_Igv = IIf(dt.Rows(0).Item("c_Igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Igv"))
            d.C_Total = IIf(dt.Rows(0).Item("c_Total") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Total"))
            d.C_Inafecto = IIf(dt.Rows(0).Item("c_Inafecto") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Inafecto"))
            d.C_TasaIGV = IIf(dt.Rows(0).Item("c_TasaIGV") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_TasaIGV"))
            d.C_ImporteISC = IIf(dt.Rows(0).Item("c_ImporteISC") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_ImporteISC"))
            d.C_Detraccion = IIf(dt.Rows(0).Item("c_Detraccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Detraccion"))
            d.C_NroDetraccion = IIf(dt.Rows(0).Item("c_NroDetraccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_NroDetraccion"))
            d.C_FechaDetraccion = IIf(dt.Rows(0).Item("c_FechaDetraccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_FechaDetraccion"))
            d.C_ImporteDetraccion = IIf(dt.Rows(0).Item("c_ImporteDetraccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_ImporteDetraccion"))
            d.C_IdTipoAnexo = IIf(dt.Rows(0).Item("c_IdTipoAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_IdTipoAnexo"))
            d.C_ImporteMN = IIf(dt.Rows(0).Item("c_ImporteMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_ImporteMN"))
            d.C_ImporteUS = IIf(dt.Rows(0).Item("c_ImporteUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_ImporteUS"))
            d.C_IGVMN = IIf(dt.Rows(0).Item("c_IGVMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_IGVMN"))
            d.C_IGVUS = IIf(dt.Rows(0).Item("c_IGVUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_IGVUS"))
            d.C_TotalMN = IIf(dt.Rows(0).Item("c_TotalMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_TotalMN"))
            d.C_TotalUS = IIf(dt.Rows(0).Item("c_TotalUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_TotalUS"))
            d.C_ImporteISCMN = IIf(dt.Rows(0).Item("c_ImporteISCMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_ImporteISCMN"))
            d.C_ImporteISCUS = IIf(dt.Rows(0).Item("c_ImporteISCUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_ImporteISCUS"))
            d.C_Agencia_Aduana = IIf(dt.Rows(0).Item("c_Agencia_Aduana") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Agencia_Aduana"))
            d.C_ImporteRta4ta = IIf(dt.Rows(0).Item("c_ImporteRta4ta") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_ImporteRta4ta"))
            d.C_IdTipoAnexo1 = IIf(dt.Rows(0).Item("c_IdTipoAnexo1") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_IdTipoAnexo1"))
            d.C_IdAnexo1 = IIf(dt.Rows(0).Item("c_IdAnexo1") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_IdAnexo1"))
            d.C_IdVinculo_Residencia = IIf(dt.Rows(0).Item("c_IdVinculo_Residencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_IdVinculo_Residencia"))
            d.C_Tipo_Renta = IIf(dt.Rows(0).Item("c_Tipo_Renta") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Tipo_Renta"))
            d.C_Mod_Serv_NoDomiciliado = IIf(dt.Rows(0).Item("c_Mod_Serv_NoDomiciliado") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Mod_Serv_NoDomiciliado"))
            d.C_Ex_Op_Nodomic = IIf(dt.Rows(0).Item("c_Ex_Op_Nodomic") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Ex_Op_Nodomic"))
            d.C_Tdref_import = IIf(dt.Rows(0).Item("c_Tdref_import") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Tdref_import"))
            d.C_Serie_import = IIf(dt.Rows(0).Item("c_Serie_import") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Serie_import"))
            d.C_Nroref_Import = IIf(dt.Rows(0).Item("c_Nroref_Import") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Nroref_Import"))
            d.C_Fecha_Aduana = IIf(dt.Rows(0).Item("c_Fecha_Aduana") Is DBNull.Value, Nothing, dt.Rows(0).Item("c_Fecha_Aduana"))
        Else
            d.FechaComprobante = Nothing
            d.idMoneda = Nothing
            d.Situacion = Nothing
            d.TipoCambio = Nothing
            d.Glosa = Nothing
            d.Total = Nothing
            d.idCambioMoneda = Nothing
            d.ConConversion = Nothing
            d.FechaRegistro = Nothing
            d.HoraRegistro = Nothing
            d.Usuario = Nothing
            d.FechaCambioMoneda = Nothing
            d.corig = Nothing
            d.cform = Nothing
            d.ctipcom = Nothing
            d.cextor = Nothing
            d.FechaComprobante2 = Nothing
            d.FechaCambioMoneda2 = Nothing
            d.FlagCom = Nothing
            d.TipoBien = Nothing
            d.C_IdAnexo = Nothing
            d.C_TipoDocumento = Nothing
            d.C_NroDocumento = Nothing
            d.C_FechaDocumento = Nothing
            d.C_FechaVencimiento = Nothing
            d.C_Importe = Nothing
            d.C_Igv = Nothing
            d.C_Total = Nothing
            d.C_Inafecto = Nothing
            d.C_TasaIGV = Nothing
            d.C_ImporteISC = Nothing
            d.C_Detraccion = Nothing
            d.C_NroDetraccion = Nothing
            d.C_FechaDetraccion = Nothing
            d.C_ImporteDetraccion = Nothing
            d.C_IdTipoAnexo = Nothing
            d.C_ImporteMN = Nothing
            d.C_ImporteUS = Nothing
            d.C_IGVMN = Nothing
            d.C_IGVUS = Nothing
            d.C_TotalMN = Nothing
            d.C_TotalUS = Nothing
            d.C_ImporteISCMN = Nothing
            d.C_ImporteISCUS = Nothing
            d.C_Agencia_Aduana = Nothing
            d.C_ImporteRta4ta = Nothing
            d.C_IdTipoAnexo1 = Nothing
            d.C_IdAnexo1 = Nothing
            d.C_IdVinculo_Residencia = Nothing
            d.C_Tipo_Renta = Nothing
            d.C_Mod_Serv_NoDomiciliado = Nothing
            d.C_Ex_Op_Nodomic = Nothing
            d.C_Tdref_import = Nothing
            d.C_Serie_import = Nothing
            d.C_Nroref_Import = Nothing
            d.C_Fecha_Aduana = Nothing
        End If
        Return d
    End Function

    Public Function Existe_Comprobante(d As NComprobante_Con)
        Dim parametros() As Object = {"@idsubdiario", "@nrocomprobante"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idSubdiario, d.NroComprobante}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar(d.Bd & ".dbo.Existe_Comprobante", parametros, valores, tipoParametro, 2)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
#End Region

End Class

