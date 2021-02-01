Imports CapaDatos
Public Class NDetComprobante_Con
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _idSubdiario As String
    Private _nroComprobante As String
    Private _secuencia As String
    Private _fechaComprobante As String
    Private _idCuenta As String
    Private _idAnexo As String
    Private _idCentroCosto As String
    Private _idMoneda As String
    Private _debeHaber As String
    Private _importe As Decimal
    Private _tipoDocumento As String
    Private _nroDocumento As String
    Private _fechaDocumento As String
    Private _fechaVencimiento As String
    Private _idArea As String
    Private _conConversion As String
    Private _fechaRegistro As System.DateTime
    Private _glosa As String
    Private _importeUS As Decimal
    Private _importeMN As Decimal
    Private _dcodarc As String
    Private _fechaComprobante2 As System.DateTime
    Private _fechaDocumento2 As System.DateTime
    Private _fechaVencimiento2 As System.DateTime
    Private _idAnexo2 As String
    Private _idTipoAnexo As String
    Private _idTipoAnexo2 As String
    Private _tipoCambio As Decimal
    Private _dcantid As Decimal
    Private _drete As String
    Private _dporre As Decimal
    Private _dtipdor As String
    Private _dnumdor As String
    Private _dfecdo2 As System.DateTime
    Private _dtiptas As String
    Private _dimptas As Decimal
    Private _dimpbmn As Decimal
    Private _dimpbus As Decimal
    Private _medioPago As String
    Private _Bd As String

#End Region

#Region "Properties"

    Public Property IdSubdiario As String
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

    Public Property Secuencia As String
        Get
            Return _secuencia
        End Get
        Set
            _secuencia = Value
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

    Public Property IdCuenta As String
        Get
            Return _idCuenta
        End Get
        Set
            _idCuenta = Value
        End Set
    End Property

    Public Property IdAnexo As String
        Get
            Return _idAnexo
        End Get
        Set
            _idAnexo = Value
        End Set
    End Property

    Public Property IdCentroCosto As String
        Get
            Return _idCentroCosto
        End Get
        Set
            _idCentroCosto = Value
        End Set
    End Property

    Public Property IdMoneda As String
        Get
            Return _idMoneda
        End Get
        Set
            _idMoneda = Value
        End Set
    End Property

    Public Property DebeHaber As String
        Get
            Return _debeHaber
        End Get
        Set
            _debeHaber = Value
        End Set
    End Property

    Public Property Importe As Decimal
        Get
            Return _importe
        End Get
        Set
            _importe = Value
        End Set
    End Property

    Public Property TipoDocumento As String
        Get
            Return _tipoDocumento
        End Get
        Set
            _tipoDocumento = Value
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

    Public Property FechaDocumento As String
        Get
            Return _fechaDocumento
        End Get
        Set
            _fechaDocumento = Value
        End Set
    End Property

    Public Property FechaVencimiento As String
        Get
            Return _fechaVencimiento
        End Get
        Set
            _fechaVencimiento = Value
        End Set
    End Property

    Public Property IdArea As String
        Get
            Return _idArea
        End Get
        Set
            _idArea = Value
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

    Public Property Glosa As String
        Get
            Return _glosa
        End Get
        Set
            _glosa = Value
        End Set
    End Property

    Public Property ImporteUS As Decimal
        Get
            Return _importeUS
        End Get
        Set
            _importeUS = Value
        End Set
    End Property

    Public Property ImporteMN As Decimal
        Get
            Return _importeMN
        End Get
        Set
            _importeMN = Value
        End Set
    End Property

    Public Property dcodarc As String
        Get
            Return _dcodarc
        End Get
        Set
            _dcodarc = Value
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

    Public Property FechaDocumento2 As System.DateTime
        Get
            Return _fechaDocumento2
        End Get
        Set
            _fechaDocumento2 = Value
        End Set
    End Property

    Public Property FechaVencimiento2 As System.DateTime
        Get
            Return _fechaVencimiento2
        End Get
        Set
            _fechaVencimiento2 = Value
        End Set
    End Property

    Public Property IdAnexo2 As String
        Get
            Return _idAnexo2
        End Get
        Set
            _idAnexo2 = Value
        End Set
    End Property

    Public Property IdTipoAnexo As String
        Get
            Return _idTipoAnexo
        End Get
        Set
            _idTipoAnexo = Value
        End Set
    End Property

    Public Property IdTipoAnexo2 As String
        Get
            Return _idTipoAnexo2
        End Get
        Set
            _idTipoAnexo2 = Value
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

    Public Property dcantid As Decimal
        Get
            Return _dcantid
        End Get
        Set
            _dcantid = Value
        End Set
    End Property

    Public Property drete As String
        Get
            Return _drete
        End Get
        Set
            _drete = Value
        End Set
    End Property

    Public Property dporre As Decimal
        Get
            Return _dporre
        End Get
        Set
            _dporre = Value
        End Set
    End Property

    Public Property dtipdor As String
        Get
            Return _dtipdor
        End Get
        Set
            _dtipdor = Value
        End Set
    End Property

    Public Property dnumdor As String
        Get
            Return _dnumdor
        End Get
        Set
            _dnumdor = Value
        End Set
    End Property

    Public Property dfecdo2 As System.DateTime
        Get
            Return _dfecdo2
        End Get
        Set
            _dfecdo2 = Value
        End Set
    End Property

    Public Property dtiptas As String
        Get
            Return _dtiptas
        End Get
        Set
            _dtiptas = Value
        End Set
    End Property

    Public Property dimptas As Decimal
        Get
            Return _dimptas
        End Get
        Set
            _dimptas = Value
        End Set
    End Property

    Public Property dimpbmn As Decimal
        Get
            Return _dimpbmn
        End Get
        Set
            _dimpbmn = Value
        End Set
    End Property

    Public Property dimpbus As Decimal
        Get
            Return _dimpbus
        End Get
        Set
            _dimpbus = Value
        End Set
    End Property

    Public Property MedioPago As String
        Get
            Return _medioPago
        End Get
        Set
            _medioPago = Value
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

    Public Sub New(ByVal idSubdiario As String, ByVal nroComprobante As String, ByVal secuencia As String, ByVal fechaComprobante As String, ByVal idCuenta As String, ByVal idAnexo As String, ByVal idCentroCosto As String, ByVal idMoneda As String, ByVal debeHaber As String, ByVal importe As Decimal, ByVal tipoDocumento As String, ByVal nroDocumento As String, ByVal fechaDocumento As String, ByVal fechaVencimiento As String, ByVal idArea As String, ByVal conConversion As String, ByVal fechaRegistro As System.DateTime, ByVal glosa As String, ByVal importeUS As Decimal, ByVal importeMN As Decimal, ByVal dcodarc As String, ByVal fechaComprobante2 As System.DateTime, ByVal fechaDocumento2 As System.DateTime, ByVal fechaVencimiento2 As System.DateTime, ByVal idAnexo2 As String, ByVal idTipoAnexo As String, ByVal idTipoAnexo2 As String, ByVal tipoCambio As Decimal, ByVal dcantid As Decimal, ByVal drete As String, ByVal dporre As Decimal, ByVal dtipdor As String, ByVal dnumdor As String, ByVal dfecdo2 As System.DateTime, ByVal dtiptas As String, ByVal dimptas As Decimal, ByVal dimpbmn As Decimal, ByVal dimpbus As Decimal, ByVal medioPago As String)
        Me.New()
    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NDetComprobante_Con)
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@secuencia", "@fechaComprobante", "@idCuenta", "@idAnexo", "@idCentroCosto", "@idMoneda", "@debeHaber", "@importe", "@tipoDocumento", "@nroDocumento", "@fechaDocumento", "@fechaVencimiento", "@idArea", "@conConversion", "@fechaRegistro", "@glosa", "@importeUS", "@importeMN", "@dcodarc", "@fechaComprobante2", "@fechaDocumento2", "@fechaVencimiento2", "@idAnexo2", "@idTipoAnexo", "@idTipoAnexo2", "@tipoCambio", "@dcantid", "@drete", "@dporre", "@dtipdor", "@dnumdor", "@dfecdo2", "@dtiptas", "@dimptas", "@dimpbmn", "@dimpbus", "@medioPago"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdSubdiario, d.NroComprobante, d.Secuencia, d.FechaComprobante, d.IdCuenta, d.IdAnexo, d.IdCentroCosto, d.IdMoneda, d.DebeHaber, d.Importe, d.TipoDocumento, d.NroDocumento, d.FechaDocumento, d.FechaVencimiento, d.IdArea, d.ConConversion, d.FechaRegistro, d.Glosa, d.ImporteUS, d.ImporteMN, d.dcodarc, d.FechaComprobante2, d.FechaDocumento2, d.FechaVencimiento2, d.IdAnexo2, d.IdTipoAnexo, d.IdTipoAnexo2, d.TipoCambio, d.dcantid, d.drete, d.dporre, d.dtipdor, d.dnumdor, d.dfecdo2, d.dtiptas, d.dimptas, d.dimpbmn, d.dimpbus, d.MedioPago}
        sql.EjecutarProcedure(Bd & ".dbo.Str_DetComprobante_I", parametros, valores, tipoParametro, 39)
    End Sub
    Public Sub Actualizar(d As NDetComprobante_Con)
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@secuencia", "@fechaComprobante", "@idCuenta", "@idAnexo", "@idCentroCosto", "@idMoneda", "@debeHaber", "@importe", "@tipoDocumento", "@nroDocumento", "@fechaDocumento", "@fechaVencimiento", "@idArea", "@conConversion", "@fechaRegistro", "@glosa", "@importeUS", "@importeMN", "@dcodarc", "@fechaComprobante2", "@fechaDocumento2", "@fechaVencimiento2", "@idAnexo2", "@idTipoAnexo", "@idTipoAnexo2", "@tipoCambio", "@dcantid", "@drete", "@dporre", "@dtipdor", "@dnumdor", "@dfecdo2", "@dtiptas", "@dimptas", "@dimpbmn", "@dimpbus", "@medioPago"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdSubdiario, d.NroComprobante, d.Secuencia, d.FechaComprobante, d.IdCuenta, d.IdAnexo, d.IdCentroCosto, d.IdMoneda, d.DebeHaber, d.Importe, d.TipoDocumento, d.NroDocumento, d.FechaDocumento, d.FechaVencimiento, d.IdArea, d.ConConversion, d.FechaRegistro, d.Glosa, d.ImporteUS, d.ImporteMN, d.dcodarc, d.FechaComprobante2, d.FechaDocumento2, d.FechaVencimiento2, d.IdAnexo2, d.IdTipoAnexo, d.IdTipoAnexo2, d.TipoCambio, d.dcantid, d.drete, d.dporre, d.dtipdor, d.dnumdor, d.dfecdo2, d.dtiptas, d.dimptas, d.dimpbmn, d.dimpbus, d.MedioPago}
        sql.EjecutarProcedure(Bd & ".dbo.Str_DetComprobante_U", parametros, valores, tipoParametro, 39)
    End Sub
    Public Sub Eliminar(d As NDetComprobante_Con)
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@secuencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdSubdiario, d.NroComprobante, d.Secuencia}
        sql.EjecutarProcedure(Bd & ".dbo.Str_DetComprobante_D", parametros, valores, tipoParametro, 3)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@secuencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_DetComprobante_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDetComprobante_Con) As NDetComprobante_Con
        Dim parametros() As Object = {"@idSubdiario", "@nroComprobante", "@secuencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdSubdiario, d.NroComprobante, d.Secuencia}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_DetComprobante_S", parametros, valores, tipoParametro, 39).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdSubdiario = IIf(dt.Rows(0).Item("idSubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idSubdiario"))
            d.NroComprobante = IIf(dt.Rows(0).Item("nroComprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroComprobante"))
            d.Secuencia = IIf(dt.Rows(0).Item("secuencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("secuencia"))
            d.FechaComprobante = IIf(dt.Rows(0).Item("fechaComprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaComprobante"))
            d.IdCuenta = IIf(dt.Rows(0).Item("idCuenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCuenta"))
            d.IdAnexo = IIf(dt.Rows(0).Item("idAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAnexo"))
            d.IdCentroCosto = IIf(dt.Rows(0).Item("idCentroCosto") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCentroCosto"))
            d.IdMoneda = IIf(dt.Rows(0).Item("idMoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMoneda"))
            d.DebeHaber = IIf(dt.Rows(0).Item("debeHaber") Is DBNull.Value, Nothing, dt.Rows(0).Item("debeHaber"))
            d.Importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))
            d.TipoDocumento = IIf(dt.Rows(0).Item("tipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoDocumento"))
            d.NroDocumento = IIf(dt.Rows(0).Item("nroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroDocumento"))
            d.FechaDocumento = IIf(dt.Rows(0).Item("fechaDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaDocumento"))
            d.FechaVencimiento = IIf(dt.Rows(0).Item("fechaVencimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaVencimiento"))
            d.IdArea = IIf(dt.Rows(0).Item("idArea") Is DBNull.Value, Nothing, dt.Rows(0).Item("idArea"))
            d.ConConversion = IIf(dt.Rows(0).Item("conConversion") Is DBNull.Value, Nothing, dt.Rows(0).Item("conConversion"))
            d.FechaRegistro = IIf(dt.Rows(0).Item("fechaRegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaRegistro"))
            d.Glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.ImporteUS = IIf(dt.Rows(0).Item("importeUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeUS"))
            d.ImporteMN = IIf(dt.Rows(0).Item("importeMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeMN"))
            d.dcodarc = IIf(dt.Rows(0).Item("dcodarc") Is DBNull.Value, Nothing, dt.Rows(0).Item("dcodarc"))
            d.FechaComprobante2 = IIf(dt.Rows(0).Item("fechaComprobante2") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaComprobante2"))
            d.FechaDocumento2 = IIf(dt.Rows(0).Item("fechaDocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaDocumento2"))
            d.FechaVencimiento2 = IIf(dt.Rows(0).Item("fechaVencimiento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaVencimiento2"))
            d.IdAnexo2 = IIf(dt.Rows(0).Item("idAnexo2") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAnexo2"))
            d.IdTipoAnexo = IIf(dt.Rows(0).Item("idTipoAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoAnexo"))
            d.IdTipoAnexo2 = IIf(dt.Rows(0).Item("idTipoAnexo2") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoAnexo2"))
            d.TipoCambio = IIf(dt.Rows(0).Item("tipoCambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoCambio"))
            d.dcantid = IIf(dt.Rows(0).Item("dcantid") Is DBNull.Value, Nothing, dt.Rows(0).Item("dcantid"))
            d.drete = IIf(dt.Rows(0).Item("drete") Is DBNull.Value, Nothing, dt.Rows(0).Item("drete"))
            d.dporre = IIf(dt.Rows(0).Item("dporre") Is DBNull.Value, Nothing, dt.Rows(0).Item("dporre"))
            d.dtipdor = IIf(dt.Rows(0).Item("dtipdor") Is DBNull.Value, Nothing, dt.Rows(0).Item("dtipdor"))
            d.dnumdor = IIf(dt.Rows(0).Item("dnumdor") Is DBNull.Value, Nothing, dt.Rows(0).Item("dnumdor"))
            d.dfecdo2 = IIf(dt.Rows(0).Item("dfecdo2") Is DBNull.Value, Nothing, dt.Rows(0).Item("dfecdo2"))
            d.dtiptas = IIf(dt.Rows(0).Item("dtiptas") Is DBNull.Value, Nothing, dt.Rows(0).Item("dtiptas"))
            d.dimptas = IIf(dt.Rows(0).Item("dimptas") Is DBNull.Value, Nothing, dt.Rows(0).Item("dimptas"))
            d.dimpbmn = IIf(dt.Rows(0).Item("dimpbmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("dimpbmn"))
            d.dimpbus = IIf(dt.Rows(0).Item("dimpbus") Is DBNull.Value, Nothing, dt.Rows(0).Item("dimpbus"))
            d.MedioPago = IIf(dt.Rows(0).Item("medioPago") Is DBNull.Value, Nothing, dt.Rows(0).Item("medioPago"))
        Else
            d.FechaComprobante = Nothing
            d.IdCuenta = Nothing
            d.IdAnexo = Nothing
            d.IdCentroCosto = Nothing
            d.IdMoneda = Nothing
            d.DebeHaber = Nothing
            d.Importe = Nothing
            d.TipoDocumento = Nothing
            d.NroDocumento = Nothing
            d.FechaDocumento = Nothing
            d.FechaVencimiento = Nothing
            d.IdArea = Nothing
            d.ConConversion = Nothing
            d.FechaRegistro = Nothing
            d.Glosa = Nothing
            d.ImporteUS = Nothing
            d.ImporteMN = Nothing
            d.dcodarc = Nothing
            d.FechaComprobante2 = Nothing
            d.FechaDocumento2 = Nothing
            d.FechaVencimiento2 = Nothing
            d.IdAnexo2 = Nothing
            d.IdTipoAnexo = Nothing
            d.IdTipoAnexo2 = Nothing
            d.TipoCambio = Nothing
            d.dcantid = Nothing
            d.drete = Nothing
            d.dporre = Nothing
            d.dtipdor = Nothing
            d.dnumdor = Nothing
            d.dfecdo2 = Nothing
            d.dtiptas = Nothing
            d.dimptas = Nothing
            d.dimpbmn = Nothing
            d.dimpbus = Nothing
            d.MedioPago = Nothing
        End If
        Return d
    End Function
#End Region

End Class