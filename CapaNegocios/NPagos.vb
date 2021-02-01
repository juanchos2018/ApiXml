Imports CapaDatos
Public Class NPagos
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _idAgencia As String
    Private _idTipoDocumento As String
    Private _serie As String
    Private _numeroDocumento As String
    Private _debeHaber As String
    Private _estado As String
    Private _glosa As String
    Private _tipoCambio As Decimal
    Private _pagoUS As Decimal
    Private _pagoMN As Decimal
    Private _importeTarjetaUS As Decimal
    Private _importeTarjetaMN As Decimal
    Private _numeroTarjeta As String
    Private _codigoTarjeta As String
    Private _adelanto As Decimal
    Private _ajuste As Decimal
    Private _codigoBanco As String
    Private _importeChequeUS As Decimal
    Private _importeChequeMN As Decimal
    Private _numeroCheque As String
    Private _usuarioCrea As String
    Private _fechaCrea As System.DateTime
    Private _usuarioMod As String
    Private _fechaMod As System.DateTime
    Private _idMoneda As String
    Private _importeEfectivo As Decimal
    Private _importeTarjeta As Decimal
    Private _importeBanco As Decimal

#End Region

#Region "Properties"

    Public Property IdAgencia As String
        Get
            Return _idAgencia
        End Get
        Set
            _idAgencia = Value
        End Set
    End Property

    Public Property IdTipoDocumento As String
        Get
            Return _idTipoDocumento
        End Get
        Set
            _idTipoDocumento = Value
        End Set
    End Property

    Public Property Serie As String
        Get
            Return _serie
        End Get
        Set
            _serie = Value
        End Set
    End Property

    Public Property NumeroDocumento As String
        Get
            Return _numeroDocumento
        End Get
        Set
            _numeroDocumento = Value
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

    Public Property Estado As String
        Get
            Return _estado
        End Get
        Set
            _estado = Value
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

    Public Property TipoCambio As Decimal
        Get
            Return _tipoCambio
        End Get
        Set
            _tipoCambio = Value
        End Set
    End Property

    Public Property PagoUS As Decimal
        Get
            Return _pagoUS
        End Get
        Set
            _pagoUS = Value
        End Set
    End Property

    Public Property PagoMN As Decimal
        Get
            Return _pagoMN
        End Get
        Set
            _pagoMN = Value
        End Set
    End Property

    Public Property ImporteTarjetaUS As Decimal
        Get
            Return _importeTarjetaUS
        End Get
        Set
            _importeTarjetaUS = Value
        End Set
    End Property

    Public Property ImporteTarjetaMN As Decimal
        Get
            Return _importeTarjetaMN
        End Get
        Set
            _importeTarjetaMN = Value
        End Set
    End Property

    Public Property NumeroTarjeta As String
        Get
            Return _numeroTarjeta
        End Get
        Set
            _numeroTarjeta = Value
        End Set
    End Property

    Public Property CodigoTarjeta As String
        Get
            Return _codigoTarjeta
        End Get
        Set
            _codigoTarjeta = Value
        End Set
    End Property

    Public Property Adelanto As Decimal
        Get
            Return _adelanto
        End Get
        Set
            _adelanto = Value
        End Set
    End Property

    Public Property Ajuste As Decimal
        Get
            Return _ajuste
        End Get
        Set
            _ajuste = Value
        End Set
    End Property

    Public Property CodigoBanco As String
        Get
            Return _codigoBanco
        End Get
        Set
            _codigoBanco = Value
        End Set
    End Property

    Public Property ImporteChequeUS As Decimal
        Get
            Return _importeChequeUS
        End Get
        Set
            _importeChequeUS = Value
        End Set
    End Property

    Public Property ImporteChequeMN As Decimal
        Get
            Return _importeChequeMN
        End Get
        Set
            _importeChequeMN = Value
        End Set
    End Property

    Public Property NumeroCheque As String
        Get
            Return _numeroCheque
        End Get
        Set
            _numeroCheque = Value
        End Set
    End Property

    Public Property UsuarioCrea As String
        Get
            Return _usuarioCrea
        End Get
        Set
            _usuarioCrea = Value
        End Set
    End Property

    Public Property FechaCrea As System.DateTime
        Get
            Return _fechaCrea
        End Get
        Set
            _fechaCrea = Value
        End Set
    End Property

    Public Property UsuarioMod As String
        Get
            Return _usuarioMod
        End Get
        Set
            _usuarioMod = Value
        End Set
    End Property

    Public Property FechaMod As System.DateTime
        Get
            Return _fechaMod
        End Get
        Set
            _fechaMod = Value
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
    Public Property ImporteEfectivo As Decimal
        Get
            Return _importeEfectivo
        End Get
        Set(value As Decimal)
            _importeEfectivo = value
        End Set
    End Property
    Public Property ImporteTarjeta As Decimal
        Get
            Return _importeTarjeta
        End Get
        Set(value As Decimal)
            _importeTarjeta = value
        End Set
    End Property
    Public Property ImporteBanco As Decimal
        Get
            Return _importeBanco
        End Get
        Set(value As Decimal)
            _importeBanco = value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idAgencia As String, ByVal idTipoDocumento As String, ByVal serie As String, ByVal numeroDocumento As String, ByVal debeHaber As String, ByVal estado As String, ByVal glosa As String, ByVal tipoCambio As Decimal, ByVal pagoUS As Decimal, ByVal pagoMN As Decimal, ByVal importeTarjetaUS As Decimal, ByVal importeTarjetaMN As Decimal, ByVal numeroTarjeta As String, ByVal codigoTarjeta As String, ByVal adelanto As Decimal, ByVal ajuste As Decimal, ByVal codigoBanco As String, ByVal importeChequeUS As Decimal, ByVal importeChequeMN As Decimal, ByVal numeroCheque As String, ByVal f7_CVENDE As String, ByVal f7_CNROCAJ As String, ByVal usuarioCrea As String, ByVal fechaCrea As System.DateTime, ByVal usuarioMod As String, ByVal fechaMod As System.DateTime, ByVal f7_CVENDE2 As String, ByVal f7_NPORVEN As Decimal, ByVal f7_NPORVE2 As Decimal, ByVal f7_CBOLBAN As String, ByVal f7_CBOLNUM As String, ByVal f7_NIUSBOL As Decimal, ByVal f7_NIMNBOL As Decimal, ByVal f7_DFECDOC As System.DateTime, ByVal f7_CCODBCO As String, ByVal f7_CNUMDE1 As String, ByVal f7_NDEPMN1 As Decimal, ByVal f7_NDEPUS1 As Decimal, ByVal f7_CNUMDE2 As String, ByVal f7_NDEPMN2 As Decimal, ByVal f7_NDEPUS2 As Decimal, ByVal f7_CNUMDE3 As String, ByVal f7_NDEPMN3 As Decimal, ByVal f7_NDEPUS3 As Decimal, ByVal f7_NREDOUS As Decimal, ByVal f7_NREDOMN As Decimal, ByVal f7_CCAJA As String, ByVal f7_N2TIMUS As Decimal, ByVal f7_N2TIMMN As Decimal, ByVal f7_C2TARCO As String, ByVal f7_C2TARNU As String, ByVal f7_N3TIMUS As Decimal, ByVal f7_N3TIMMN As Decimal, ByVal f7_C3TARCO As String, ByVal f7_C3TARNU As String, ByVal idMoneda As String)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    Public Sub agregar(p As NPagos)
        Dim params1() As Object = {"@IdAgencia", "@IdTipoDocumento", "@Serie", "@NumeroDocumento", "@DebeHaber", "@Glosa",
            "@TipoCambio", "@TipoMoneda", "@ImporteEfectivo", "@ImporteTarjeta", "@NumeroTarjeta", "@CodigoTarjeta", "@CodigoBanco", "@ImporteCheque", "@NumeroCheque", "@Usuario", "@Fecha"}
        Dim tipoParametro1() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Float, SqlDbType.Char, SqlDbType.Float, SqlDbType.Float, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime}
        Dim vals1() As Object = {
            p.IdAgencia, p.IdTipoDocumento, p.Serie, p.NumeroDocumento, p.DebeHaber, p.Glosa,
            p.TipoCambio, p.IdMoneda, p.ImporteEfectivo, p.ImporteTarjeta, p.NumeroTarjeta, p.CodigoTarjeta, p.CodigoBanco, p.ImporteBanco, p.NumeroCheque, p.UsuarioCrea, p.FechaCrea}
        sql.EjecutarProcedure("proc_AddPago", params1, vals1, tipoParametro1, 17)
    End Sub

#End Region

End Class
