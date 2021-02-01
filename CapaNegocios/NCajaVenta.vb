Imports CapaDatos
Public Class NCajaVenta
    Dim sql As New ClsConexion
#Region "Declarations"
    Private _idCajaVenta As Integer
    Private _idAgencia As String
    Private _idAlmacen As String
    Private _idTipoDocumento As String
    Private _serie As String
    Private _numeroDocumento As String
    Private _debeHaber As String
    Private _estado As String
    Private _glosa As String
    Private _tipoCambio As Decimal
    Private _idMoneda As String
    Private _pago As Decimal
    Private _pagoUS As Decimal
    Private _pagoMN As Decimal
    Private _importeTarjeta As Decimal
    Private _importeTarjetaUS As Decimal
    Private _importeTarjetaMN As Decimal
    Private _numeroTarjeta As String
    Private _codigoTarjeta As String
    Private _adelanto As Decimal
    Private _ajuste As Decimal
    Private _codigoBanco As String
    Private _importeCheque As Decimal
    Private _importeChequeUS As Decimal
    Private _importeChequeMN As Decimal
    Private _numeroCheque As String
    Private _usuarioCrea As String
    Private _fechaCrea As System.DateTime
    Private _usuarioMod As String
    Private _fechaMod As System.DateTime
    Private _idDetUsuarioCaja As Integer
    Private _nroDocumentoPago As String
    Private _tipoDocumentoPago As String
    Private _idMonedaPago As String
    Private _FechaDocumento As DateTime
    Public Property IdSubdiario As String
    Public Property NroComprobante As String
    Public Property Secuencia As String

#End Region
#Region "Properties"

    Public Property IdCajaVenta As Integer
        Get
            Return _idCajaVenta
        End Get
        Set
            _idCajaVenta = Value
        End Set
    End Property

    Public Property IdAgencia As String
        Get
            Return _idAgencia
        End Get
        Set
            _idAgencia = Value
        End Set
    End Property

    Public Property IdAlmacen As String
        Get
            Return _idAlmacen
        End Get
        Set
            _idAlmacen = Value
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

    Public Property IdMoneda As String
        Get
            Return _idMoneda
        End Get
        Set
            _idMoneda = Value
        End Set
    End Property

    Public Property Pago As Decimal
        Get
            Return _pago
        End Get
        Set
            _pago = Value
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

    Public Property ImporteTarjeta As Decimal
        Get
            Return _importeTarjeta
        End Get
        Set
            _importeTarjeta = Value
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

    Public Property ImporteCheque As Decimal
        Get
            Return _importeCheque
        End Get
        Set
            _importeCheque = Value
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

    Public Property IdDetUsuarioCaja As Integer
        Get
            Return _idDetUsuarioCaja
        End Get
        Set
            _idDetUsuarioCaja = Value
        End Set
    End Property

    Public Property NroDocumentoPago As String
        Get
            Return _nroDocumentoPago
        End Get
        Set
            _nroDocumentoPago = Value
        End Set
    End Property

    Public Property TipoDocumentoPago As String
        Get
            Return _tipoDocumentoPago
        End Get
        Set
            _tipoDocumentoPago = Value
        End Set
    End Property

    Public Property IdMonedaPago As String
        Get
            Return _idMonedaPago
        End Get
        Set
            _idMonedaPago = Value
        End Set
    End Property

    Public Property FechaDocumento As Date
        Get
            Return _FechaDocumento
        End Get
        Set(value As Date)
            _FechaDocumento = value
        End Set
    End Property


#End Region
#Region "Constructors"
    Public Sub New()
    End Sub
    Public Sub New(ByVal idCajaVenta As Integer, ByVal idAgencia As String, ByVal idAlmacen As String, ByVal idTipoDocumento As String, ByVal serie As String, ByVal numeroDocumento As String, ByVal debeHaber As String, ByVal estado As String, ByVal glosa As String, ByVal tipoCambio As Decimal, ByVal idMoneda As String, ByVal pago As Decimal, ByVal pagoUS As Decimal, ByVal pagoMN As Decimal, ByVal importeTarjeta As Decimal, ByVal importeTarjetaUS As Decimal, ByVal importeTarjetaMN As Decimal, ByVal numeroTarjeta As String, ByVal codigoTarjeta As String, ByVal adelanto As Decimal, ByVal ajuste As Decimal, ByVal codigoBanco As String, ByVal importeCheque As Decimal, ByVal importeChequeUS As Decimal, ByVal importeChequeMN As Decimal, ByVal numeroCheque As String, ByVal usuarioCrea As String, ByVal fechaCrea As System.DateTime, ByVal usuarioMod As String, ByVal fechaMod As System.DateTime, ByVal idDetUsuarioCaja As Integer, ByVal nroDocumentoPago As String, ByVal tipoDocumentoPago As String, ByVal idMonedaPago As String)
        Me.New()
    End Sub
#End Region
#Region "Metodos"
    Public Sub Agregar(d As NCajaVenta)
        Dim parametros() As Object = {"@idAgencia", "@idAlmacen", "@idTipoDocumento", "@serie", "@numeroDocumento", "@debeHaber", "@estado", "@glosa", "@tipoCambio", "@idMoneda", "@pago", "@importeTarjeta", "@numeroTarjeta", "@codigoTarjeta", "@adelanto", "@ajuste", "@codigoBanco", "@importeCheque", "@numeroCheque", "@usuarioCrea", "@fechaCrea", "@usuarioMod", "@fechaMod", "@idDetUsuarioCaja", "@nroDocumentoPago", "@tipoDocumentoPago", "@idMonedaPago", "@FechaDocumento", "@IdSubdiario", "@NroComprobante", "@Secuencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAgencia, d.IdAlmacen, d.IdTipoDocumento, d.Serie, d.NumeroDocumento, d.DebeHaber, d.Estado, d.Glosa, d.TipoCambio, d.IdMoneda, d.Pago, d.ImporteTarjeta, d.NumeroTarjeta, d.CodigoTarjeta, d.Adelanto, d.Ajuste, d.CodigoBanco, d.ImporteCheque, d.NumeroCheque, d.UsuarioCrea, d.FechaCrea, d.UsuarioMod, d.FechaMod, d.IdDetUsuarioCaja, d.NroDocumentoPago, d.TipoDocumentoPago, d.IdMonedaPago, d.FechaDocumento, d.IdSubdiario, d.NroComprobante, d.Secuencia}
        sql.EjecutarProcedure("Str_Tbl_Caja_Venta_I", parametros, valores, tipoParametro, 31)
    End Sub
    Public Sub Actualizar(d As NCajaVenta)
        Dim parametros() As Object = {"@idCajaVenta", "@idAgencia", "@idAlmacen", "@idTipoDocumento", "@serie", "@numeroDocumento", "@debeHaber", "@estado", "@glosa", "@tipoCambio", "@idMoneda", "@pago", "@importeTarjeta", "@numeroTarjeta", "@codigoTarjeta", "@adelanto", "@ajuste", "@codigoBanco", "@importeCheque", "@numeroCheque", "@usuarioCrea", "@fechaCrea", "@usuarioMod", "@fechaMod", "@idDetUsuarioCaja", "@nroDocumentoPago", "@tipoDocumentoPago", "@idMonedaPago", "@FechaDocumento", "@IdSubdiario", "@NroComprobante", "@Secuencia"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdCajaVenta, d.IdAgencia, d.IdAlmacen, d.IdTipoDocumento, d.Serie, d.NumeroDocumento, d.DebeHaber, d.Estado, d.Glosa, d.TipoCambio, d.IdMoneda, d.Pago, d.ImporteTarjeta, d.NumeroTarjeta, d.CodigoTarjeta, d.Adelanto, d.Ajuste, d.CodigoBanco, d.ImporteCheque, d.NumeroCheque, d.UsuarioCrea, d.FechaCrea, d.UsuarioMod, d.FechaMod, d.IdDetUsuarioCaja, d.NroDocumentoPago, d.TipoDocumentoPago, d.IdMonedaPago, d.FechaDocumento, d.IdSubdiario, d.NroComprobante, d.Secuencia}
        sql.EjecutarProcedure("Str_Tbl_Caja_Venta_U", parametros, valores, tipoParametro, 32)
    End Sub
    Public Sub Eliminar(d As NCajaVenta)
        Dim parametros() As Object = {"@idCajaVenta"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.IdCajaVenta}
        sql.EjecutarProcedure("Str_Tbl_Caja_Venta_D", parametros, valores, tipoParametro, 1)
    End Sub
    ''' <summary>
    ''' Elima un registro por su FK comprobante de pago
    ''' </summary>
    ''' <param name="d"></param>
    Public Sub EliminarFk(d As NCajaVenta)
        sql.EjecutarConsulta("de", "delete from Tbl_Caja_Venta where IdTipoDocumento='" & d.IdTipoDocumento & "' and Serie='" & d.Serie & "' and NumeroDocumento='" & d.NumeroDocumento & "'")
    End Sub

    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idCajaVenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Caja_Venta_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro_FK(d As NCajaVenta) As NCajaVenta
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@idAlmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAgencia, d.IdTipoDocumento, d.Serie, d.NumeroDocumento, d.IdAlmacen}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Caja_Venta_FK", parametros, valores, tipoParametro, 5).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdCajaVenta = IIf(dt.Rows(0).Item("idCajaVenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCajaVenta"))
            d.IdAgencia = IIf(dt.Rows(0).Item("idAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAgencia"))
            d.IdAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.IdTipoDocumento = IIf(dt.Rows(0).Item("idTipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoDocumento"))
            d.Serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.NumeroDocumento = IIf(dt.Rows(0).Item("numeroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento"))
            d.DebeHaber = IIf(dt.Rows(0).Item("debeHaber") Is DBNull.Value, Nothing, dt.Rows(0).Item("debeHaber"))
            d.Estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.Glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.TipoCambio = IIf(dt.Rows(0).Item("tipoCambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoCambio"))
            d.IdMoneda = IIf(dt.Rows(0).Item("idMoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMoneda"))
            d.Pago = IIf(dt.Rows(0).Item("pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("pago"))
            d.PagoUS = IIf(dt.Rows(0).Item("pagoUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagoUS"))
            d.PagoMN = IIf(dt.Rows(0).Item("pagoMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagoMN"))
            d.ImporteTarjeta = IIf(dt.Rows(0).Item("importeTarjeta") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeTarjeta"))
            d.ImporteTarjetaUS = IIf(dt.Rows(0).Item("importeTarjetaUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeTarjetaUS"))
            d.ImporteTarjetaMN = IIf(dt.Rows(0).Item("importeTarjetaMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeTarjetaMN"))
            d.NumeroTarjeta = IIf(dt.Rows(0).Item("numeroTarjeta") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroTarjeta"))
            d.CodigoTarjeta = IIf(dt.Rows(0).Item("codigoTarjeta") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigoTarjeta"))
            d.Adelanto = IIf(dt.Rows(0).Item("adelanto") Is DBNull.Value, Nothing, dt.Rows(0).Item("adelanto"))
            d.Ajuste = IIf(dt.Rows(0).Item("ajuste") Is DBNull.Value, Nothing, dt.Rows(0).Item("ajuste"))
            d.CodigoBanco = IIf(dt.Rows(0).Item("codigoBanco") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigoBanco"))
            d.ImporteCheque = IIf(dt.Rows(0).Item("importeCheque") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeCheque"))
            d.ImporteChequeUS = IIf(dt.Rows(0).Item("importeChequeUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeChequeUS"))
            d.ImporteChequeMN = IIf(dt.Rows(0).Item("importeChequeMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeChequeMN"))
            d.NumeroCheque = IIf(dt.Rows(0).Item("numeroCheque") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroCheque"))
            d.UsuarioCrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.FechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.UsuarioMod = IIf(dt.Rows(0).Item("usuarioMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioMod"))
            d.FechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.IdDetUsuarioCaja = IIf(dt.Rows(0).Item("idDetUsuarioCaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idDetUsuarioCaja"))
            d.NroDocumentoPago = IIf(dt.Rows(0).Item("nroDocumentoPago") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroDocumentoPago"))
            d.TipoDocumentoPago = IIf(dt.Rows(0).Item("tipoDocumentoPago") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoDocumentoPago"))
            d.IdMonedaPago = IIf(dt.Rows(0).Item("idMonedaPago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMonedaPago"))
            d.FechaDocumento = IIf(dt.Rows(0).Item("fechaDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaDocumento"))
            d.IdSubdiario = IIf(dt.Rows(0).Item("IdSubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("IdSubdiario"))
            d.NroComprobante = IIf(dt.Rows(0).Item("NroComprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("NroComprobante"))
            d.Secuencia = IIf(dt.Rows(0).Item("Secuencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("Secuencia"))
        Else
            d.IdAgencia = Nothing
            d.IdAlmacen = Nothing
            d.IdTipoDocumento = Nothing
            d.Serie = Nothing
            d.NumeroDocumento = Nothing
            d.DebeHaber = Nothing
            d.Estado = Nothing
            d.Glosa = Nothing
            d.TipoCambio = Nothing
            d.IdMoneda = Nothing
            d.Pago = Nothing
            d.PagoUS = Nothing
            d.PagoMN = Nothing
            d.ImporteTarjeta = Nothing
            d.ImporteTarjetaUS = Nothing
            d.ImporteTarjetaMN = Nothing
            d.NumeroTarjeta = Nothing
            d.CodigoTarjeta = Nothing
            d.Adelanto = Nothing
            d.Ajuste = Nothing
            d.CodigoBanco = Nothing
            d.ImporteCheque = Nothing
            d.ImporteChequeUS = Nothing
            d.ImporteChequeMN = Nothing
            d.NumeroCheque = Nothing
            d.UsuarioCrea = Nothing
            d.FechaCrea = Nothing
            d.UsuarioMod = Nothing
            d.FechaMod = Nothing
            d.IdDetUsuarioCaja = Nothing
            d.NroDocumentoPago = Nothing
            d.TipoDocumentoPago = Nothing
            d.IdMonedaPago = Nothing
            d.FechaDocumento = Nothing
            d.IdSubdiario = Nothing
            d.NroComprobante = Nothing
            d.Secuencia = Nothing
        End If
        Return d
    End Function

    Public Function Registro(d As NCajaVenta) As NCajaVenta
        Dim parametros() As Object = {"@idCajaVenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.IdCajaVenta}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Caja_Venta_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdCajaVenta = IIf(dt.Rows(0).Item("idCajaVenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCajaVenta"))
            d.IdAgencia = IIf(dt.Rows(0).Item("idAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAgencia"))
            d.IdAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.IdTipoDocumento = IIf(dt.Rows(0).Item("idTipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoDocumento"))
            d.Serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.NumeroDocumento = IIf(dt.Rows(0).Item("numeroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento"))
            d.DebeHaber = IIf(dt.Rows(0).Item("debeHaber") Is DBNull.Value, Nothing, dt.Rows(0).Item("debeHaber"))
            d.Estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.Glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.TipoCambio = IIf(dt.Rows(0).Item("tipoCambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoCambio"))
            d.IdMoneda = IIf(dt.Rows(0).Item("idMoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMoneda"))
            d.Pago = IIf(dt.Rows(0).Item("pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("pago"))
            d.PagoUS = IIf(dt.Rows(0).Item("pagoUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagoUS"))
            d.PagoMN = IIf(dt.Rows(0).Item("pagoMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagoMN"))
            d.ImporteTarjeta = IIf(dt.Rows(0).Item("importeTarjeta") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeTarjeta"))
            d.ImporteTarjetaUS = IIf(dt.Rows(0).Item("importeTarjetaUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeTarjetaUS"))
            d.ImporteTarjetaMN = IIf(dt.Rows(0).Item("importeTarjetaMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeTarjetaMN"))
            d.NumeroTarjeta = IIf(dt.Rows(0).Item("numeroTarjeta") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroTarjeta"))
            d.CodigoTarjeta = IIf(dt.Rows(0).Item("codigoTarjeta") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigoTarjeta"))
            d.Adelanto = IIf(dt.Rows(0).Item("adelanto") Is DBNull.Value, Nothing, dt.Rows(0).Item("adelanto"))
            d.Ajuste = IIf(dt.Rows(0).Item("ajuste") Is DBNull.Value, Nothing, dt.Rows(0).Item("ajuste"))
            d.CodigoBanco = IIf(dt.Rows(0).Item("codigoBanco") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigoBanco"))
            d.ImporteCheque = IIf(dt.Rows(0).Item("importeCheque") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeCheque"))
            d.ImporteChequeUS = IIf(dt.Rows(0).Item("importeChequeUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeChequeUS"))
            d.ImporteChequeMN = IIf(dt.Rows(0).Item("importeChequeMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeChequeMN"))
            d.NumeroCheque = IIf(dt.Rows(0).Item("numeroCheque") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroCheque"))
            d.UsuarioCrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.FechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.UsuarioMod = IIf(dt.Rows(0).Item("usuarioMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioMod"))
            d.FechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.IdDetUsuarioCaja = IIf(dt.Rows(0).Item("idDetUsuarioCaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idDetUsuarioCaja"))
            d.NroDocumentoPago = IIf(dt.Rows(0).Item("nroDocumentoPago") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroDocumentoPago"))
            d.TipoDocumentoPago = IIf(dt.Rows(0).Item("tipoDocumentoPago") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoDocumentoPago"))
            d.IdMonedaPago = IIf(dt.Rows(0).Item("idMonedaPago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMonedaPago"))
            d.FechaDocumento = IIf(dt.Rows(0).Item("fechaDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaDocumento"))
            d.IdSubdiario = IIf(dt.Rows(0).Item("IdSubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("IdSubdiario"))
            d.NroComprobante = IIf(dt.Rows(0).Item("NroComprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("NroComprobante"))
            d.Secuencia = IIf(dt.Rows(0).Item("Secuencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("Secuencia"))
        Else
            d.IdAgencia = Nothing
            d.IdAlmacen = Nothing
            d.IdTipoDocumento = Nothing
            d.Serie = Nothing
            d.NumeroDocumento = Nothing
            d.DebeHaber = Nothing
            d.Estado = Nothing
            d.Glosa = Nothing
            d.TipoCambio = Nothing
            d.IdMoneda = Nothing
            d.Pago = Nothing
            d.PagoUS = Nothing
            d.PagoMN = Nothing
            d.ImporteTarjeta = Nothing
            d.ImporteTarjetaUS = Nothing
            d.ImporteTarjetaMN = Nothing
            d.NumeroTarjeta = Nothing
            d.CodigoTarjeta = Nothing
            d.Adelanto = Nothing
            d.Ajuste = Nothing
            d.CodigoBanco = Nothing
            d.ImporteCheque = Nothing
            d.ImporteChequeUS = Nothing
            d.ImporteChequeMN = Nothing
            d.NumeroCheque = Nothing
            d.UsuarioCrea = Nothing
            d.FechaCrea = Nothing
            d.UsuarioMod = Nothing
            d.FechaMod = Nothing
            d.IdDetUsuarioCaja = Nothing
            d.NroDocumentoPago = Nothing
            d.TipoDocumentoPago = Nothing
            d.IdMonedaPago = Nothing
            d.FechaDocumento = Nothing
            d.IdSubdiario = Nothing
            d.NroComprobante = Nothing
            d.Secuencia = Nothing
        End If
        Return d
    End Function

    Public Function AsientoContado(fecha As DateTime, Idalmacen As String, idcaja As String, mostrar As String) As DataTable
        Dim dt As DataTable
        Dim campos() As Object = {"@FechaDia", "@IdCaja", "@IdAlmacen", "@Mostrar"}
        Dim tipoparametro() As Object = {SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {fecha, idcaja, Idalmacen, mostrar}
        dt = sql.ProcedureSQL("AsientoContados", campos, valores, tipoparametro, 4).Tables(0)
        Return dt
    End Function
    Public Function AsientoContadoCaja(fecha As DateTime, idcaja As String, mostrar As String) As DataTable
        Dim dt As DataTable
        Dim campos() As Object = {"@FechaDia", "@IdCaja", "@Mostrar"}
        Dim tipoparametro() As Object = {SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {fecha, idcaja, mostrar}
        dt = sql.ProcedureSQL("AsientoContadosCaja", campos, valores, tipoparametro, 3).Tables(0)
        Return dt
    End Function

#End Region


End Class
