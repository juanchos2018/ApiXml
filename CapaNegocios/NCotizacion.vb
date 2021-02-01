Imports CapaDatos

Public Class NCotizacion
    Dim px As New ClsConexion

#Region "Declarations"

    Private _idAgencia As String
    Private _idTipoDocumento As String
    Private _serie As String
    Private _numeroDocumento As String
    Private _numeroPedido As String
    Private _fechaDocumento As System.DateTime
    Private _fechaVencimineto As System.DateTime
    Private _debeHaber As String
    Private _idVendedor As String
    Private _idCaja As String
    Private _idCliente As String
    Private _nombreCliente As String
    Private _direccion As String
    Private _rUC As String
    Private _idAlmacen As String
    Private _idFormaVenta As String
    Private _idMoneda As String
    Private _tipoCambio As Decimal
    Private _importeTotal As Decimal
    Private _importeIGV As Decimal
    Private _saldo As Decimal
    Private _importeDescuento As Decimal
    Private _numeroOrden As String
    Private _idTipoDocumento1 As String
    Private _serie1 As String
    Private _numeroDocumento1 As String
    Private _descripcion As String
    Private _estado As String
    Private _facturaGuia As String
    Private _idTransportista As String
    Private _idCentroCosto As String
    Private _idMaquina As String
    Private _destino As String
    Private _idTipoFactura As String
    Private _idTipoAnexo As String
    Private _idAnexo As String
    Private _descuneto1 As Decimal
    Private _descuento2 As Decimal
    Private _flete As Decimal
    Private _embalaje As Decimal
    Private _tasa As Decimal
    Private _idUsuarioOperador As String
    Private _idUsuarioSectorista As String
    Private _idCadena As String
    Private _idInternoCadena As String
    Private _idAutorizacion As String
    Private _reparto As String
    Private _usuarioCrea As String
    Private _fechaCrea As System.DateTime
    Private _usuarioMod As String
    Private _fechaMod As System.DateTime
    Private _idTipoNotaCredito As String
    Private _linea As String
    Private _impreso As String
    Private _anuladoNC As String
    Private _idVendedor1 As String
    Private _iGV As Decimal
    Private _idchofer As String
    Private _idZonaVenta As String

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

    Public Property NumeroPedido As String
        Get
            Return _numeroPedido
        End Get
        Set
            _numeroPedido = Value
        End Set
    End Property

    Public Property FechaDocumento As System.DateTime
        Get
            Return _fechaDocumento
        End Get
        Set
            _fechaDocumento = Value
        End Set
    End Property

    Public Property FechaVencimineto As System.DateTime
        Get
            Return _fechaVencimineto
        End Get
        Set
            _fechaVencimineto = Value
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

    Public Property IdVendedor As String
        Get
            Return _idVendedor
        End Get
        Set
            _idVendedor = Value
        End Set
    End Property

    Public Property IdCaja As String
        Get
            Return _idCaja
        End Get
        Set
            _idCaja = Value
        End Set
    End Property

    Public Property IdCliente As String
        Get
            Return _idCliente
        End Get
        Set
            _idCliente = Value
        End Set
    End Property

    Public Property NombreCliente As String
        Get
            Return _nombreCliente
        End Get
        Set
            _nombreCliente = Value
        End Set
    End Property

    Public Property Direccion As String
        Get
            Return _direccion
        End Get
        Set
            _direccion = Value
        End Set
    End Property

    Public Property RUC As String
        Get
            Return _rUC
        End Get
        Set
            _rUC = Value
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

    Public Property IdFormaVenta As String
        Get
            Return _idFormaVenta
        End Get
        Set
            _idFormaVenta = Value
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

    Public Property TipoCambio As Decimal
        Get
            Return _tipoCambio
        End Get
        Set
            _tipoCambio = Value
        End Set
    End Property

    Public Property ImporteTotal As Decimal
        Get
            Return _importeTotal
        End Get
        Set
            _importeTotal = Value
        End Set
    End Property

    Public Property ImporteIGV As Decimal
        Get
            Return _importeIGV
        End Get
        Set
            _importeIGV = Value
        End Set
    End Property

    Public Property Saldo As Decimal
        Get
            Return _saldo
        End Get
        Set
            _saldo = Value
        End Set
    End Property

    Public Property ImporteDescuento As Decimal
        Get
            Return _importeDescuento
        End Get
        Set
            _importeDescuento = Value
        End Set
    End Property

    Public Property NumeroOrden As String
        Get
            Return _numeroOrden
        End Get
        Set
            _numeroOrden = Value
        End Set
    End Property

    Public Property IdTipoDocumento1 As String
        Get
            Return _idTipoDocumento1
        End Get
        Set
            _idTipoDocumento1 = Value
        End Set
    End Property

    Public Property Serie1 As String
        Get
            Return _serie1
        End Get
        Set
            _serie1 = Value
        End Set
    End Property

    Public Property NumeroDocumento1 As String
        Get
            Return _numeroDocumento1
        End Get
        Set
            _numeroDocumento1 = Value
        End Set
    End Property

    Public Property Descripcion As String
        Get
            Return _descripcion
        End Get
        Set
            _descripcion = Value
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

    Public Property FacturaGuia As String
        Get
            Return _facturaGuia
        End Get
        Set
            _facturaGuia = Value
        End Set
    End Property

    Public Property IdTransportista As String
        Get
            Return _idTransportista
        End Get
        Set
            _idTransportista = Value
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

    Public Property IdMaquina As String
        Get
            Return _idMaquina
        End Get
        Set
            _idMaquina = Value
        End Set
    End Property

    Public Property Destino As String
        Get
            Return _destino
        End Get
        Set
            _destino = Value
        End Set
    End Property

    Public Property IdTipoFactura As String
        Get
            Return _idTipoFactura
        End Get
        Set
            _idTipoFactura = Value
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

    Public Property IdAnexo As String
        Get
            Return _idAnexo
        End Get
        Set
            _idAnexo = Value
        End Set
    End Property

    Public Property Descuneto1 As Decimal
        Get
            Return _descuneto1
        End Get
        Set
            _descuneto1 = Value
        End Set
    End Property

    Public Property Descuento2 As Decimal
        Get
            Return _descuento2
        End Get
        Set
            _descuento2 = Value
        End Set
    End Property

    Public Property Flete As Decimal
        Get
            Return _flete
        End Get
        Set
            _flete = Value
        End Set
    End Property

    Public Property Embalaje As Decimal
        Get
            Return _embalaje
        End Get
        Set
            _embalaje = Value
        End Set
    End Property

    Public Property Tasa As Decimal
        Get
            Return _tasa
        End Get
        Set
            _tasa = Value
        End Set
    End Property

    Public Property IdUsuarioOperador As String
        Get
            Return _idUsuarioOperador
        End Get
        Set
            _idUsuarioOperador = Value
        End Set
    End Property

    Public Property IdUsuarioSectorista As String
        Get
            Return _idUsuarioSectorista
        End Get
        Set
            _idUsuarioSectorista = Value
        End Set
    End Property

    Public Property IdCadena As String
        Get
            Return _idCadena
        End Get
        Set
            _idCadena = Value
        End Set
    End Property

    Public Property IdInternoCadena As String
        Get
            Return _idInternoCadena
        End Get
        Set
            _idInternoCadena = Value
        End Set
    End Property

    Public Property IdAutorizacion As String
        Get
            Return _idAutorizacion
        End Get
        Set
            _idAutorizacion = Value
        End Set
    End Property

    Public Property Reparto As String
        Get
            Return _reparto
        End Get
        Set
            _reparto = Value
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

    Public Property IdTipoNotaCredito As String
        Get
            Return _idTipoNotaCredito
        End Get
        Set
            _idTipoNotaCredito = Value
        End Set
    End Property

    Public Property Linea As String
        Get
            Return _linea
        End Get
        Set
            _linea = Value
        End Set
    End Property

    Public Property Impreso As String
        Get
            Return _impreso
        End Get
        Set
            _impreso = Value
        End Set
    End Property

    Public Property AnuladoNC As String
        Get
            Return _anuladoNC
        End Get
        Set
            _anuladoNC = Value
        End Set
    End Property

    Public Property IdVendedor1 As String
        Get
            Return _idVendedor1
        End Get
        Set
            _idVendedor1 = Value
        End Set
    End Property

    Public Property IGV As Decimal
        Get
            Return _iGV
        End Get
        Set
            _iGV = Value
        End Set
    End Property

    Public Property Idchofer As String
        Get
            Return _idchofer
        End Get
        Set
            _idchofer = Value
        End Set
    End Property

    Public Property IdZonaVenta As String
        Get
            Return _idZonaVenta
        End Get
        Set
            _idZonaVenta = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idAgencia As String, ByVal idTipoDocumento As String, ByVal serie As String, ByVal numeroDocumento As String, ByVal numeroPedido As String, ByVal fechaDocumento As System.DateTime, ByVal fechaVencimineto As System.DateTime, ByVal debeHaber As String, ByVal idVendedor As String, ByVal idCaja As String, ByVal idCliente As String, ByVal nombreCliente As String, ByVal direccion As String, ByVal rUC As String, ByVal idAlmacen As String, ByVal idFormaVenta As String, ByVal idMoneda As String, ByVal tipoCambio As Decimal, ByVal importeTotal As Decimal, ByVal importeIGV As Decimal, ByVal saldo As Decimal, ByVal importeDescuento As Decimal, ByVal numeroOrden As String, ByVal idTipoDocumento1 As String, ByVal serie1 As String, ByVal numeroDocumento1 As String, ByVal descripcion As String, ByVal estado As String, ByVal facturaGuia As String, ByVal idTransportista As String, ByVal idCentroCosto As String, ByVal idMaquina As String, ByVal destino As String, ByVal idTipoFactura As String, ByVal idTipoAnexo As String, ByVal idAnexo As String, ByVal descuneto1 As Decimal, ByVal descuento2 As Decimal, ByVal flete As Decimal, ByVal embalaje As Decimal, ByVal tasa As Decimal, ByVal idUsuarioOperador As String, ByVal idUsuarioSectorista As String, ByVal idCadena As String, ByVal idInternoCadena As String, ByVal idAutorizacion As String, ByVal reparto As String, ByVal usuarioCrea As String, ByVal fechaCrea As System.DateTime, ByVal usuarioMod As String, ByVal fechaMod As System.DateTime, ByVal idTipoNotaCredito As String, ByVal linea As String, ByVal impreso As String, ByVal anuladoNC As String, ByVal idVendedor1 As String, ByVal iGV As Decimal, ByVal idchofer As String, ByVal idZonaVenta As String)
        Me.New()
    End Sub
#End Region

#Region "Metodos"


    Public Function ObtenerCabecera(p As NCotizacion) As NCotizacion
        Dim dt_cab As New DataTable
        Dim cadena As String = " SELECT     IdAgencia, IdTipoDocumento, Serie, NumeroDocumento, NumeroPedido, FechaDocumento, FechaVencimineto, DebeHaber, IdVendedor, IdCaja, IdCliente,  "
        cadena += " NombreCliente, Direccion, RUC, IdAlmacen, IdFormaVenta, IdMoneda, TipoCambio, ImporteTotal, ImporteIGV, Saldo, ImporteDescuento, ISNULL(NumeroOrden,'')AS NumeroOrden,  "
        cadena += " IdTipoDocumento1, Serie1, NumeroDocumento1, Descripcion, Estado, FacturaGuia, IdTransportista, IdCentroCosto, IdMaquina, Destino, IdTipoFactura, IdTipoAnexo, "
        cadena += " IdAnexo, Descuneto1, Descuento2, Flete, Embalaje, Tasa, IdUsuarioOperador, IdUsuarioSectorista, IdCadena, IdInternoCadena, IdAutorizacion, Reparto, UsuarioCrea, "
        cadena += " FechaCrea, UsuarioMod, FechaMod, IdTipoNotaCredito, Linea, Impreso, AnuladoNC, IdVendedor1, IGV, Idchofer, IdZonaVenta  "
        cadena += "  FROM Pedido "
        cadena += " where IdAgencia='" & p.IdAgencia & "' AND IdAlmacen='" & p.IdAlmacen & "' and Serie='" & p.Serie & "' and IdTipoDocumento='" & p.IdTipoDocumento & "' and NumeroDocumento='" & p.NumeroDocumento & "' "
        dt_cab = px.EjecutarConsulta("cab", cadena).Tables(0)
        If dt_cab.Rows.Count > 0 Then
            With dt_cab
                p.IdAgencia = .Rows(0).Item("IdAgencia")
                p.IdVendedor = .Rows(0).Item("IdVendedor")
                p.IdCliente = .Rows(0).Item("IdCliente")
                p.NombreCliente = .Rows(0).Item("NombreCliente")
                p.Direccion = .Rows(0).Item("Direccion")
                p.RUC = .Rows(0).Item("RUC")
                p.IdFormaVenta = .Rows(0).Item("IdFormaVenta")
                p.IdMoneda = .Rows(0).Item("IdMoneda")
                p.TipoCambio = .Rows(0).Item("TipoCambio")
                p.ImporteTotal = .Rows(0).Item("ImporteTotal")
                p.ImporteIGV = .Rows(0).Item("ImporteIGV")
                '                TxtSubTotal.Text = .Rows(0).Item("ImporteTotal") - .Rows(0).Item("ImporteIGV")
                p.Descripcion = .Rows(0).Item("Descripcion")
                p.IGV = .Rows(0).Item("IGV")
                p.Idchofer = .Rows(0).Item("IdChofer")
                p.NumeroOrden = .Rows(0).Item("NumeroOrden")
                p.FechaDocumento = .Rows(0).Item("FechaDocumento")
            End With
        End If
        Return p
    End Function

    Public Function Lista(p As NCotizacion) As DataTable
        Dim cadena As String = "select IdTipoDocumento,Serie,NumeroDocumento,FechaDocumento,IdCliente,Nombrecliente,importetotal from pedido"
        cadena += " where idAlmacen='" & p.IdAlmacen & "'"
        'cadena += " where IdTipoDocumento='" & p.IdTipoDocumento & "' AND SERIE='" & p.Serie & "'and numeroDocumento='" & p.NumeroDocumento & "' and idAlmacen='" & p.IdAlmacen & "'"
        Dim dt As New DataTable
        dt = px.EjecutarConsulta("Pedido", cadena).Tables(0)
        Return dt
    End Function

#End Region

End Class
