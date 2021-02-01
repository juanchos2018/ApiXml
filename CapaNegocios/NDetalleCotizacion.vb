Imports CapaDatos
Public Class NDetalleCotizacion
    Dim px As New ClsConexion
#Region "Declarations"

    Private _idAgencia As String
    Private _idTipoDocumento As String
    Private _serie As String
    Private _numeroDocumento As String
    Private _item As String
    Private _idArticulo As String
    Private _descripcion As String
    Private _texto As String
    Private _cantidad As Decimal
    Private _unidad As String
    Private _serie1 As String
    Private _cantidad1 As Decimal
    Private _unidadEnvase As Decimal
    Private _numeroEnvase As Decimal
    Private _saldoEntrega As Decimal
    Private _precioVenta As Decimal
    Private _precioVentaH As Decimal
    Private _precioVentaImportacion As Decimal
    Private _precioVentaImportacionH As Decimal
    Private _precioSIGV As Decimal
    Private _importeDescuento As Decimal
    Private _descuentoDocumento As Decimal
    Private _cargoDistribucion As Decimal
    Private _iGV As Decimal
    Private _importeIGV As Decimal
    Private _importeUS As Decimal
    Private _importeMN As Decimal
    Private _idTipoITemDescuento As String
    Private _descuento1 As Decimal
    Private _importeDescuento1 As Decimal
    Private _descuento2 As Decimal
    Private _importeDescuento2 As Decimal
    Private _descuento3 As Decimal
    Private _importeDescuento3 As Decimal
    Private _descuento4 As Decimal
    Private _importeDescuento4 As Decimal
    Private _descuento5 As Decimal
    Private _importeDescuento5 As Decimal
    Private _descuento6 As Decimal
    Private _estado As String
    Private _vendedor As String
    Private _idAlmacen As String
    Private _numeroCaja As String
    Private _stock As String
    Private _fechaSDocumento As System.DateTime
    Private _idLinea As String
    Private _idCampania As String
    Private _numeroPaquete As String
    Private _nroDescuentoFinaciero As String
    Private _nroDescuentoLaboratorio As String
    Private _nroDescuentoAdicional As String
    Private _nroDescuentoBonificacion As String
    Private _nroDescuentoFlag As String
    Private _comision As Decimal
    Private _importeComision As Decimal
    Private _usuarioCrea As String
    Private _fechaCrea As System.DateTime
    Private _precioUnitarioOrigen As Decimal
    Private _idVendedor2 As String
    Private _identrada As String
    Private _nPFacturado As String
    Private _idLista As Integer

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

    Public Property Item As String
        Get
            Return _item
        End Get
        Set
            _item = Value
        End Set
    End Property

    Public Property IdArticulo As String
        Get
            Return _idArticulo
        End Get
        Set
            _idArticulo = Value
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

    Public Property Texto As String
        Get
            Return _texto
        End Get
        Set
            _texto = Value
        End Set
    End Property

    Public Property Cantidad As Decimal
        Get
            Return _cantidad
        End Get
        Set
            _cantidad = Value
        End Set
    End Property

    Public Property Unidad As String
        Get
            Return _unidad
        End Get
        Set
            _unidad = Value
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

    Public Property Cantidad1 As Decimal
        Get
            Return _cantidad1
        End Get
        Set
            _cantidad1 = Value
        End Set
    End Property

    Public Property UnidadEnvase As Decimal
        Get
            Return _unidadEnvase
        End Get
        Set
            _unidadEnvase = Value
        End Set
    End Property

    Public Property NumeroEnvase As Decimal
        Get
            Return _numeroEnvase
        End Get
        Set
            _numeroEnvase = Value
        End Set
    End Property

    Public Property SaldoEntrega As Decimal
        Get
            Return _saldoEntrega
        End Get
        Set
            _saldoEntrega = Value
        End Set
    End Property

    Public Property PrecioVenta As Decimal
        Get
            Return _precioVenta
        End Get
        Set
            _precioVenta = Value
        End Set
    End Property

    Public Property PrecioVentaH As Decimal
        Get
            Return _precioVentaH
        End Get
        Set
            _precioVentaH = Value
        End Set
    End Property

    Public Property PrecioVentaImportacion As Decimal
        Get
            Return _precioVentaImportacion
        End Get
        Set
            _precioVentaImportacion = Value
        End Set
    End Property

    Public Property PrecioVentaImportacionH As Decimal
        Get
            Return _precioVentaImportacionH
        End Get
        Set
            _precioVentaImportacionH = Value
        End Set
    End Property

    Public Property PrecioSIGV As Decimal
        Get
            Return _precioSIGV
        End Get
        Set
            _precioSIGV = Value
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

    Public Property DescuentoDocumento As Decimal
        Get
            Return _descuentoDocumento
        End Get
        Set
            _descuentoDocumento = Value
        End Set
    End Property

    Public Property CargoDistribucion As Decimal
        Get
            Return _cargoDistribucion
        End Get
        Set
            _cargoDistribucion = Value
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

    Public Property ImporteIGV As Decimal
        Get
            Return _importeIGV
        End Get
        Set
            _importeIGV = Value
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

    Public Property IdTipoITemDescuento As String
        Get
            Return _idTipoITemDescuento
        End Get
        Set
            _idTipoITemDescuento = Value
        End Set
    End Property

    Public Property Descuento1 As Decimal
        Get
            Return _descuento1
        End Get
        Set
            _descuento1 = Value
        End Set
    End Property

    Public Property ImporteDescuento1 As Decimal
        Get
            Return _importeDescuento1
        End Get
        Set
            _importeDescuento1 = Value
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

    Public Property ImporteDescuento2 As Decimal
        Get
            Return _importeDescuento2
        End Get
        Set
            _importeDescuento2 = Value
        End Set
    End Property

    Public Property Descuento3 As Decimal
        Get
            Return _descuento3
        End Get
        Set
            _descuento3 = Value
        End Set
    End Property

    Public Property ImporteDescuento3 As Decimal
        Get
            Return _importeDescuento3
        End Get
        Set
            _importeDescuento3 = Value
        End Set
    End Property

    Public Property Descuento4 As Decimal
        Get
            Return _descuento4
        End Get
        Set
            _descuento4 = Value
        End Set
    End Property

    Public Property ImporteDescuento4 As Decimal
        Get
            Return _importeDescuento4
        End Get
        Set
            _importeDescuento4 = Value
        End Set
    End Property

    Public Property Descuento5 As Decimal
        Get
            Return _descuento5
        End Get
        Set
            _descuento5 = Value
        End Set
    End Property

    Public Property ImporteDescuento5 As Decimal
        Get
            Return _importeDescuento5
        End Get
        Set
            _importeDescuento5 = Value
        End Set
    End Property

    Public Property Descuento6 As Decimal
        Get
            Return _descuento6
        End Get
        Set
            _descuento6 = Value
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

    Public Property Vendedor As String
        Get
            Return _vendedor
        End Get
        Set
            _vendedor = Value
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

    Public Property NumeroCaja As String
        Get
            Return _numeroCaja
        End Get
        Set
            _numeroCaja = Value
        End Set
    End Property

    Public Property Stock As String
        Get
            Return _stock
        End Get
        Set
            _stock = Value
        End Set
    End Property

    Public Property FechaSDocumento As System.DateTime
        Get
            Return _fechaSDocumento
        End Get
        Set
            _fechaSDocumento = Value
        End Set
    End Property

    Public Property IdLinea As String
        Get
            Return _idLinea
        End Get
        Set
            _idLinea = Value
        End Set
    End Property

    Public Property IdCampania As String
        Get
            Return _idCampania
        End Get
        Set
            _idCampania = Value
        End Set
    End Property

    Public Property NumeroPaquete As String
        Get
            Return _numeroPaquete
        End Get
        Set
            _numeroPaquete = Value
        End Set
    End Property

    Public Property NroDescuentoFinaciero As String
        Get
            Return _nroDescuentoFinaciero
        End Get
        Set
            _nroDescuentoFinaciero = Value
        End Set
    End Property

    Public Property NroDescuentoLaboratorio As String
        Get
            Return _nroDescuentoLaboratorio
        End Get
        Set
            _nroDescuentoLaboratorio = Value
        End Set
    End Property

    Public Property NroDescuentoAdicional As String
        Get
            Return _nroDescuentoAdicional
        End Get
        Set
            _nroDescuentoAdicional = Value
        End Set
    End Property

    Public Property NroDescuentoBonificacion As String
        Get
            Return _nroDescuentoBonificacion
        End Get
        Set
            _nroDescuentoBonificacion = Value
        End Set
    End Property

    Public Property NroDescuentoFlag As String
        Get
            Return _nroDescuentoFlag
        End Get
        Set
            _nroDescuentoFlag = Value
        End Set
    End Property

    Public Property Comision As Decimal
        Get
            Return _comision
        End Get
        Set
            _comision = Value
        End Set
    End Property

    Public Property ImporteComision As Decimal
        Get
            Return _importeComision
        End Get
        Set
            _importeComision = Value
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

    Public Property PrecioUnitarioOrigen As Decimal
        Get
            Return _precioUnitarioOrigen
        End Get
        Set
            _precioUnitarioOrigen = Value
        End Set
    End Property

    Public Property IdVendedor2 As String
        Get
            Return _idVendedor2
        End Get
        Set
            _idVendedor2 = Value
        End Set
    End Property

    Public Property identrada As String
        Get
            Return _identrada
        End Get
        Set
            _identrada = Value
        End Set
    End Property

    Public Property NPFacturado As String
        Get
            Return _nPFacturado
        End Get
        Set
            _nPFacturado = Value
        End Set
    End Property

    Public Property IdLista As Integer
        Get
            Return _idLista
        End Get
        Set
            _idLista = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idAgencia As String, ByVal idTipoDocumento As String, ByVal serie As String, ByVal numeroDocumento As String, ByVal item As String, ByVal idArticulo As String, ByVal descripcion As String, ByVal texto As String, ByVal cantidad As Decimal, ByVal unidad As String, ByVal serie1 As String, ByVal cantidad1 As Decimal, ByVal unidadEnvase As Decimal, ByVal numeroEnvase As Decimal, ByVal saldoEntrega As Decimal, ByVal precioVenta As Decimal, ByVal precioVentaH As Decimal, ByVal precioVentaImportacion As Decimal, ByVal precioVentaImportacionH As Decimal, ByVal precioSIGV As Decimal, ByVal importeDescuento As Decimal, ByVal descuentoDocumento As Decimal, ByVal cargoDistribucion As Decimal, ByVal iGV As Decimal, ByVal importeIGV As Decimal, ByVal importeUS As Decimal, ByVal importeMN As Decimal, ByVal idTipoITemDescuento As String, ByVal descuento1 As Decimal, ByVal importeDescuento1 As Decimal, ByVal descuento2 As Decimal, ByVal importeDescuento2 As Decimal, ByVal descuento3 As Decimal, ByVal importeDescuento3 As Decimal, ByVal descuento4 As Decimal, ByVal importeDescuento4 As Decimal, ByVal descuento5 As Decimal, ByVal importeDescuento5 As Decimal, ByVal descuento6 As Decimal, ByVal estado As String, ByVal vendedor As String, ByVal idAlmacen As String, ByVal numeroCaja As String, ByVal stock As String, ByVal fechaSDocumento As System.DateTime, ByVal idLinea As String, ByVal idCampania As String, ByVal numeroPaquete As String, ByVal nroDescuentoFinaciero As String, ByVal nroDescuentoLaboratorio As String, ByVal nroDescuentoAdicional As String, ByVal nroDescuentoBonificacion As String, ByVal nroDescuentoFlag As String, ByVal comision As Decimal, ByVal importeComision As Decimal, ByVal usuarioCrea As String, ByVal fechaCrea As System.DateTime, ByVal precioUnitarioOrigen As Decimal, ByVal idVendedor2 As String, ByVal identrada As String, ByVal nPFacturado As String, ByVal idLista As Integer)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    Public Function obtenerDetalle(d As NDetalleCotizacion) As DataTable
        Dim cadena As String = " SELECT IdAgencia, IdTipoDocumento, Serie, NumeroDocumento, Item, IdArticulo, Descripcion, Texto, Cantidad, Unidad, Serie1, Cantidad1, UnidadEnvase, NumeroEnvase,  "
        cadena += " SaldoEntrega, PrecioVenta, PrecioVentaH, PrecioVentaImportacion, PrecioVentaImportacionH, PrecioSIGV, ImporteDescuento, DescuentoDocumento,  "
        cadena += " CargoDistribucion, IGV, ImporteIGV, ImporteUS, ImporteMN, IdTipoITemDescuento, Descuento1, ImporteDescuento1, Descuento2, ImporteDescuento2, Descuento3, "
        cadena += " ImporteDescuento3, Descuento4, ImporteDescuento4, Descuento5, ImporteDescuento5, Descuento6, Estado, Vendedor, IdAlmacen, NumeroCaja, Stock, "
        cadena += " FechaSDocumento, IdLinea, IdCampania, NumeroPaquete, NroDescuentoFinaciero, NroDescuentoLaboratorio, NroDescuentoAdicional, NroDescuentoBonificacion, "
        cadena += " NroDescuentoFlag, Comision, ImporteComision, UsuarioCrea, FechaCrea, PrecioUnitarioOrigen, IdVendedor2, identrada, NPFacturado, IdLista "
        cadena += " FROM         DetallePedido "
        cadena += " WHERE     (IdAgencia = '" & d.IdAgencia & "') AND (IdTipoDocumento = '" & d.IdTipoDocumento & "') AND (Serie = '" & d.Serie & "') AND (NumeroDocumento = '" & d.NumeroDocumento & "')"
        Dim dt_det As New DataTable
        dt_det = px.EjecutarConsulta("det", cadena).Tables(0)

        Return dt_det
    End Function
#End Region
End Class
