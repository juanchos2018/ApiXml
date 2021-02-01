Imports CapaDatos
Public Class NDetallePreVenta
    Dim sql As New ClsConexion
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

    'Public Sub Agregar(d As NDetallePreVenta)
    '    Dim paramsD() As Object = {"@IdAgencia", "@IdTipoDocumento", "@Serie", "@NumeroDocumento", "@Item",
    '        "@IdArticulo", "@Descripcion", "@Cantidad", "@Unidad", "@PrecioUnitario", "@PrecioSIGV",
    '        "@ImporteDescuento", "@IGV", "@ImporteIGV", "@ImporteUS", "@ImporteMN", "@Descuento1",
    '        "@ImporteDescuento1", "@Descuento2", "@ImporteDescuento2", "@Descuento3",
    '        "@ImporteDescuento3", "@IdAlmacen", "@Usuario", "@Fecha", "@idEntrada"}
    '    Dim tipoParametroD() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Float, SqlDbType.Char, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char}
    '    Dim valsD() As Object = {d.IdAgencia, d.IdTipoDocumento, d.Serie, d.NumeroDocumento, d.Item,
    '        d.IdArticulo, d.Descripcion, d.Cantidad, d.Unidad, d.PrecioVenta, d.PrecioSIGV,
    '        d.ImporteDescuento, d.IGV, d.ImporteIGV, d.ImporteUS, d.ImporteMN, d.Descuento1,
    '        d.ImporteDescuento1, d.Descuento2, d.ImporteDescuento2, d.Descuento3,
    '        d.ImporteDescuento3, d.IdAlmacen, d.UsuarioCrea, d.FechaCrea, d.identrada} ' FillCeros(k,4)
    '    sql.EjecutarProcedure("proc_AddDetallePreVenta", paramsD, valsD, tipoParametroD, 26)

    'End Sub

    Public Function obtenerDetalle(d As NDetallePreVenta) As DataTable
        Dim cadena As String = " SELECT IdAgencia, IdTipoDocumento, Serie, NumeroDocumento, Item, IdArticulo, Descripcion, Texto, Cantidad, Unidad, Serie1, Cantidad1, UnidadEnvase, NumeroEnvase,  "
        cadena += " SaldoEntrega, PrecioVenta, PrecioVentaH, PrecioVentaImportacion, PrecioVentaImportacionH, PrecioSIGV, ImporteDescuento, DescuentoDocumento,  "
        cadena += " CargoDistribucion, IGV, ImporteIGV, ImporteUS, ImporteMN, IdTipoITemDescuento, Descuento1, ImporteDescuento1, Descuento2, ImporteDescuento2, Descuento3, "
        cadena += " ImporteDescuento3, Descuento4, ImporteDescuento4, Descuento5, ImporteDescuento5, Descuento6, Estado, Vendedor, IdAlmacen, NumeroCaja, Stock, "
        cadena += " FechaSDocumento, IdLinea, IdCampania, NumeroPaquete, NroDescuentoFinaciero, NroDescuentoLaboratorio, NroDescuentoAdicional, NroDescuentoBonificacion, "
        cadena += " NroDescuentoFlag, Comision, ImporteComision, UsuarioCrea, FechaCrea, PrecioUnitarioOrigen, IdVendedor2, identrada, NPFacturado, IdLista "
        cadena += " FROM         detallepreventa "
        cadena += " WHERE     (IdAgencia = '" & d.IdAgencia & "') AND (IdTipoDocumento = '" & d.IdTipoDocumento & "') AND (Serie = '" & d.Serie & "') AND (NumeroDocumento = '" & d.NumeroDocumento & "') "
        cadena += " and saldoEntrega<>0 "
        Dim dt_det As New DataTable
        dt_det = sql.EjecutarConsulta("det", cadena).Tables(0)
        Return dt_det
    End Function
    Public Function Detalle(d As NDetallePreVenta) As DataTable
        Dim cadena As String = " SELECT Item, IdArticulo, Descripcion,Unidad, Cantidad,"
        cadena += " PrecioVenta as PUnit,ImporteUS, ImporteMN "
        cadena += " FROM  detallepreventa "
        cadena += " WHERE     (IdAgencia = '" & d.IdAgencia & "') AND (IdTipoDocumento = '" & d.IdTipoDocumento & "') AND (Serie = '" & d.Serie & "') AND (NumeroDocumento = '" & d.NumeroDocumento & "')"
        Dim dt_det As New DataTable
        dt_det = sql.EjecutarConsulta("det", cadena).Tables(0)
        Return dt_det
    End Function
    Public Sub Agregar(d As NDetallePreVenta)
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@item", "@idArticulo", "@descripcion", "@texto", "@cantidad", "@unidad", "@serie1", "@cantidad1", "@unidadEnvase", "@numeroEnvase", "@saldoEntrega", "@precioVenta", "@precioVentaH", "@precioVentaImportacion", "@precioVentaImportacionH", "@precioSIGV", "@importeDescuento", "@descuentoDocumento", "@cargoDistribucion", "@iGV", "@importeIGV", "@importeUS", "@importeMN", "@idTipoITemDescuento", "@descuento1", "@importeDescuento1", "@descuento2", "@importeDescuento2", "@descuento3", "@importeDescuento3", "@descuento4", "@importeDescuento4", "@descuento5", "@importeDescuento5", "@descuento6", "@estado", "@vendedor", "@idAlmacen", "@numeroCaja", "@stock", "@fechaSDocumento", "@idLinea", "@idCampania", "@numeroPaquete", "@nroDescuentoFinaciero", "@nroDescuentoLaboratorio", "@nroDescuentoAdicional", "@nroDescuentoBonificacion", "@nroDescuentoFlag", "@comision", "@importeComision", "@usuarioCrea", "@fechaCrea", "@precioUnitarioOrigen", "@idVendedor2", "@identrada", "@nPFacturado", "@idLista"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Int}
        Dim valores() As Object = {d.IdAgencia, d.IdTipoDocumento, d.Serie, d.NumeroDocumento, d.Item, d.IdArticulo, d.Descripcion, d.Texto, d.Cantidad, d.Unidad, d.Serie1, d.Cantidad1, d.UnidadEnvase, d.NumeroEnvase, d.SaldoEntrega, d.PrecioVenta, d.PrecioVentaH, d.PrecioVentaImportacion, d.PrecioVentaImportacionH, d.PrecioSIGV, d.ImporteDescuento, d.DescuentoDocumento, d.CargoDistribucion, d.IGV, d.ImporteIGV, d.ImporteUS, d.ImporteMN, d.IdTipoITemDescuento, d.Descuento1, d.ImporteDescuento1, d.Descuento2, d.ImporteDescuento2, d.Descuento3, d.ImporteDescuento3, d.Descuento4, d.ImporteDescuento4, d.Descuento5, d.ImporteDescuento5, d.Descuento6, d.Estado, d.Vendedor, d.IdAlmacen, d.NumeroCaja, d.Stock, d.FechaSDocumento, d.IdLinea, d.IdCampania, d.NumeroPaquete, d.NroDescuentoFinaciero, d.NroDescuentoLaboratorio, d.NroDescuentoAdicional, d.NroDescuentoBonificacion, d.NroDescuentoFlag, d.Comision, d.ImporteComision, d.UsuarioCrea, d.FechaCrea, d.PrecioUnitarioOrigen, d.IdVendedor2, d.identrada, d.NPFacturado, d.IdLista}
        sql.EjecutarProcedure("Str_DetallePreVenta_I", parametros, valores, tipoParametro, 62)
    End Sub
    Public Sub Actualizar(d As NDetallePreVenta)
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@item", "@idArticulo", "@descripcion", "@texto", "@cantidad", "@unidad", "@serie1", "@cantidad1", "@unidadEnvase", "@numeroEnvase", "@saldoEntrega", "@precioVenta", "@precioVentaH", "@precioVentaImportacion", "@precioVentaImportacionH", "@precioSIGV", "@importeDescuento", "@descuentoDocumento", "@cargoDistribucion", "@iGV", "@importeIGV", "@importeUS", "@importeMN", "@idTipoITemDescuento", "@descuento1", "@importeDescuento1", "@descuento2", "@importeDescuento2", "@descuento3", "@importeDescuento3", "@descuento4", "@importeDescuento4", "@descuento5", "@importeDescuento5", "@descuento6", "@estado", "@vendedor", "@idAlmacen", "@numeroCaja", "@stock", "@fechaSDocumento", "@idLinea", "@idCampania", "@numeroPaquete", "@nroDescuentoFinaciero", "@nroDescuentoLaboratorio", "@nroDescuentoAdicional", "@nroDescuentoBonificacion", "@nroDescuentoFlag", "@comision", "@importeComision", "@usuarioCrea", "@fechaCrea", "@precioUnitarioOrigen", "@idVendedor2", "@identrada", "@nPFacturado", "@idLista"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Int}
        Dim valores() As Object = {d.IdAgencia, d.IdTipoDocumento, d.Serie, d.NumeroDocumento, d.Item, d.IdArticulo, d.Descripcion, d.Texto, d.Cantidad, d.Unidad, d.Serie1, d.Cantidad1, d.UnidadEnvase, d.NumeroEnvase, d.SaldoEntrega, d.PrecioVenta, d.PrecioVentaH, d.PrecioVentaImportacion, d.PrecioVentaImportacionH, d.PrecioSIGV, d.ImporteDescuento, d.DescuentoDocumento, d.CargoDistribucion, d.IGV, d.ImporteIGV, d.ImporteUS, d.ImporteMN, d.IdTipoITemDescuento, d.Descuento1, d.ImporteDescuento1, d.Descuento2, d.ImporteDescuento2, d.Descuento3, d.ImporteDescuento3, d.Descuento4, d.ImporteDescuento4, d.Descuento5, d.ImporteDescuento5, d.Descuento6, d.Estado, d.Vendedor, d.IdAlmacen, d.NumeroCaja, d.Stock, d.FechaSDocumento, d.IdLinea, d.IdCampania, d.NumeroPaquete, d.NroDescuentoFinaciero, d.NroDescuentoLaboratorio, d.NroDescuentoAdicional, d.NroDescuentoBonificacion, d.NroDescuentoFlag, d.Comision, d.ImporteComision, d.UsuarioCrea, d.FechaCrea, d.PrecioUnitarioOrigen, d.IdVendedor2, d.identrada, d.NPFacturado, d.IdLista}
        sql.EjecutarProcedure("Str_DetallePreVenta_U", parametros, valores, tipoParametro, 62)
    End Sub
    Public Sub Eliminar(d As NDetallePreVenta)
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@item", "@idArticulo", "@idAlmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAgencia, d.IdTipoDocumento, d.Serie, d.NumeroDocumento, d.Item, d.IdArticulo, d.IdAlmacen}
        sql.EjecutarProcedure("Str_DetallePreVenta_D", parametros, valores, tipoParametro, 7)
    End Sub
    Public Function Lista(d As NDetallePreVenta) As DataTable
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@item", "@idArticulo", "@idAlmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAgencia, d.IdTipoDocumento, d.Serie, d.NumeroDocumento, d.Item, d.IdArticulo, d.IdAlmacen}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetallePreVenta_S", parametros, valores, tipoParametro, 7).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDetallePreVenta) As NDetallePreVenta
        Dim parametros() As Object = {"@idAgencia", "@idTipoDocumento", "@serie", "@numeroDocumento", "@item", "@idArticulo", "@idAlmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAgencia, d.IdTipoDocumento, d.Serie, d.NumeroDocumento, d.Item, d.IdArticulo, d.IdAlmacen}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetallePreVenta_S", parametros, valores, tipoParametro, 7).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdAgencia = IIf(dt.Rows(0).Item("idAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAgencia"))
            d.IdTipoDocumento = IIf(dt.Rows(0).Item("idTipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoDocumento"))
            d.Serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.NumeroDocumento = IIf(dt.Rows(0).Item("numeroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento"))
            d.Item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.IdArticulo = IIf(dt.Rows(0).Item("idArticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idArticulo"))
            d.Descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.Texto = IIf(dt.Rows(0).Item("texto") Is DBNull.Value, Nothing, dt.Rows(0).Item("texto"))
            d.Cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.Unidad = IIf(dt.Rows(0).Item("unidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidad"))
            d.Serie1 = IIf(dt.Rows(0).Item("serie1") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie1"))
            d.Cantidad1 = IIf(dt.Rows(0).Item("cantidad1") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad1"))
            d.UnidadEnvase = IIf(dt.Rows(0).Item("unidadEnvase") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidadEnvase"))
            d.NumeroEnvase = IIf(dt.Rows(0).Item("numeroEnvase") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroEnvase"))
            d.SaldoEntrega = IIf(dt.Rows(0).Item("saldoEntrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldoEntrega"))
            d.PrecioVenta = IIf(dt.Rows(0).Item("precioVenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioVenta"))
            d.PrecioVentaH = IIf(dt.Rows(0).Item("precioVentaH") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioVentaH"))
            d.PrecioVentaImportacion = IIf(dt.Rows(0).Item("precioVentaImportacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioVentaImportacion"))
            d.PrecioVentaImportacionH = IIf(dt.Rows(0).Item("precioVentaImportacionH") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioVentaImportacionH"))
            d.PrecioSIGV = IIf(dt.Rows(0).Item("precioSIGV") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioSIGV"))
            d.ImporteDescuento = IIf(dt.Rows(0).Item("importeDescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeDescuento"))
            d.DescuentoDocumento = IIf(dt.Rows(0).Item("descuentoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuentoDocumento"))
            d.CargoDistribucion = IIf(dt.Rows(0).Item("cargoDistribucion") Is DBNull.Value, Nothing, dt.Rows(0).Item("cargoDistribucion"))
            d.IGV = IIf(dt.Rows(0).Item("iGV") Is DBNull.Value, Nothing, dt.Rows(0).Item("iGV"))
            d.ImporteIGV = IIf(dt.Rows(0).Item("importeIGV") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeIGV"))
            d.ImporteUS = IIf(dt.Rows(0).Item("importeUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeUS"))
            d.ImporteMN = IIf(dt.Rows(0).Item("importeMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeMN"))
            d.IdTipoITemDescuento = IIf(dt.Rows(0).Item("idTipoITemDescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoITemDescuento"))
            d.Descuento1 = IIf(dt.Rows(0).Item("descuento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento1"))
            d.ImporteDescuento1 = IIf(dt.Rows(0).Item("importeDescuento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeDescuento1"))
            d.Descuento2 = IIf(dt.Rows(0).Item("descuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento2"))
            d.ImporteDescuento2 = IIf(dt.Rows(0).Item("importeDescuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeDescuento2"))
            d.Descuento3 = IIf(dt.Rows(0).Item("descuento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento3"))
            d.ImporteDescuento3 = IIf(dt.Rows(0).Item("importeDescuento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeDescuento3"))
            d.Descuento4 = IIf(dt.Rows(0).Item("descuento4") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento4"))
            d.ImporteDescuento4 = IIf(dt.Rows(0).Item("importeDescuento4") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeDescuento4"))
            d.Descuento5 = IIf(dt.Rows(0).Item("descuento5") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento5"))
            d.ImporteDescuento5 = IIf(dt.Rows(0).Item("importeDescuento5") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeDescuento5"))
            d.Descuento6 = IIf(dt.Rows(0).Item("descuento6") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento6"))
            d.Estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.Vendedor = IIf(dt.Rows(0).Item("vendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("vendedor"))
            d.IdAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.NumeroCaja = IIf(dt.Rows(0).Item("numeroCaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroCaja"))
            d.Stock = IIf(dt.Rows(0).Item("stock") Is DBNull.Value, Nothing, dt.Rows(0).Item("stock"))
            d.FechaSDocumento = IIf(dt.Rows(0).Item("fechaSDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaSDocumento"))
            d.IdLinea = IIf(dt.Rows(0).Item("idLinea") Is DBNull.Value, Nothing, dt.Rows(0).Item("idLinea"))
            d.IdCampania = IIf(dt.Rows(0).Item("idCampania") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCampania"))
            d.NumeroPaquete = IIf(dt.Rows(0).Item("numeroPaquete") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroPaquete"))
            d.NroDescuentoFinaciero = IIf(dt.Rows(0).Item("nroDescuentoFinaciero") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroDescuentoFinaciero"))
            d.NroDescuentoLaboratorio = IIf(dt.Rows(0).Item("nroDescuentoLaboratorio") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroDescuentoLaboratorio"))
            d.NroDescuentoAdicional = IIf(dt.Rows(0).Item("nroDescuentoAdicional") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroDescuentoAdicional"))
            d.NroDescuentoBonificacion = IIf(dt.Rows(0).Item("nroDescuentoBonificacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroDescuentoBonificacion"))
            d.NroDescuentoFlag = IIf(dt.Rows(0).Item("nroDescuentoFlag") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroDescuentoFlag"))
            d.Comision = IIf(dt.Rows(0).Item("comision") Is DBNull.Value, Nothing, dt.Rows(0).Item("comision"))
            d.ImporteComision = IIf(dt.Rows(0).Item("importeComision") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeComision"))
            d.UsuarioCrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.FechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.PrecioUnitarioOrigen = IIf(dt.Rows(0).Item("precioUnitarioOrigen") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioUnitarioOrigen"))
            d.IdVendedor2 = IIf(dt.Rows(0).Item("idVendedor2") Is DBNull.Value, Nothing, dt.Rows(0).Item("idVendedor2"))
            d.identrada = IIf(dt.Rows(0).Item("identrada") Is DBNull.Value, Nothing, dt.Rows(0).Item("identrada"))
            d.NPFacturado = IIf(dt.Rows(0).Item("nPFacturado") Is DBNull.Value, Nothing, dt.Rows(0).Item("nPFacturado"))
            d.IdLista = IIf(dt.Rows(0).Item("idLista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idLista"))
        Else
            d.Descripcion = Nothing
            d.Texto = Nothing
            d.Cantidad = Nothing
            d.Unidad = Nothing
            d.Serie1 = Nothing
            d.Cantidad1 = Nothing
            d.UnidadEnvase = Nothing
            d.NumeroEnvase = Nothing
            d.SaldoEntrega = Nothing
            d.PrecioVenta = Nothing
            d.PrecioVentaH = Nothing
            d.PrecioVentaImportacion = Nothing
            d.PrecioVentaImportacionH = Nothing
            d.PrecioSIGV = Nothing
            d.ImporteDescuento = Nothing
            d.DescuentoDocumento = Nothing
            d.CargoDistribucion = Nothing
            d.IGV = Nothing
            d.ImporteIGV = Nothing
            d.ImporteUS = Nothing
            d.ImporteMN = Nothing
            d.IdTipoITemDescuento = Nothing
            d.Descuento1 = Nothing
            d.ImporteDescuento1 = Nothing
            d.Descuento2 = Nothing
            d.ImporteDescuento2 = Nothing
            d.Descuento3 = Nothing
            d.ImporteDescuento3 = Nothing
            d.Descuento4 = Nothing
            d.ImporteDescuento4 = Nothing
            d.Descuento5 = Nothing
            d.ImporteDescuento5 = Nothing
            d.Descuento6 = Nothing
            d.Estado = Nothing
            d.Vendedor = Nothing
            d.NumeroCaja = Nothing
            d.Stock = Nothing
            d.FechaSDocumento = Nothing
            d.IdLinea = Nothing
            d.IdCampania = Nothing
            d.NumeroPaquete = Nothing
            d.NroDescuentoFinaciero = Nothing
            d.NroDescuentoLaboratorio = Nothing
            d.NroDescuentoAdicional = Nothing
            d.NroDescuentoBonificacion = Nothing
            d.NroDescuentoFlag = Nothing
            d.Comision = Nothing
            d.ImporteComision = Nothing
            d.UsuarioCrea = Nothing
            d.FechaCrea = Nothing
            d.PrecioUnitarioOrigen = Nothing
            d.IdVendedor2 = Nothing
            d.identrada = Nothing
            d.NPFacturado = Nothing
            d.IdLista = Nothing
        End If
        Return d
    End Function

#End Region
End Class
