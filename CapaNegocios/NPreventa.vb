Imports CapaDatos
Public Class NPreventa
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
    Public Property IsFacturado As Boolean = False

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
#Region "Metodos"
    'Public Function Lista(p As NPreventa) As DataTable
    '    Dim cadena As String = "select distinct p.IdTipoDocumento,p.Serie,p.NumeroDocumento,p.FechaDocumento,p.IdCliente,p.Nombrecliente,p.importetotal,p.IdAlmacen,d.IdAgencia,isnull(isfacturado,0)as IsFacturado from PreVenta p "
    '    cadena += " inner join detallepreventa d on p.idalmacen=d.idalmacen and p.idtipodocumento=d.idtipodocumento and p.serie=d.serie and p.numerodocumento=d.numerodocumento "
    '    cadena += "where saldoentrega<>0 and p.idAlmacen='" & p.IdAlmacen & "'"
    '    Dim dt As New DataTable
    '    dt = px.EjecutarConsulta("Pedido", cadena).Tables(0)
    '    Return dt
    'End Function
    Public Function Lista(p As NPreventa) As DataTable
        Dim cadena As String = "select p.IdTipoDocumento,p.Serie,p.NumeroDocumento,p.FechaDocumento,p.IdCliente,p.Nombrecliente,p.importetotal,p.IdAlmacen,p.IdAgencia,isnull(isfacturado,0)as IsFacturado from PreVenta p "
        cadena += "where p.idAlmacen='" & p.IdAlmacen & "'"
        Dim dt As New DataTable
        dt = px.EjecutarConsulta("Pedido", cadena).Tables(0)
        Return dt
    End Function


    Public Sub Agregar(p As NPreventa)
        Dim params() As Object = {"@IdAgencia", "@IdTipoDocumento", "@Serie", "@NumeroDocumento", "@FechaDocumento", "@DebeHaber", "@IdVendedor", "@IdCliente", "@NombreCliente", "@DirCliente", "@RUC", "@idAlmacen", "@IdFormaVenta", "@IdMoneda", "@TipoCambio", "@IGV", "@ImporteTotal", "@ImporteIGV", "@DescuentoTotal", "@IdTipoDocumento1", "@Serie1", "@NumeroDocumento2", "@Observacion", "@Usuario", "@Fecha", "@FacturaGuia"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Float, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char}
        Dim vals() As Object
        vals = {p.IdAgencia, p.IdTipoDocumento, p.Serie, p.NumeroDocumento, p.FechaDocumento, "D", p.IdVendedor, p.IdCliente, p.NombreCliente, p.Direccion _
            , p.RUC, p.IdAlmacen, p.IdFormaVenta, p.IdMoneda, p.TipoCambio, p.IGV, p.ImporteTotal, p.ImporteIGV, p.ImporteDescuento, "", "", "", p.Descripcion, p.UsuarioCrea, p.FechaCrea, "'S'"}
        px.EjecutarProcedure("proc_AddPreVenta", params, vals, tipoParametro, 26)
    End Sub
    Public Sub AgregarG(d As NPreventa)
        Dim parametros() As Object = {"@anuladoNC", "@debeHaber", "@descripcion", "@descuento2", "@descuneto1", "@destino", "@direccion", "@embalaje", "@estado", "@facturaGuia", "@fechaCrea", "@fechaDocumento", "@fechaMod", "@fechaVencimineto", "@flete", "@idAgencia", "@idAlmacen", "@idAnexo", "@idAutorizacion", "@idCadena", "@idCaja", "@idCentroCosto", "@idchofer", "@idCliente", "@idFormaVenta", "@idInternoCadena", "@idMaquina", "@idMoneda", "@idTipoAnexo", "@idTipoDocumento", "@idTipoDocumento1", "@idTipoFactura", "@idTipoNotaCredito", "@idTransportista", "@idUsuarioOperador", "@idUsuarioSectorista", "@idVendedor", "@idVendedor1", "@idZonaVenta", "@iGV", "@importeDescuento", "@importeIGV", "@importeTotal", "@impreso", "@isFacturado", "@linea", "@nombreCliente", "@numeroDocumento", "@numeroDocumento1", "@numeroOrden", "@numeroPedido", "@reparto", "@rUC", "@saldo", "@serie", "@serie1", "@tasa", "@tipoCambio", "@usuarioCrea", "@usuarioMod"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.AnuladoNC, d.DebeHaber, d.Descripcion, d.Descuento2, d.Descuneto1, d.Destino, d.Direccion, d.Embalaje, d.Estado, d.FacturaGuia, d.FechaCrea, d.FechaDocumento, d.FechaMod, d.FechaVencimineto, d.Flete, d.IdAgencia, d.IdAlmacen, d.IdAnexo, d.IdAutorizacion, d.IdCadena, d.IdCaja, d.IdCentroCosto, d.Idchofer, d.IdCliente, d.IdFormaVenta, d.IdInternoCadena, d.IdMaquina, d.IdMoneda, d.IdTipoAnexo, d.IdTipoDocumento, d.IdTipoDocumento1, d.IdTipoFactura, d.IdTipoNotaCredito, d.IdTransportista, d.IdUsuarioOperador, d.IdUsuarioSectorista, d.IdVendedor, d.IdVendedor1, d.IdZonaVenta, d.IGV, d.ImporteDescuento, d.ImporteIGV, d.ImporteTotal, d.Impreso, d.IsFacturado, d.Linea, d.NombreCliente, d.NumeroDocumento, d.NumeroDocumento1, d.NumeroOrden, d.NumeroPedido, d.Reparto, d.RUC, d.Saldo, d.Serie, d.Serie1, d.Tasa, d.TipoCambio, d.UsuarioCrea, d.UsuarioMod}
        px.EjecutarProcedure("Str_Preventa_I", parametros, valores, tipoParametro, 60)
    End Sub

    Public Sub Actualizar(d As NPreventa)
        Dim parametros() As Object = {"@anuladoNC", "@debeHaber", "@descripcion", "@descuento2", "@descuneto1", "@destino", "@direccion", "@embalaje", "@estado", "@facturaGuia", "@fechaCrea", "@fechaDocumento", "@fechaMod", "@fechaVencimineto", "@flete", "@idAgencia", "@idAlmacen", "@idAnexo", "@idAutorizacion", "@idCadena", "@idCaja", "@idCentroCosto", "@idchofer", "@idCliente", "@idFormaVenta", "@idInternoCadena", "@idMaquina", "@idMoneda", "@idTipoAnexo", "@idTipoDocumento", "@idTipoDocumento1", "@idTipoFactura", "@idTipoNotaCredito", "@idTransportista", "@idUsuarioOperador", "@idUsuarioSectorista", "@idVendedor", "@idVendedor1", "@idZonaVenta", "@iGV", "@importeDescuento", "@importeIGV", "@importeTotal", "@impreso", "@linea", "@nombreCliente", "@numeroDocumento", "@numeroDocumento1", "@numeroOrden", "@numeroPedido", "@reparto", "@rUC", "@saldo", "@serie", "@serie1", "@tasa", "@tipoCambio", "@usuarioCrea", "@usuarioMod", "@IsFacturado"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.AnuladoNC, d.DebeHaber, d.Descripcion, d.Descuento2, d.Descuneto1, d.Destino, d.Direccion, d.Embalaje, d.Estado, d.FacturaGuia, d.FechaCrea, d.FechaDocumento, d.FechaMod, d.FechaVencimineto, d.Flete, d.IdAgencia, d.IdAlmacen, d.IdAnexo, d.IdAutorizacion, d.IdCadena, d.IdCaja, d.IdCentroCosto, d.Idchofer, d.IdCliente, d.IdFormaVenta, d.IdInternoCadena, d.IdMaquina, d.IdMoneda, d.IdTipoAnexo, d.IdTipoDocumento, d.IdTipoDocumento1, d.IdTipoFactura, d.IdTipoNotaCredito, d.IdTransportista, d.IdUsuarioOperador, d.IdUsuarioSectorista, d.IdVendedor, d.IdVendedor1, d.IdZonaVenta, d.IGV, d.ImporteDescuento, d.ImporteIGV, d.ImporteTotal, d.Impreso, d.Linea, d.NombreCliente, d.NumeroDocumento, d.NumeroDocumento1, d.NumeroOrden, d.NumeroPedido, d.Reparto, d.RUC, d.Saldo, d.Serie, d.Serie1, d.Tasa, d.TipoCambio, d.UsuarioCrea, d.UsuarioMod, d.IsFacturado}
        px.EjecutarProcedure("Str_Preventa_U", parametros, valores, tipoParametro, 60)
    End Sub
    Public Sub Eliminar(d As NPreventa)
        Dim parametros() As Object = {"@idAgencia", "@idAlmacen", "@idTipoDocumento", "@numeroDocumento", "@serie"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAgencia, d.IdAlmacen, d.IdTipoDocumento, d.NumeroDocumento, d.Serie}
        px.EjecutarProcedure("Str_Preventa_D", parametros, valores, tipoParametro, 5)
    End Sub
    Public Function ListaG() As DataTable
        Dim parametros() As Object = {"@idAgencia", "@idAlmacen", "@idTipoDocumento", "@numeroDocumento", "@serie"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = px.ProcedureSQL("Str_Preventa_S", parametros, valores, tipoParametro, 5).Tables(0)
        ' dt = px.Proc_BindSource("Str_Preventa_S", parametros, valores, tipoParametro, 5).DataSource
        Return dt
    End Function
    Public Function ListaG(d As NPreventa) As DataTable
        Dim parametros() As Object = {"@idAgencia", "@idAlmacen", "@idTipoDocumento", "@numeroDocumento", "@serie"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAgencia, d.IdAlmacen, d.IdTipoDocumento, d.NumeroDocumento, d.Serie}
        Dim dt As New DataTable
        dt = px.ProcedureSQL("Str_Preventa_S", parametros, valores, tipoParametro, 5).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NPreventa) As NPreventa
        Dim parametros() As Object = {"@idAgencia", "@idAlmacen", "@idTipoDocumento", "@numeroDocumento", "@serie"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAgencia, d.IdAlmacen, d.IdTipoDocumento, d.NumeroDocumento, d.Serie}
        Dim dt As New DataTable
        dt = px.ProcedureSQL("Str_Preventa_S", parametros, valores, tipoParametro, 5).Tables(0)
        If dt.Rows.Count > 0 Then
            d.AnuladoNC = IIf(dt.Rows(0).Item("anuladoNC") Is DBNull.Value, Nothing, dt.Rows(0).Item("anuladoNC"))
            d.DebeHaber = IIf(dt.Rows(0).Item("debeHaber") Is DBNull.Value, Nothing, dt.Rows(0).Item("debeHaber"))
            d.Descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.Descuento2 = IIf(dt.Rows(0).Item("descuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento2"))
            d.Descuneto1 = IIf(dt.Rows(0).Item("descuneto1") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuneto1"))
            d.Destino = IIf(dt.Rows(0).Item("destino") Is DBNull.Value, Nothing, dt.Rows(0).Item("destino"))
            d.Direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.Embalaje = IIf(dt.Rows(0).Item("embalaje") Is DBNull.Value, Nothing, dt.Rows(0).Item("embalaje"))
            d.Estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.FacturaGuia = IIf(dt.Rows(0).Item("facturaGuia") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaGuia"))
            d.FechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.FechaDocumento = IIf(dt.Rows(0).Item("fechaDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaDocumento"))
            d.FechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.FechaVencimineto = IIf(dt.Rows(0).Item("fechaVencimineto") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaVencimineto"))
            d.Flete = IIf(dt.Rows(0).Item("flete") Is DBNull.Value, Nothing, dt.Rows(0).Item("flete"))
            d.IdAgencia = IIf(dt.Rows(0).Item("idAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAgencia"))
            d.IdAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.IdAnexo = IIf(dt.Rows(0).Item("idAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAnexo"))
            d.IdAutorizacion = IIf(dt.Rows(0).Item("idAutorizacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAutorizacion"))
            d.IdCadena = IIf(dt.Rows(0).Item("idCadena") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCadena"))
            d.IdCaja = IIf(dt.Rows(0).Item("idCaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCaja"))
            d.IdCentroCosto = IIf(dt.Rows(0).Item("idCentroCosto") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCentroCosto"))
            d.Idchofer = IIf(dt.Rows(0).Item("idchofer") Is DBNull.Value, Nothing, dt.Rows(0).Item("idchofer"))
            d.IdCliente = IIf(dt.Rows(0).Item("idCliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCliente"))
            d.IdFormaVenta = IIf(dt.Rows(0).Item("idFormaVenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idFormaVenta"))
            d.IdInternoCadena = IIf(dt.Rows(0).Item("idInternoCadena") Is DBNull.Value, Nothing, dt.Rows(0).Item("idInternoCadena"))
            d.IdMaquina = IIf(dt.Rows(0).Item("idMaquina") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMaquina"))
            d.IdMoneda = IIf(dt.Rows(0).Item("idMoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMoneda"))
            d.IdTipoAnexo = IIf(dt.Rows(0).Item("idTipoAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoAnexo"))
            d.IdTipoDocumento = IIf(dt.Rows(0).Item("idTipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoDocumento"))
            d.IdTipoDocumento1 = IIf(dt.Rows(0).Item("idTipoDocumento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoDocumento1"))
            d.IdTipoFactura = IIf(dt.Rows(0).Item("idTipoFactura") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoFactura"))
            d.IdTipoNotaCredito = IIf(dt.Rows(0).Item("idTipoNotaCredito") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoNotaCredito"))
            d.IdTransportista = IIf(dt.Rows(0).Item("idTransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTransportista"))
            d.IdUsuarioOperador = IIf(dt.Rows(0).Item("idUsuarioOperador") Is DBNull.Value, Nothing, dt.Rows(0).Item("idUsuarioOperador"))
            d.IdUsuarioSectorista = IIf(dt.Rows(0).Item("idUsuarioSectorista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idUsuarioSectorista"))
            d.IdVendedor = IIf(dt.Rows(0).Item("idVendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idVendedor"))
            d.IdVendedor1 = IIf(dt.Rows(0).Item("idVendedor1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idVendedor1"))
            d.IdZonaVenta = IIf(dt.Rows(0).Item("idZonaVenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idZonaVenta"))
            d.IGV = IIf(dt.Rows(0).Item("iGV") Is DBNull.Value, Nothing, dt.Rows(0).Item("iGV"))
            d.ImporteDescuento = IIf(dt.Rows(0).Item("importeDescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeDescuento"))
            d.ImporteIGV = IIf(dt.Rows(0).Item("importeIGV") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeIGV"))
            d.ImporteTotal = IIf(dt.Rows(0).Item("importeTotal") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeTotal"))
            d.Impreso = IIf(dt.Rows(0).Item("impreso") Is DBNull.Value, Nothing, dt.Rows(0).Item("impreso"))
            d.Linea = IIf(dt.Rows(0).Item("linea") Is DBNull.Value, Nothing, dt.Rows(0).Item("linea"))
            d.NombreCliente = IIf(dt.Rows(0).Item("nombreCliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreCliente"))
            d.NumeroDocumento = IIf(dt.Rows(0).Item("numeroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento"))
            d.NumeroDocumento1 = IIf(dt.Rows(0).Item("numeroDocumento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento1"))
            d.NumeroOrden = IIf(dt.Rows(0).Item("numeroOrden") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroOrden"))
            d.NumeroPedido = IIf(dt.Rows(0).Item("numeroPedido") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroPedido"))
            d.Reparto = IIf(dt.Rows(0).Item("reparto") Is DBNull.Value, Nothing, dt.Rows(0).Item("reparto"))
            d.RUC = IIf(dt.Rows(0).Item("rUC") Is DBNull.Value, Nothing, dt.Rows(0).Item("rUC"))
            d.Saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.Serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.Serie1 = IIf(dt.Rows(0).Item("serie1") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie1"))
            d.Tasa = IIf(dt.Rows(0).Item("tasa") Is DBNull.Value, Nothing, dt.Rows(0).Item("tasa"))
            d.TipoCambio = IIf(dt.Rows(0).Item("tipoCambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoCambio"))
            d.UsuarioCrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.UsuarioMod = IIf(dt.Rows(0).Item("usuarioMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioMod"))
            d.IsFacturado = IIf(dt.Rows(0).Item("IsFacturado") Is DBNull.Value, Nothing, dt.Rows(0).Item("IsFacturado"))
        Else
            d.AnuladoNC = Nothing
            d.DebeHaber = Nothing
            d.Descripcion = Nothing
            d.Descuento2 = Nothing
            d.Descuneto1 = Nothing
            d.Destino = Nothing
            d.Direccion = Nothing
            d.Embalaje = Nothing
            d.Estado = Nothing
            d.FacturaGuia = Nothing
            d.FechaCrea = Nothing
            d.FechaDocumento = Nothing
            d.FechaMod = Nothing
            d.FechaVencimineto = Nothing
            d.Flete = Nothing
            d.IdAgencia = Nothing
            d.IdAlmacen = Nothing
            d.IdAnexo = Nothing
            d.IdAutorizacion = Nothing
            d.IdCadena = Nothing
            d.IdCaja = Nothing
            d.IdCentroCosto = Nothing
            d.Idchofer = Nothing
            d.IdCliente = Nothing
            d.IdFormaVenta = Nothing
            d.IdInternoCadena = Nothing
            d.IdMaquina = Nothing
            d.IdMoneda = Nothing
            d.IdTipoAnexo = Nothing
            d.IdTipoDocumento = Nothing
            d.IdTipoDocumento1 = Nothing
            d.IdTipoFactura = Nothing
            d.IdTipoNotaCredito = Nothing
            d.IdTransportista = Nothing
            d.IdUsuarioOperador = Nothing
            d.IdUsuarioSectorista = Nothing
            d.IdVendedor = Nothing
            d.IdVendedor1 = Nothing
            d.IdZonaVenta = Nothing
            d.IGV = Nothing
            d.ImporteDescuento = Nothing
            d.ImporteIGV = Nothing
            d.ImporteTotal = Nothing
            d.Impreso = Nothing
            d.Linea = Nothing
            d.NombreCliente = Nothing
            d.NumeroDocumento = Nothing
            d.NumeroDocumento1 = Nothing
            d.NumeroOrden = Nothing
            d.NumeroPedido = Nothing
            d.Reparto = Nothing
            d.RUC = Nothing
            d.Saldo = Nothing
            d.Serie = Nothing
            d.Serie1 = Nothing
            d.Tasa = Nothing
            d.TipoCambio = Nothing
            d.UsuarioCrea = Nothing
            d.UsuarioMod = Nothing
            d.IsFacturado = False
        End If
        Return d
    End Function

    Public Function Existe(p As NPreventa) As Boolean
        Dim existeC As String
        Dim bandera As Boolean = False
        Dim valoresC() As Object = {"'" & p.IdAgencia & "'", "'" & p.IdTipoDocumento & "'", "'" & Trim(p.Serie) & "'", "'" & Trim(p.NumeroDocumento) & "'"}
        existeC = px.ValorEscalar("dbo.FPreVenta_Existe", valoresC, 4)
        If existeC = "1" Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function ObtenerCabecera(p As NPreventa) As NPreventa
        Dim dt_cab As New DataTable
        Dim cadena As String = " SELECT     IdAgencia, IdTipoDocumento, Serie, NumeroDocumento, NumeroPedido, FechaDocumento, FechaVencimineto, DebeHaber, IdVendedor, IdCaja, IdCliente,  "
        cadena += " NombreCliente, Direccion, RUC, IdAlmacen, IdFormaVenta, IdMoneda, TipoCambio, ImporteTotal, ImporteIGV, Saldo, ImporteDescuento, ISNULL(NumeroOrden,'')AS NumeroOrden,  "
        cadena += " IdTipoDocumento1, Serie1, NumeroDocumento1, Descripcion, Estado, FacturaGuia, IdTransportista, IdCentroCosto, IdMaquina, Destino, IdTipoFactura, IdTipoAnexo, "
        cadena += " IdAnexo, Descuneto1, Descuento2, Flete, Embalaje, Tasa, IdUsuarioOperador, IdUsuarioSectorista, IdCadena, IdInternoCadena, IdAutorizacion, Reparto, UsuarioCrea, "
        cadena += " FechaCrea, UsuarioMod, FechaMod, IdTipoNotaCredito, Linea, Impreso, AnuladoNC, IdVendedor1, IGV, Idchofer, IdZonaVenta  "
        cadena += "  FROM PreVenta "
        cadena += " where IdAgencia='" & p.IdAgencia & "' AND IdAlmacen='" & p.IdAlmacen & "' and Serie='" & p.Serie & "' and IdTipoDocumento='" & p.IdTipoDocumento & "' and NumeroDocumento='" & p.NumeroDocumento & "' "
        dt_cab = px.EjecutarConsulta("cab", cadena).Tables(0)
        If dt_cab.Rows.Count > 0 Then
            With dt_cab
                p.IdAgencia = .Rows(0).Item("IdAgencia").ToString
                p.IdVendedor = .Rows(0).Item("IdVendedor").ToString
                p.IdCliente = .Rows(0).Item("IdCliente").ToString
                p.NombreCliente = .Rows(0).Item("NombreCliente").ToString
                p.Direccion = .Rows(0).Item("Direccion").ToString
                p.RUC = .Rows(0).Item("RUC").ToString
                p.IdFormaVenta = .Rows(0).Item("IdFormaVenta").ToString
                p.IdMoneda = .Rows(0).Item("IdMoneda")
                p.TipoCambio = .Rows(0).Item("TipoCambio")
                p.ImporteTotal = .Rows(0).Item("ImporteTotal")
                p.ImporteIGV = .Rows(0).Item("ImporteIGV")
                p.FechaDocumento = .Rows(0).Item("FechaDocumento")
                p.Descripcion = .Rows(0).Item("Descripcion").ToString
                p.IGV = .Rows(0).Item("IGV")
                p.Idchofer = .Rows(0).Item("IdChofer").ToString
                p.NumeroOrden = .Rows(0).Item("NumeroOrden").ToString
            End With
        End If
        Return p
    End Function

    Public Function documentos(idalmacen As String) As DataTable
        Dim cadena As String = " select IdAlmacen,IdTipoDocumento,Serie,NumeroDocumento,FechaDocumento,IdCliente,NombreCliente,ImporteTotal from preventa "
        cadena += " where isnull(Estado,'V')='V'  and idalmacen='" & idalmacen & "'"
        Dim dt As DataTable = px.EjecutarConsulta("ca", cadena).Tables(0)
        Return dt
    End Function

#End Region




#End Region



End Class
