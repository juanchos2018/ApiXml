Imports System.Diagnostics.Eventing.Reader
Imports CapaDatos
Public Class NStock
    Dim sql As New ClsConexion
#Region "Declarations"
    Private _idAlmacen As String
    Private _idArticulo As String
    Private _stockDisponible As Decimal
    Private _stockComprometido As Decimal
    Private _stockMinimo As Decimal
    Private _stockMaximo As Decimal
    Private _mesProceso As String
    Private _precioUnitario As Decimal
    Private _ultFechaMov As System.DateTime
    Private _stockMes As Decimal
    Private _stockValorizado As Decimal
    Private _puntoReposicion As Decimal
    Private _semanaReposicion As Decimal
    Private _tipoReposicion As String
    Private _ubicacionFisica As String
    Private _ubicacionFisica2 As String
    Private _ubicacionFisica3 As String
    Private _ubicacionFisica4 As String
    Private _loteCompra As Decimal
    Private _tipoCompra As String
    Private _ingreso As Decimal
    Private _salida As Decimal
    Private _tipoAfectacion As String



#End Region

#Region "Properties"

    Public Property IdAlmacen As String
        Get
            Return _idAlmacen
        End Get
        Set
            _idAlmacen = Value
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

    Public Property StockDisponible As Decimal
        Get
            Return _stockDisponible
        End Get
        Set
            _stockDisponible = Value
        End Set
    End Property

    Public Property StockComprometido As Decimal
        Get
            Return _stockComprometido
        End Get
        Set
            _stockComprometido = Value
        End Set
    End Property

    Public Property StockMinimo As Decimal
        Get
            Return _stockMinimo
        End Get
        Set
            _stockMinimo = Value
        End Set
    End Property

    Public Property StockMaximo As Decimal
        Get
            Return _stockMaximo
        End Get
        Set
            _stockMaximo = Value
        End Set
    End Property

    Public Property MesProceso As String
        Get
            Return _mesProceso
        End Get
        Set
            _mesProceso = Value
        End Set
    End Property

    Public Property PrecioUnitario As Decimal
        Get
            Return _precioUnitario
        End Get
        Set
            _precioUnitario = Value
        End Set
    End Property

    Public Property UltFechaMov As System.DateTime
        Get
            Return _ultFechaMov
        End Get
        Set
            _ultFechaMov = Value
        End Set
    End Property

    Public Property StockMes As Decimal
        Get
            Return _stockMes
        End Get
        Set
            _stockMes = Value
        End Set
    End Property

    Public Property StockValorizado As Decimal
        Get
            Return _stockValorizado
        End Get
        Set
            _stockValorizado = Value
        End Set
    End Property

    Public Property PuntoReposicion As Decimal
        Get
            Return _puntoReposicion
        End Get
        Set
            _puntoReposicion = Value
        End Set
    End Property

    Public Property SemanaReposicion As Decimal
        Get
            Return _semanaReposicion
        End Get
        Set
            _semanaReposicion = Value
        End Set
    End Property

    Public Property TipoReposicion As String
        Get
            Return _tipoReposicion
        End Get
        Set
            _tipoReposicion = Value
        End Set
    End Property

    Public Property UbicacionFisica As String
        Get
            Return _ubicacionFisica
        End Get
        Set
            _ubicacionFisica = Value
        End Set
    End Property

    Public Property UbicacionFisica2 As String
        Get
            Return _ubicacionFisica2
        End Get
        Set
            _ubicacionFisica2 = Value
        End Set
    End Property

    Public Property UbicacionFisica3 As String
        Get
            Return _ubicacionFisica3
        End Get
        Set
            _ubicacionFisica3 = Value
        End Set
    End Property

    Public Property UbicacionFisica4 As String
        Get
            Return _ubicacionFisica4
        End Get
        Set
            _ubicacionFisica4 = Value
        End Set
    End Property

    Public Property LoteCompra As Decimal
        Get
            Return _loteCompra
        End Get
        Set
            _loteCompra = Value
        End Set
    End Property

    Public Property TipoCompra As String
        Get
            Return _tipoCompra
        End Get
        Set
            _tipoCompra = Value
        End Set
    End Property

    Public Property Ingreso As Decimal
        Get
            Return _ingreso
        End Get
        Set
            _ingreso = Value
        End Set
    End Property

    Public Property salida As Decimal
        Get
            Return _salida
        End Get
        Set
            _salida = Value
        End Set
    End Property

    Public Property tipoAfectacion As String
        Get
            Return _tipoAfectacion
        End Get
        Set(value As String)
            _tipoAfectacion = value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idAlmacen As String, ByVal idArticulo As String, ByVal stockDisponible As Decimal, ByVal stockComprometido As Decimal, ByVal stockMinimo As Decimal, ByVal stockMaximo As Decimal, ByVal mesProceso As String, ByVal precioUnitario As Decimal, ByVal ultFechaMov As System.DateTime, ByVal stockMes As Decimal, ByVal stockValorizado As Decimal, ByVal puntoReposicion As Decimal, ByVal semanaReposicion As Decimal, ByVal tipoReposicion As String, ByVal ubicacionFisica As String, ByVal ubicacionFisica2 As String, ByVal ubicacionFisica3 As String, ByVal ubicacionFisica4 As String, ByVal loteCompra As Decimal, ByVal tipoCompra As String, ByVal ingreso As Decimal, ByVal salida As Decimal)
        Me.New()
    End Sub

#End Region

#Region "Metodos"

    Public Sub agregar(s As NStock)
        Dim valsS() As Object = {s.IdAlmacen, s.IdArticulo, s.StockDisponible, s.UltFechaMov, s.tipoAfectacion}
        Dim paramsS() As Object = {"@IdAlmacen", "@IdArticulo", "@StockDisponible", "@UltFechaMov", "@Afectacion"}
        Dim tipoParametroS() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Float, SqlDbType.DateTime, SqlDbType.Char}
        sql.EjecutarProcedure("proc_Stock", paramsS, valsS, tipoParametroS, 5)
    End Sub
    Public Sub Actualizar(s As NStock)
        Dim valsS() As Object = {s.IdAlmacen, s.IdArticulo, s.StockDisponible, s.UltFechaMov, s.tipoAfectacion}
        Dim paramsS() As Object = {"@IdAlmacen", "@IdArticulo", "@StockDisponible", "@UltFechaMov", "@Afectacion"}
        Dim tipoParametroS() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Float, SqlDbType.DateTime, SqlDbType.Char}
        sql.EjecutarProcedure("proc_Stock", paramsS, valsS, tipoParametroS, 5)
    End Sub
    Public Function ActualizarStock(s As NStock)
        Dim valsS() As Object = {s.IdAlmacen, s.IdArticulo, s.StockDisponible, s.UltFechaMov}
        Dim paramsS() As Object = {"@IdAlmacen", "@IdArticulo", "@StockDisponible", "@UltFechaMov"}
        Dim tipoParametroS() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Float, SqlDbType.DateTime}
        sql.EjecutarProcedure("Str_StockDisponible_U", paramsS, valsS, tipoParametroS, 4)
    End Function
    Public Function ActualizarStockComprometido(s As NStock)
        Dim valsS() As Object = {s.IdAlmacen, s.IdArticulo, s.StockComprometido, s.UltFechaMov}
        Dim paramsS() As Object = {"@IdAlmacen", "@IdArticulo", "@StockComprometido", "@UltFechaMov"}
        Dim tipoParametroS() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Float, SqlDbType.DateTime}
        sql.EjecutarProcedure("Str_StockComprometido_U", paramsS, valsS, tipoParametroS, 4)
    End Function
    Public Function Listar(s As NStock) As DataTable
        Return sql.EjecutarConsulta("Stock", "Select StockDisponible,StockComprometido from stock where idArticulo = '" + s.IdArticulo + "' and idAlmacen = '" + s.IdAlmacen + "'").Tables(0)
    End Function
    Public Function ListaId(d As NStock) As NStock
        Dim valsS() As Object = {d.IdAlmacen, d.IdArticulo}
        Dim paramsS() As Object = {"@IdAlmacen", "@IdArticulo"}
        Dim tipoParametroS() As Object = {SqlDbType.Char, SqlDbType.Char}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_stock_S", paramsS, valsS, tipoParametroS, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdAlmacen = Convert.ToString(dt.Rows(0).Item("idalmacen"))
            d.IdArticulo = Convert.ToString(dt.Rows(0).Item("idarticulo"))
            d.StockDisponible = Convert.ToDecimal(dt.Rows(0).Item("stockDisponible"))
            If dt.Rows(0).Item("stockComprometido") Is DBNull.Value Then
                d.StockComprometido = 0
            Else
                d.StockComprometido = Convert.ToDecimal(dt.Rows(0).Item("stockComprometido").ToString())
            End If

            'd.StockMinimo = Convert.ToDecimal(dt.Rows(0).Item("stockMinimo"))
            'd.StockMaximo = Convert.ToDecimal(dt.Rows(0).Item("stockMaximo"))
            'd.MesProceso = Convert.ToString(dt.Rows(0).Item("mesproceso"))
            'd.PrecioUnitario = Convert.ToDecimal(dt.Rows(0).Item("precioUnitario"))
            'd.UltFechaMov = Convert.ToDateTime(dt.Rows(0).Item("ultFechaMov"))
            'd.StockMes = Convert.ToDecimal(dt.Rows(0).Item("stockMes"))
            'd.StockValorizado = Convert.ToDecimal(dt.Rows(0).Item("stockValorizado"))
            'd.PuntoReposicion = Convert.ToDecimal(dt.Rows(0).Item("puntoReposicion"))
            'd.SemanaReposicion = Convert.ToDecimal(dt.Rows(0).Item("semanaReposicion"))
            'd.TipoReposicion = Convert.ToString(dt.Rows(0).Item("tiporeposicion"))
            'd.UbicacionFisica = Convert.ToString(dt.Rows(0).Item("ubicacionfisica"))
            'd.UbicacionFisica2 = Convert.ToString(dt.Rows(0).Item("ubicacionfisica2"))
            'd.UbicacionFisica3 = Convert.ToString(dt.Rows(0).Item("ubicacionfisica3"))
            'd.UbicacionFisica4 = Convert.ToString(dt.Rows(0).Item("ubicacionfisica4"))
            'd.LoteCompra = Convert.ToDecimal(dt.Rows(0).Item("loteCompra"))
            'd.TipoCompra = Convert.ToString(dt.Rows(0).Item("tipocompra"))
            'd.Ingreso = Convert.ToDecimal(dt.Rows(0).Item("ingreso"))
            'd.salida = Convert.ToDecimal(dt.Rows(0).Item("salida"))
        Else

        End If
        Return d

    End Function

    Public Function item(s As NStock) As NStock
        Dim dt As New DataTable
        Dim valores() As Object = {s.IdAlmacen, s.IdArticulo.ToString.Trim}
        Dim campos() As Object = {"@IdAlmacen", "@IdArticulo"}
        Dim tipodatos() As Object = {SqlDbType.Char, SqlDbType.Char}
        dt = sql.ProcedureSQL("dbo.str_FndStock", campos, valores, tipodatos, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            Dim row As DataRow = dt.Rows(0)
            With row
                s.IdArticulo = .Item("Idarticulo").ToString.Trim
                s.IdAlmacen = .Item("IdAlmacen")
                s.StockDisponible = .Item("StockDisponible")
            End With
        Else
            s.StockDisponible = 0.0
        End If
        Return s
    End Function
    Public Function Existe(s As NStock) As Boolean
        Dim existeC As String
        Dim bandera As Boolean = False
        Dim valoresC() As Object = {"'" & s.IdAlmacen & "'", "'" & s.IdArticulo & "'"}
        existeC = sql.ValorEscalar("dbo.Stock_Existe", valoresC, 2)
        If existeC = "1" Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function

#End Region
End Class
