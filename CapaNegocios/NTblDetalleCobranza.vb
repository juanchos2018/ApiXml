Imports CapaDatos
Public Class NTblDetalleCobranza
    Dim sql As New ClsConexion
#Region "Declarations"
    Private _tipoLiq As String
    Private _nroLiq As String
    Private _item As String
    Private _tipoDocumento As String
    Private _nroDocumento As String
    Private _idCliente As String
    Private _idCaja As String
    Private _formaPago As String
    Private _fechaMovimiento As System.DateTime
    Private _importeCobrado As Decimal
    Private _glosa As String
    Private _usuarioCrea As String
    Private _fechaCrea As System.DateTime
    Private _usuarioMod As String
    Private _fechaMod As System.DateTime
    Private _tipoDocRef As String
    Private _nroDocRef As String
    Private _estado As String
    Private _tipoCambio As Decimal
    Private _idMoneda As String
    Private _impCobMN As Decimal
    Private _impCobUS As Decimal
    Private _idAlmacen As String
    Private _idSubdiario As String
    Private _nroComprobante As String
    Private _idBanco As String
    Private _idMonedaO As String
    Private _idDetUsuarioCaja As Integer
    Public Property IdArea As String


#End Region

#Region "Properties"

    Public Property TipoLiq As String
        Get
            Return _tipoLiq
        End Get
        Set
            _tipoLiq = Value
        End Set
    End Property

    Public Property NroLiq As String
        Get
            Return _nroLiq
        End Get
        Set
            _nroLiq = Value
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

    Public Property IdCliente As String
        Get
            Return _idCliente
        End Get
        Set
            _idCliente = Value
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

    Public Property FormaPago As String
        Get
            Return _formaPago
        End Get
        Set
            _formaPago = Value
        End Set
    End Property

    Public Property FechaMovimiento As System.DateTime
        Get
            Return _fechaMovimiento
        End Get
        Set
            _fechaMovimiento = Value
        End Set
    End Property

    Public Property ImporteCobrado As Decimal
        Get
            Return _importeCobrado
        End Get
        Set
            _importeCobrado = Value
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

    Public Property TipoDocRef As String
        Get
            Return _tipoDocRef
        End Get
        Set
            _tipoDocRef = Value
        End Set
    End Property

    Public Property NroDocRef As String
        Get
            Return _nroDocRef
        End Get
        Set
            _nroDocRef = Value
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

    Public Property ImpCobMN As Decimal
        Get
            Return _impCobMN
        End Get
        Set
            _impCobMN = Value
        End Set
    End Property

    Public Property ImpCobUS As Decimal
        Get
            Return _impCobUS
        End Get
        Set
            _impCobUS = Value
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

    Public Property IdBanco As String
        Get
            Return _idBanco
        End Get
        Set
            _idBanco = Value
        End Set
    End Property

    Public Property IdMonedaO As String
        Get
            Return _idMonedaO
        End Get
        Set
            _idMonedaO = Value
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


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal tipoLiq As String, ByVal nroLiq As String, ByVal item As String, ByVal tipoDocumento As String, ByVal nroDocumento As String, ByVal idCliente As String, ByVal idCaja As String, ByVal formaPago As String, ByVal fechaMovimiento As System.DateTime, ByVal importeCobrado As Decimal, ByVal glosa As String, ByVal usuarioCrea As String, ByVal fechaCrea As System.DateTime, ByVal usuarioMod As String, ByVal fechaMod As System.DateTime, ByVal tipoDocRef As String, ByVal nroDocRef As String, ByVal estado As String, ByVal tipoCambio As Decimal, ByVal idMoneda As String, ByVal impCobMN As Decimal, ByVal impCobUS As Decimal, ByVal idAlmacen As String, ByVal idSubdiario As String, ByVal nroComprobante As String, ByVal idBanco As String, ByVal idMonedaO As String, ByVal idDetUsuarioCaja As Integer, idArea As String)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NTblDetalleCobranza)

        Dim parametros() As Object = {"@estado", "@fechaCrea", "@fechaMod", "@fechaMovimiento", "@formaPago", "@glosa", "@idAlmacen", "@idArea", "@idBanco", "@idCaja", "@idCliente", "@idDetUsuarioCaja", "@idMoneda", "@idMonedaO", "@idSubdiario", "@importeCobrado", "@item", "@nroComprobante", "@nroDocRef", "@nroDocumento", "@nroLiq", "@tipoCambio", "@tipoDocRef", "@tipoDocumento", "@tipoLiq", "@usuarioCrea", "@usuarioMod"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.Estado, d.FechaCrea, d.FechaMod, d.FechaMovimiento, d.FormaPago, d.Glosa, d.IdAlmacen, d.IdArea, d.IdBanco, d.IdCaja, d.IdCliente, d.IdDetUsuarioCaja, d.IdMoneda, d.IdMonedaO, d.IdSubdiario, d.ImporteCobrado, d.Item, d.NroComprobante, d.NroDocRef, d.NroDocumento, d.NroLiq, d.TipoCambio, d.TipoDocRef, d.TipoDocumento, d.TipoLiq, d.UsuarioCrea, d.UsuarioMod}
        sql.EjecutarProcedure("Str_Tbl_DetalleCobranza_I", parametros, valores, tipoParametro, 27)
    End Sub
    Public Sub Actualizar(d As NTblDetalleCobranza)
        Dim parametros() As Object = {"@estado", "@fechaCrea", "@fechaMod", "@fechaMovimiento", "@formaPago", "@glosa", "@idAlmacen", "@idArea", "@idBanco", "@idCaja", "@idCliente", "@idDetUsuarioCaja", "@idMoneda", "@idMonedaO", "@idSubdiario", "@importeCobrado", "@item", "@nroComprobante", "@nroDocRef", "@nroDocumento", "@nroLiq", "@tipoCambio", "@tipoDocRef", "@tipoDocumento", "@tipoLiq", "@usuarioCrea", "@usuarioMod"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.Estado, d.FechaCrea, d.FechaMod, d.FechaMovimiento, d.FormaPago, d.Glosa, d.IdAlmacen, d.IdArea, d.IdBanco, d.IdCaja, d.IdCliente, d.IdDetUsuarioCaja, d.IdMoneda, d.IdMonedaO, d.IdSubdiario, d.ImporteCobrado, d.Item, d.NroComprobante, d.NroDocRef, d.NroDocumento, d.NroLiq, d.TipoCambio, d.TipoDocRef, d.TipoDocumento, d.TipoLiq, d.UsuarioCrea, d.UsuarioMod}
        sql.EjecutarProcedure("Str_Tbl_DetalleCobranza_U", parametros, valores, tipoParametro, 27)
    End Sub
    Public Sub Eliminar(d As NTblDetalleCobranza)
        Dim parametros() As Object = {"@item", "@nroLiq", "@tipoLiq"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.Item, d.NroLiq, d.TipoLiq}
        sql.EjecutarProcedure("Str_Tbl_DetalleCobranza_D", parametros, valores, tipoParametro, 3)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@tipoLiq", "@nroLiq", "@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_DetalleCobranza_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTblDetalleCobranza) As DataTable
        Dim parametros() As Object = {"@item", "@nroLiq", "@tipoLiq"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.Item, d.NroLiq, d.TipoLiq}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_DetalleCobranza_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTblDetalleCobranza) As NTblDetalleCobranza
        Dim parametros() As Object = {"@item", "@nroLiq", "@tipoLiq"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.Item, d.NroLiq, d.TipoLiq}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_DetalleCobranza_S", parametros, valores, tipoParametro, 3).Tables(0)
        If dt.Rows.Count > 0 Then
            d.Estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.FechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.FechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.FechaMovimiento = IIf(dt.Rows(0).Item("fechaMovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMovimiento"))
            d.FormaPago = IIf(dt.Rows(0).Item("formaPago") Is DBNull.Value, Nothing, dt.Rows(0).Item("formaPago"))
            d.Glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.IdAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.IdArea = IIf(dt.Rows(0).Item("idArea") Is DBNull.Value, Nothing, dt.Rows(0).Item("idArea"))
            d.IdBanco = IIf(dt.Rows(0).Item("idBanco") Is DBNull.Value, Nothing, dt.Rows(0).Item("idBanco"))
            d.IdCaja = IIf(dt.Rows(0).Item("idCaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCaja"))
            d.IdCliente = IIf(dt.Rows(0).Item("idCliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCliente"))
            d.IdDetUsuarioCaja = IIf(dt.Rows(0).Item("idDetUsuarioCaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idDetUsuarioCaja"))
            d.IdMoneda = IIf(dt.Rows(0).Item("idMoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMoneda"))
            d.IdMonedaO = IIf(dt.Rows(0).Item("idMonedaO") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMonedaO"))
            d.IdSubdiario = IIf(dt.Rows(0).Item("idSubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idSubdiario"))
            d.ImpCobMN = IIf(dt.Rows(0).Item("impCobMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("impCobMN"))
            d.ImpCobUS = IIf(dt.Rows(0).Item("impCobUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("impCobUS"))
            d.ImporteCobrado = IIf(dt.Rows(0).Item("importeCobrado") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeCobrado"))
            d.Item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.NroComprobante = IIf(dt.Rows(0).Item("nroComprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroComprobante"))
            d.NroDocRef = IIf(dt.Rows(0).Item("nroDocRef") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroDocRef"))
            d.NroDocumento = IIf(dt.Rows(0).Item("nroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroDocumento"))
            d.NroLiq = IIf(dt.Rows(0).Item("nroLiq") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroLiq"))
            d.TipoCambio = IIf(dt.Rows(0).Item("tipoCambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoCambio"))
            d.TipoDocRef = IIf(dt.Rows(0).Item("tipoDocRef") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoDocRef"))
            d.TipoDocumento = IIf(dt.Rows(0).Item("tipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoDocumento"))
            d.TipoLiq = IIf(dt.Rows(0).Item("tipoLiq") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoLiq"))
            d.UsuarioCrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.UsuarioMod = IIf(dt.Rows(0).Item("usuarioMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioMod"))
        Else
            d.Estado = Nothing
            d.FechaCrea = Nothing
            d.FechaMod = Nothing
            d.FechaMovimiento = Nothing
            d.FormaPago = Nothing
            d.Glosa = Nothing
            d.IdAlmacen = Nothing
            d.IdArea = Nothing
            d.IdBanco = Nothing
            d.IdCaja = Nothing
            d.IdCliente = Nothing
            d.IdDetUsuarioCaja = Nothing
            d.IdMoneda = Nothing
            d.IdMonedaO = Nothing
            d.IdSubdiario = Nothing
            d.ImpCobMN = Nothing
            d.ImpCobUS = Nothing
            d.ImporteCobrado = Nothing
            d.Item = Nothing
            d.NroComprobante = Nothing
            d.NroDocRef = Nothing
            d.NroDocumento = Nothing
            d.NroLiq = Nothing
            d.TipoCambio = Nothing
            d.TipoDocRef = Nothing
            d.TipoDocumento = Nothing
            d.TipoLiq = Nothing
            d.UsuarioCrea = Nothing
            d.UsuarioMod = Nothing
        End If
        Return d
    End Function

    Public Function generaasiento(d As NTblDetalleCobranza) As DataTable
        Dim cadena As String = " SELECT     dcb.TipoLiq, dcb.NroLiq, dcb.Item, dcb.TipoDocumento, LEFT(dcb.NroDocumento, 4) AS Serie, SUBSTRING(dcb.NroDocumento, 5, 7) AS Numero, dcb.NroDocumento, dcb.IdCliente, dcb.IdCaja,  "
        cadena += " dcb.FormaPago, dcb.FechaMovimiento, dcb.ImporteCobrado, dcb.Glosa, dcb.UsuarioCrea, dcb.FechaCrea, dcb.UsuarioMod, dcb.FechaMod, dcb.TipoDocRef, dcb.NroDocRef, dcb.Estado,  "
        cadena += " dcb.TipoCambio, dcb.Id,Moneda, dcb.ImpCobMN, dcb.ImpCobUS, dcb.IdAlmacen, dcb.IdSubdiario, dcb.NroComprobante, dcb.IdBanco, dcb.IdMonedaO, dcb.IdDetUsuarioCaja, du.SubDiarioIngreso,  "
        cadena += " du.IdCuenta AS IdCuentaCja, Cliente.Nombre,isnull(n.idcuentaC,de.IdCuenta) as idcuentaC FROM Tbl_DetalleCobranza AS dcb INNER JOIN Tbl_DetalleUsuarioCaja AS du ON dcb.IdDetUsuarioCaja = du.IdDetUsuarioCaja INNER JOIN "
        cadena += " Cliente ON dcb.IdCliente = Cliente.IdCliente left join numeracion n on dcb.TipoDocumento=n.idtipodocumento and LEFT(dcb.NroDocumento, 4)=n.serie "
        cadena += " left join deuda de on dcb.TipoDocumento=de.IdTipoDocumento and dcb.NroDocumento=de.NumeroDocumento "
        cadena += " WHERE  IdAlmacen='" & d.IdAlmacen & "' and TipoDocumento='" & d.TipoDocumento & "' and Nrodocumento='" & d.NroDocumento & "' and dcb.IdCliente='" & d.IdCliente & "' and NroDocRef='" & d.NroDocRef & "'"
        Return sql.EjecutarConsulta("d", cadena).Tables(0)
    End Function
    Public Function GeneraPlanilla(id As String, idplanilla As String) As DataTable
        Dim cadena As String = " SELECT     dcb.TipoLiq, dcb.NroLiq, dcb.Item, dcb.TipoDocumento, LEFT(dcb.NroDocumento, 4) AS Serie, SUBSTRING(dcb.NroDocumento, 5, 7) AS Numero, dcb.NroDocumento, dcb.IdCliente, dcb.IdCaja,  "
        cadena += " dcb.FormaPago, dcb.FechaMovimiento, dcb.ImporteCobrado, dcb.Glosa, dcb.UsuarioCrea, dcb.FechaCrea, dcb.UsuarioMod, dcb.FechaMod, dcb.TipoDocRef, dcb.NroDocRef, dcb.Estado,  "
        cadena += " dcb.TipoCambio, dcb.IdMoneda, dcb.ImpCobMN, dcb.ImpCobUS, dcb.IdAlmacen, dcb.IdSubdiario, dcb.NroComprobante, dcb.IdBanco, dcb.IdMonedaO, dcb.IdDetUsuarioCaja, du.SubDiarioIngreso,  "
        cadena += " du.IdCuenta AS IdCuentaCja, Cliente.Nombre,isnull(de.IdCuenta,n.idcuentaC) as idcuentaC  FROM Tbl_DetalleCobranza AS dcb left JOIN Tbl_DetalleUsuarioCaja AS du ON dcb.IdDetUsuarioCaja = du.IdDetUsuarioCaja INNER JOIN "
        cadena += " Cliente ON dcb.IdCliente = Cliente.IdCliente left join numeracion n on dcb.TipoDocumento=n.idtipodocumento and LEFT(dcb.NroDocumento, 4)=n.serie "
        cadena += " left join deuda de on dcb.TipoDocumento=de.IdTipoDocumento and dcb.NroDocumento=de.NumeroDocumento "
        cadena += " WHERE dcb.Nroliq='" & idplanilla & "' and dcb.tipoliq='" & id & "'"
        Return sql.EjecutarConsulta("d", cadena).Tables(0)
        'cadena += " du.IdCuenta AS IdCuentaCja, Cliente.Nombre,isnull(n.idcuentaC,de.IdCuenta) as idcuentaC  FROM Tbl_DetalleCobranza AS dcb left JOIN Tbl_DetalleUsuarioCaja AS du ON dcb.IdDetUsuarioCaja = du.IdDetUsuarioCaja INNER JOIN "
        'cadena += " du.IdCuenta AS IdCuentaCja, Cliente.Nombre,n.idcuentaC  FROM Tbl_DetalleCobranza AS dcb INNER JOIN Tbl_DetalleUsuarioCaja AS du ON dcb.IdDetUsuarioCaja = du.IdDetUsuarioCaja INNER JOIN "
    End Function

    Public Function Genera_asiento_cobros(idalm As String, sb As String, Fi As Date, Ff As Date, ismanual As Boolean) As DataTable
        Dim parametros() As Object = {"@idAlmacen", "@IdSubdiario", "@FechaI", "@FechaF", "@AsientoManual"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Bit}
        Dim valores() As Object = {idalm, sb, Fi, Ff, ismanual}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Genera_AsientoCobranza", parametros, valores, tipoParametro, 5).Tables(0)
        Return dt
    End Function

#End Region


End Class
