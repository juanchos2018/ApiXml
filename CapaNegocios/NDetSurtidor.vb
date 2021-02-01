Imports CapaDatos
Public Class NDetSurtidor
    Dim sql As New ClsConexion
#Region "Declarations"
    Private _item_DS As Integer
    Private _idSurtidor As Integer
    Private _item_D As Integer
    Private _contInicial As Decimal
    Private _contFinal As Decimal
    Private _precioUnitario As Decimal
    Private _gal_Des As Decimal
    Private _total As Decimal

#End Region

#Region "Properties"

    Public Property Item_DS As Integer
        Get
            Return _item_DS
        End Get
        Set
            _item_DS = Value
        End Set
    End Property

    Public Property IdSurtidor As Integer
        Get
            Return _idSurtidor
        End Get
        Set
            _idSurtidor = Value
        End Set
    End Property

    Public Property Item_D As Integer
        Get
            Return _item_D
        End Get
        Set
            _item_D = Value
        End Set
    End Property

    Public Property ContInicial As Decimal
        Get
            Return _contInicial
        End Get
        Set
            _contInicial = Value
        End Set
    End Property

    Public Property ContFinal As Decimal
        Get
            Return _contFinal
        End Get
        Set
            _contFinal = Value
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

    Public Property Gal_Des As Decimal
        Get
            Return _gal_Des
        End Get
        Set
            _gal_Des = Value
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


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal item_DS As Integer, ByVal idSurtidor As Integer, ByVal item_D As Integer, ByVal contInicial As Decimal, ByVal contFinal As Decimal, ByVal precioUnitario As Decimal, ByVal gal_Des As Decimal, ByVal total As Decimal)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NDetSurtidor)
        Dim parametros() As Object = {"@idSurtidor", "@item_D", "@contInicial", "@contFinal", "@precioUnitario", "@gal_Des", "@total"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Int, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal}
        Dim valores() As Object = {d.IdSurtidor, d.Item_D, d.ContInicial, d.ContFinal, d.PrecioUnitario, d.Gal_Des, d.Total}
        sql.EjecutarProcedure("Str_tbl_Det_Surtidor_I", parametros, valores, tipoParametro, 7)
    End Sub
    Public Sub Actualizar(d As NDetSurtidor)
        Dim parametros() As Object = {"@item_DS", "@idSurtidor", "@item_D", "@contInicial", "@contFinal", "@precioUnitario", "@gal_Des", "@total"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Int, SqlDbType.Int, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal}
        Dim valores() As Object = {d.Item_DS, d.IdSurtidor, d.Item_D, d.ContInicial, d.ContFinal, d.PrecioUnitario, d.Gal_Des, d.Total}
        sql.EjecutarProcedure("Str_tbl_Det_Surtidor_U", parametros, valores, tipoParametro, 8)
    End Sub
    Public Sub Eliminar(d As NDetSurtidor)
        Dim parametros() As Object = {"@item_DS"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.Item_DS}
        sql.EjecutarProcedure("Str_tbl_Det_Surtidor_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Registro(d As NDetSurtidor) As NDetSurtidor
        Dim parametros() As Object = {"@item_DS"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.Item_DS}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Det_Surtidor_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.Item_DS = dt.Rows(0).Item("Item_DS")
            d.IdSurtidor = dt.Rows(0).Item("IdSurtidor")
            d.Item_D = dt.Rows(0).Item("Item_D")
            d.ContInicial = dt.Rows(0).Item("ContInicial")
            d.ContFinal = dt.Rows(0).Item("ContFinal")
            d.PrecioUnitario = dt.Rows(0).Item("PrecioUnitario")
            d.Gal_Des = dt.Rows(0).Item("Gal_Des")
            d.Total = dt.Rows(0).Item("Total")
        End If
        Return d
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@item_DS"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Det_Surtidor_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    ''' <summary>
    ''' Obtiene los registro relacionados a la tabla maestro
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    Public Function DetalleMaestro(d As NDetSurtidor) As DataTable
        Dim parametros() As Object = {"@IdSurtidor"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.IdSurtidor}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Det_Surtidor_M", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function

    Public Function Lista_Det_Lado(idsurtidor As String, lado As String) As DataTable
        Dim c As String = " SELECT     s.IdSurtidor, s.item, s.IdCaja, s.Nombre, ts.Descripcion, ts.NroMangueras, dts.IdArticulo, dts.Lado, ds.Item_DS, ds.Item_D, ds.producto_sem "
        c += " FROM         tbl_Surtidor AS s INNER JOIN tbl_Tipo_Surtidor AS ts ON s.item = ts.Item INNER JOIN "
        c += " tbl_Det_Surtidor AS ds ON s.IdSurtidor = ds.IdSurtidor INNER JOIN tbl_Det_Tipo_Surtidor AS dts ON ts.Item = dts.Item AND ds.Item_D = dts.Item_D "
        c += " where s.idsurtidor=" & idsurtidor & " and lado='" & lado & "'"
        Return sql.EjecutarConsulta("d", c).Tables(0)
    End Function
#End Region
End Class
