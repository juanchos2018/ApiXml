Imports CapaDatos
Public Class NCambioMoneda
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _idMoneda As String
    Private _fechaCambio As String
    Private _compra As Decimal
    Private _xMNIMP As Decimal
    Private _venta As Decimal
    Private _xMNIMP2 As Decimal
    Private _fechaRegistro As System.DateTime
    Private _fechaCambio2 As System.DateTime

#End Region

#Region "Properties"

    Public Property IdMoneda As String
        Get
            Return _idMoneda
        End Get
        Set
            _idMoneda = Value
        End Set
    End Property

    Public Property FechaCambio As String
        Get
            Return _fechaCambio
        End Get
        Set
            _fechaCambio = Value
        End Set
    End Property

    Public Property Compra As Decimal
        Get
            Return _compra
        End Get
        Set
            _compra = Value
        End Set
    End Property

    Public Property XMNIMP As Decimal
        Get
            Return _xMNIMP
        End Get
        Set
            _xMNIMP = Value
        End Set
    End Property

    Public Property Venta As Decimal
        Get
            Return _venta
        End Get
        Set
            _venta = Value
        End Set
    End Property

    Public Property XMNIMP2 As Decimal
        Get
            Return _xMNIMP2
        End Get
        Set
            _xMNIMP2 = Value
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

    Public Property FechaCambio2 As System.DateTime
        Get
            Return _fechaCambio2
        End Get
        Set
            _fechaCambio2 = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idMoneda As String, ByVal fechaCambio As String, ByVal compra As Decimal, ByVal xMNIMP As Decimal, ByVal venta As Decimal, ByVal xMNIMP2 As Decimal, ByVal fechaRegistro As System.DateTime, ByVal fechaCambio2 As System.DateTime)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    ''' <summary>
    ''' Agrega un registro en la tabla tipo de cambio
    ''' </summary>
    ''' <param name="d"></param>
    Public Sub Agregar(d As NCambioMoneda)
        Dim parametros() As Object = {"@idMoneda", "@fechaCambio", "@compra", "@xMNIMP", "@venta", "@xMNIMP2", "@fechaRegistro", "@fechaCambio2"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.DateTime}
        Dim valores() As Object = {d.IdMoneda, d.FechaCambio, d.Compra, d.XMNIMP, d.Venta, d.XMNIMP2, d.FechaRegistro, d.FechaCambio2}
        sql.EjecutarProcedure("Str_CambioMoneda_I", parametros, valores, tipoParametro, 8)
    End Sub
    ''' <summary>
    ''' Actualiza los cambios realizados en la tabla tipo de cambio
    ''' </summary>
    ''' <param name="d"></param>
    Public Sub Actualizar(d As NCambioMoneda)
        Dim parametros() As Object = {"@idMoneda", "@fechaCambio", "@compra", "@xMNIMP", "@venta", "@xMNIMP2", "@fechaRegistro", "@fechaCambio2"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.DateTime}
        Dim valores() As Object = {d.IdMoneda, d.FechaCambio, d.Compra, d.XMNIMP, d.Venta, d.XMNIMP2, d.FechaRegistro, d.FechaCambio2}
        sql.EjecutarProcedure("Str_CambioMoneda_U", parametros, valores, tipoParametro, 8)
    End Sub
    ''' <summary>
    ''' Elimina un registro de la tabla tipo de cambio
    ''' </summary>
    ''' <param name="d"></param>
    Public Sub Eliminar(d As NCambioMoneda)
        Dim parametros() As Object = {"@idMoneda", "@fechaCambio2"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.IdMoneda, d.FechaCambio2}
        sql.EjecutarProcedure("Str_CambioMoneda_D", parametros, valores, tipoParametro, 2)
    End Sub
    ''' <summary>
    ''' Obtiene un registro de la tabla tipo de cambio
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    Public Function Item(d As NCambioMoneda) As NCambioMoneda
        Dim parametros() As Object = {"@idMoneda", "@fechaCambio2"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.IdMoneda, d.FechaCambio2}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_CambioMoneda_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdMoneda = dt.Rows(0).Item("IdMoneda")
            d.FechaCambio2 = dt.Rows(0).Item("FechaCambio2")
            d.FechaCambio = dt.Rows(0).Item("FechaCambio")
            d.Compra = dt.Rows(0).Item("Compra")
            d.Venta = dt.Rows(0).Item("Venta")
            '  d.FechaRegistro = dt.Rows(0).Item("FechaRegistro")
        Else
            d.IdMoneda = Nothing
            d.Compra = 0
            d.Venta = 0

        End If
        Return d
    End Function
    Public Function Existe_CambioMoneda(d As NCambioMoneda) As Boolean
        Dim parametros() As Object = {"@idmoneda", "@fechacambio2"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.DateTime}
        Dim valores() As Object = {d.IdMoneda, d.FechaCambio2}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_CambioMoneda", parametros, valores, tipoParametro, 2)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
#End Region
End Class
