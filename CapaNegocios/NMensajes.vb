Imports CapaDatos
Public Class NMensajes
    Dim sql As New ClsConexion
    Public Property articulo() As String
        Get
            Return m_articulo
        End Get
        Set
            m_articulo = Value
        End Set
    End Property
    Private m_articulo As String
    Public Property cantidad() As String
        Get
            Return m_cantidad
        End Get
        Set
            m_cantidad = Value
        End Set
    End Property
    Private m_cantidad As String
    Public Property NroDocumento() As String
        Get
            Return m_NroDocumento
        End Get
        Set
            m_NroDocumento = Value
        End Set
    End Property
    Private m_NroDocumento As String
    Public Property id() As Integer
        Get
            Return m_id
        End Get
        Set
            m_id = Value
        End Set
    End Property
    Private m_id As Integer
    Public Property fechacrea() As DateTime
        Get
            Return m_fechaCrea
        End Get
        Set
            m_fechaCrea = Value
        End Set
    End Property
    Private m_fechacrea As DateTime
    Public Property idarticulo() As String
        Get
            Return m_idarticulo
        End Get
        Set
            m_idarticulo = Value
        End Set
    End Property
    Private m_idarticulo As String
    Public Property item() As Integer
        Get
            Return m_item
        End Get
        Set
            m_item = Value
        End Set
    End Property
    Private m_item As Integer
    Public Property precioAnterior() As [Decimal]
        Get
            Return m_precioAnterior
        End Get
        Set
            m_precioAnterior = Value
        End Set
    End Property
    Private m_precioAnterior As [Decimal]

    Public Function ListaCompras() As DataTable
        Dim parametros As Object() = New Object() {}
        Dim tipoParametro As SqlDbType() = New SqlDbType() {}
        Dim valores As Object() = New Object() {}
        Dim dt As New DataTable()
        dt = sql.ProcedureSQL("Str_Movimiento_ultimo", parametros, valores, tipoParametro, 0).Tables(0)
        Return dt
    End Function
    Public Function ListaUltimosPrecios() As DataTable
        Dim parametros As Object() = New Object() {}
        Dim tipoParametro As SqlDbType() = New SqlDbType() {}
        Dim valores As Object() = New Object() {}
        Dim dt As New DataTable()
        dt = sql.ProcedureSQL("Str_UltimoPrecioCtalogo", parametros, valores, tipoParametro, 0).Tables(0)
        Return dt
    End Function

End Class
