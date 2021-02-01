Imports CapaDatos
Public Class NNumeracionGrifo
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _idNumeracion As Integer
    Private _idCaja As String
    Private _idSurtidor As Integer
    Private _idTipoDocumento As String
    Private _serie As String
    Private _numeroInicial As Decimal
    Private _numeroFinal As Decimal
    Private _descripcion As String
    Private _numeroActual As Decimal
    Private _usuarioCrea As String
    Private _fechaCrea As System.DateTime
    Private _idAgencia As String
    Private _items As Integer

#End Region

#Region "Properties"

    Public Property IdNumeracion As Integer
        Get
            Return _idNumeracion
        End Get
        Set
            _idNumeracion = Value
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

    Public Property IdSurtidor As Integer
        Get
            Return _idSurtidor
        End Get
        Set
            _idSurtidor = Value
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

    Public Property NumeroInicial As Decimal
        Get
            Return _numeroInicial
        End Get
        Set
            _numeroInicial = Value
        End Set
    End Property

    Public Property NumeroFinal As Decimal
        Get
            Return _numeroFinal
        End Get
        Set
            _numeroFinal = Value
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

    Public Property NumeroActual As Decimal
        Get
            Return _numeroActual
        End Get
        Set
            _numeroActual = Value
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

    Public Property IdAgencia As String
        Get
            Return _idAgencia
        End Get
        Set
            _idAgencia = Value
        End Set
    End Property

    Public Property items As Integer
        Get
            Return _items
        End Get
        Set
            _items = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idNumeracion As Integer, ByVal idCaja As String, ByVal idSurtidor As Integer, ByVal idTipoDocumento As String, ByVal serie As String, ByVal numeroInicial As Decimal, ByVal numeroFinal As Decimal, ByVal descripcion As String, ByVal numeroActual As Decimal, ByVal usuarioCrea As String, ByVal fechaCrea As System.DateTime, ByVal idAgencia As String, ByVal items As Integer)
        Me.New()
    End Sub

#Region "Metodos"
    Public Sub Agregar(d As NNumeracionGrifo)
        Dim parametros() As Object = {"@idCaja", "@idSurtidor", "@idTipoDocumento", "@serie", "@numeroInicial", "@numeroFinal", "@descripcion", "@numeroActual", "@usuarioCrea", "@fechaCrea", "@idAgencia", "@items"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Int}
        Dim valores() As Object = {d.IdCaja, d.IdSurtidor, d.IdTipoDocumento, d.Serie, d.NumeroInicial, d.NumeroFinal, d.Descripcion, d.NumeroActual, d.UsuarioCrea, d.FechaCrea, d.IdAgencia, d.items}
        sql.EjecutarProcedure("Str_tbl_Numeracion_Grifo_I", parametros, valores, tipoParametro, 12)
    End Sub
    Public Sub Actualizar(d As NNumeracionGrifo)
        Dim parametros() As Object = {"@idNumeracion", "@idCaja", "@idSurtidor", "@idTipoDocumento", "@serie", "@numeroInicial", "@numeroFinal", "@descripcion", "@numeroActual", "@usuarioCrea", "@fechaCrea", "@idAgencia", "@items"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Int}
        Dim valores() As Object = {d.IdNumeracion, d.IdCaja, d.IdSurtidor, d.IdTipoDocumento, d.Serie, d.NumeroInicial, d.NumeroFinal, d.Descripcion, d.NumeroActual, d.UsuarioCrea, d.FechaCrea, d.IdAgencia, d.items}
        sql.EjecutarProcedure("Str_tbl_Numeracion_Grifo_U", parametros, valores, tipoParametro, 13)
    End Sub
    Public Sub Eliminar(d As NNumeracionGrifo)
        Dim parametros() As Object = {"@idNumeracion"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.IdNumeracion}
        sql.EjecutarProcedure("Str_tbl_Numeracion_Grifo_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function registro(d As NNumeracionGrifo) As NNumeracionGrifo
        Dim parametros() As Object = {"@idNumeracion"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.IdNumeracion}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Numeracion_Grifo_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdSurtidor = dt.Rows(0).Item("IdSurtidor")
            d.IdCaja = dt.Rows(0).Item("IdCaja")
            d.IdTipoDocumento = dt.Rows(0).Item("IdTipoDocumento")
            d.IdAgencia = dt.Rows(0).Item("IdAgencia")
            d.items = dt.Rows(0).Item("items")
            d.NumeroActual = dt.Rows(0).Item("NumeroActual")
            d.NumeroInicial = dt.Rows(0).Item("NumeroInicial")
            d.NumeroFinal = dt.Rows(0).Item("NumeroFinal")
            d.Serie = dt.Rows(0).Item("Serie")
            d.Descripcion = dt.Rows(0).Item("Descripcion")
            d.UsuarioCrea = dt.Rows(0).Item("UsuarioCrea").ToString
            '  d.FechaCrea = dt.Rows(0).Item("FechaCrea")
        Else
            d.NumeroActual = 0
            d.NumeroInicial = 0
            d.NumeroFinal = 0
            d.Serie = ""
            d.Descripcion = ""
        End If
        Return d
    End Function

    Public Function lista() As DataTable
        Dim parametros() As Object = {"@idNumeracion"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Numeracion_Grifo_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function SerieTipo(d As NNumeracionGrifo) As DataTable
        Dim dt As New DataTable
        Dim s As String = " SELECT     IdNumeracion, IdCaja, IdSurtidor, IdTipoDocumento, Serie, NumeroInicial, NumeroFinal, Descripcion, NumeroActual, UsuarioCrea, FechaCrea, IdAgencia, items "
        s += " FROM         tbl_Numeracion_Grifo "
        s += " WHERE     (IdCaja='" & d.IdCaja & "') AND (IdSurtidor=" & d.IdSurtidor & ") AND (IdTipoDocumento='" & d.IdTipoDocumento & "')"
        dt = sql.EjecutarConsulta("d", s).Tables(0)
        Return dt
    End Function
    Public Function SerieItem(d As NNumeracionGrifo) As NNumeracionGrifo
        Dim dt As New DataTable
        Dim s As String = " SELECT     IdNumeracion, IdCaja, IdSurtidor, IdTipoDocumento, Serie, NumeroInicial, NumeroFinal, Descripcion, NumeroActual, UsuarioCrea, FechaCrea, IdAgencia, items "
        s += " FROM         tbl_Numeracion_Grifo "
        s += " WHERE     (IdCaja='" & d.IdCaja & "') AND (IdSurtidor=" & d.IdSurtidor & ") AND (IdTipoDocumento='" & d.IdTipoDocumento & "') AND (Serie='" & d.Serie & "')"
        dt = sql.EjecutarConsulta("d", s).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdNumeracion = dt.Rows(0).Item("IdNumeracion")
            d.IdCaja = dt.Rows(0).Item("IdCaja")
            d.IdSurtidor = dt.Rows(0).Item("IdSurtidor")
            d.IdTipoDocumento = dt.Rows(0).Item("IdTipoDocumento")
            d.Serie = dt.Rows(0).Item("Serie")
            d.NumeroActual = dt.Rows(0).Item("NumeroActual")
            d.NumeroInicial = dt.Rows(0).Item("NumeroInicial")
            d.NumeroFinal = dt.Rows(0).Item("NumeroFinal")
            d.items = dt.Rows(0).Item("Items")
            d.IdAgencia = dt.Rows(0).Item("IdAgencia")
        End If
        Return d
    End Function
    Public Function SerieItem(d As NNumeracionGrifo, sinsurtidor As Boolean) As NNumeracionGrifo
        Dim dt As New DataTable
        Dim text As String = " SELECT     IdNumeracion, IdCaja, IdSurtidor, IdTipoDocumento, Serie, NumeroInicial, NumeroFinal, Descripcion, NumeroActual, UsuarioCrea, FechaCrea, IdAgencia, items "
        text += " FROM         tbl_Numeracion_Grifo "
        text = " WHERE     (IdCaja='" & d.IdCaja & "') AND (IdTipoDocumento='" & d.IdTipoDocumento & "') AND (Serie='" & d.Serie & "')"
        dt = sql.EjecutarConsulta("d", text).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdNumeracion = dt.Rows(0).Item("IdNumeracion")
            d.IdCaja = dt.Rows(0).Item("IdCaja")
            d.IdSurtidor = dt.Rows(0).Item("IdSurtidor")
            d.IdTipoDocumento = dt.Rows(0).Item("IdTipoDocumento")
            d.Serie = dt.Rows(0).Item("Serie")
            d.NumeroActual = dt.Rows(0).Item("NumeroActual")
            d.NumeroInicial = dt.Rows(0).Item("NumeroInicial")
            d.NumeroFinal = dt.Rows(0).Item("NumeroFinal")
            d.items = dt.Rows(0).Item("Items")
            d.IdAgencia = dt.Rows(0).Item("IdAgencia")
        End If
        Return d
    End Function


    Public Sub actualizarnumero(d As NNumeracionGrifo)
        sql.Editar("tbl_Numeracion_Grifo", "NumeroInicial=" & d.NumeroInicial & ",NumeroActual=" & d.NumeroActual, "IdCaja='" & d.IdCaja & "' and IdSurtidor=" & d.IdSurtidor & " and idtipodocumento='" & d.IdTipoDocumento & "' and serie='" & d.Serie & "'")
    End Sub

#End Region



#End Region
End Class
