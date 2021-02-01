Imports CapaDatos
Public Class NSurtidor
    Dim sql As New ClsConexion
#Region "Declarations"
    Private _idSurtidor As Integer
    Private _item As Integer
    Private _idCaja As String
    Private _usuarioCrea As String
    Private _fechaCrea As System.DateTime
    Private _Nombre As String

#End Region

#Region "Properties"

    Public Property IdSurtidor As Integer
        Get
            Return _idSurtidor
        End Get
        Set
            _idSurtidor = Value
        End Set
    End Property

    Public Property item As Integer
        Get
            Return _item
        End Get
        Set
            _item = Value
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

    Public Property Nombre As String
        Get
            Return _Nombre
        End Get
        Set(value As String)
            _Nombre = value
        End Set
    End Property



#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idSurtidor As Integer, ByVal item As Integer, ByVal idCaja As String, ByVal usuarioCrea As String, ByVal fechaCrea As System.DateTime)
        Me.New()
    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NSurtidor)
        Dim parametros() As Object = {"@item", "@idCaja", "@usuarioCrea", "@fechaCrea", "@Nombre"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.item, d.IdCaja, d.UsuarioCrea, d.FechaCrea, d.Nombre}
        sql.EjecutarProcedure("Str_tbl_Surtidor_I", parametros, valores, tipoParametro, 5)
    End Sub
    Public Sub Actualizar(d As NSurtidor)
        Dim parametros() As Object = {"@idSurtidor", "@item", "@idCaja", "@usuarioCrea", "@fechaCrea", "@Nombre"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdSurtidor, d.item, d.IdCaja, d.UsuarioCrea, d.FechaCrea, d.Nombre}
        sql.EjecutarProcedure("Str_tbl_Surtidor_U", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Eliminar(d As NSurtidor)
        Dim parametros() As Object = {"@idSurtidor"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.IdSurtidor}
        sql.EjecutarProcedure("Str_tbl_Surtidor_D", parametros, valores, tipoParametro, 1)
    End Sub
    ''' <summary>
    ''' obtine un registro según la clave primaria
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    Public Function Registro(d As NSurtidor) As NSurtidor
        Dim parametros() As Object = {"@idSurtidor"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.IdSurtidor}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Surtidor_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdSurtidor = dt.Rows(0).Item("IdSurtidor")
            d.item = dt.Rows(0).Item("Item")
            d.IdCaja = dt.Rows(0).Item("IdCaja")
        Else

        End If
        Return d
    End Function
    ''' <summary>
    ''' Obtiene todos los registros de la tabla
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idSurtidor"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Surtidor_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function

    Public Function DetalleCaja(idcaja As String) As DataTable
        Dim dt As New DataTable
        Dim cadena As String = " SELECT     tbl_Det_Tipo_Surtidor.Lado, tbl_Det_Tipo_Surtidor.IdArticulo, Articulo.Descripcion1,tbl_Det_Surtidor.ContFinal as ContInicial,0.00 as ContFinal,tbl_Det_Surtidor.Gal_Des,0.0 as Gal_Credito,0.0 as Gal_Calibrados,0.0 as Gal_Contado, Articulo.Precio1, "
        cadena += " 0.0 AS Total,tbl_Surtidor.Nombre, tbl_Surtidor.IdSurtidor,tbl_Det_Surtidor.Item_DS "
        cadena += " FROM         tbl_Det_Surtidor INNER JOIN "
        cadena += " tbl_Surtidor ON tbl_Det_Surtidor.IdSurtidor = tbl_Surtidor.IdSurtidor INNER JOIN "
        cadena += " tbl_Tipo_Surtidor ON tbl_Surtidor.item = tbl_Tipo_Surtidor.Item INNER JOIN "
        cadena += " tbl_Det_Tipo_Surtidor ON tbl_Det_Surtidor.Item_D = tbl_Det_Tipo_Surtidor.Item_D AND tbl_Tipo_Surtidor.Item = tbl_Det_Tipo_Surtidor.Item INNER JOIN "
        cadena += " Articulo ON tbl_Det_Tipo_Surtidor.IdArticulo = Articulo.IdArticulo where  tbl_Surtidor.IdCaja='" & idcaja & "' "
        cadena += " ORDER BY tbl_Det_Tipo_Surtidor.Lado,tbl_Det_Tipo_Surtidor.IdArticulo,tbl_Surtidor.Nombre"
        dt = sql.EjecutarConsulta("c", cadena).Tables(0)
        Return dt
    End Function


#End Region

End Class
