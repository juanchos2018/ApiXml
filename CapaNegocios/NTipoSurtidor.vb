Imports CapaDatos

Public Class NTipoSurtidor
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _item As Integer
    Private _descripcion As String
    Private _nroMangueras As Integer
    Private _usuarioCrea As String
    Private _fechaCrea As System.DateTime

#End Region

#Region "Properties"

    Public Property Item As Integer
        Get
            Return _item
        End Get
        Set
            _item = Value
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


    Public Property NroMangueras As Integer
        Get
            Return _nroMangueras
        End Get
        Set
            _nroMangueras = Value
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


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal item As Integer, ByVal descripcion As String, ByVal nroMangueras As Integer, ByVal usuarioCrea As String, ByVal fechaCrea As System.DateTime)
        Me.New()
    End Sub

#End Region
#Region "Metodos"
    ''' <summary>
    ''' Agrega un item a tipo surtidor
    ''' </summary>
    ''' <param name="d"></param>
    Public Sub Agregar(d As NTipoSurtidor)
        Dim parametros() As Object = {"@descripcion", "@nroMangueras", "@usuarioCrea", "@fechaCrea"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.Descripcion, d.NroMangueras, d.UsuarioCrea, d.FechaCrea}
        sql.EjecutarProcedure("Str_tbl_Tipo_Surtidor_I", parametros, valores, tipoParametro, 4)
    End Sub
    ''' <summary>
    ''' Actualizar el tipo surtidor
    ''' </summary>
    ''' <param name="d"></param>
    Public Sub Actualizar(d As NTipoSurtidor)
        Dim parametros() As Object = {"@item", "@descripcion", "@nroMangueras", "@usuarioCrea", "@fechaCrea"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.Item, d.Descripcion, d.NroMangueras, d.UsuarioCrea, d.FechaCrea}
        sql.EjecutarProcedure("Str_tbl_Tipo_Surtidor_U", parametros, valores, tipoParametro, 5)
    End Sub
    ''' <summary>
    ''' Eliminar registro de la tabla
    ''' </summary>
    ''' <param name="d"></param>
    Public Sub Eliminar(d As NTipoSurtidor)
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.Item}
        sql.EjecutarProcedure("Str_tbl_Tipo_Surtidor_D", parametros, valores, tipoParametro, 1)
    End Sub
    ''' <summary>
    ''' Obtiene un registro de tipo surtidor
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    Public Function Registro(d As NTipoSurtidor) As NTipoSurtidor
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.Item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Tipo_Surtidor_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.Descripcion = dt.Rows(0).Item("descripcion")
            d.NroMangueras = dt.Rows(0).Item("nroMangueras")
            d.UsuarioCrea = dt.Rows(0).Item("usuarioCrea")
            d.FechaCrea = dt.Rows(0).Item("fechaCrea")
        Else
            d.Descripcion = ""
            d.NroMangueras = 0
        End If
        Return d
    End Function
    ''' <summary>
    ''' Lista todos los tipos de surtidores existentes
    ''' </summary>
    ''' <returns></returns>
    Public Function Listar() As DataTable
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Tipo_Surtidor_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function

#End Region

End Class
