Imports CapaDatos
Public Class NAnexoComplementario_Con
    Dim sql As New ClsConexion
#Region "Declarations"
    Private _idTipoAnexo As String
    Private _idAnexo As String
    Private _apellidoPaterno As String
    Private _apellidoMaterno As String
    Private _nombre As String
    Private _formuSuspen As String
    Private _telefono As String
    Private _idTipoEntidad As String
    Private _fechaRegistro As System.DateTime
    Private _usuario As String
    Private _direccion As String
    Private _provincia As String
    Private _departamento As String
    Private _pais As String
    Private _zonaPostal As String
    Private _idDocumIden As String
    Private _numDocumIden As String
    Private _tipoProcedencia As String
    Private _tasaDetraccion As String
    Private _tasaPercepcion As String
    Private _dobleTributacion As String
    Private _tipoPension As String
    Private _bd As String
#End Region
#Region "Properties"
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

    Public Property ApellidoPaterno As String
        Get
            Return _apellidoPaterno
        End Get
        Set
            _apellidoPaterno = Value
        End Set
    End Property

    Public Property ApellidoMaterno As String
        Get
            Return _apellidoMaterno
        End Get
        Set
            _apellidoMaterno = Value
        End Set
    End Property

    Public Property Nombre As String
        Get
            Return _nombre
        End Get
        Set
            _nombre = Value
        End Set
    End Property

    Public Property FormuSuspen As String
        Get
            Return _formuSuspen
        End Get
        Set
            _formuSuspen = Value
        End Set
    End Property

    Public Property Telefono As String
        Get
            Return _telefono
        End Get
        Set
            _telefono = Value
        End Set
    End Property

    Public Property IdTipoEntidad As String
        Get
            Return _idTipoEntidad
        End Get
        Set
            _idTipoEntidad = Value
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

    Public Property Usuario As String
        Get
            Return _usuario
        End Get
        Set
            _usuario = Value
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

    Public Property Provincia As String
        Get
            Return _provincia
        End Get
        Set
            _provincia = Value
        End Set
    End Property

    Public Property Departamento As String
        Get
            Return _departamento
        End Get
        Set
            _departamento = Value
        End Set
    End Property

    Public Property Pais As String
        Get
            Return _pais
        End Get
        Set
            _pais = Value
        End Set
    End Property

    Public Property ZonaPostal As String
        Get
            Return _zonaPostal
        End Get
        Set
            _zonaPostal = Value
        End Set
    End Property

    Public Property IdDocumIden As String
        Get
            Return _idDocumIden
        End Get
        Set
            _idDocumIden = Value
        End Set
    End Property

    Public Property NumDocumIden As String
        Get
            Return _numDocumIden
        End Get
        Set
            _numDocumIden = Value
        End Set
    End Property

    Public Property TipoProcedencia As String
        Get
            Return _tipoProcedencia
        End Get
        Set
            _tipoProcedencia = Value
        End Set
    End Property

    Public Property TasaDetraccion As String
        Get
            Return _tasaDetraccion
        End Get
        Set
            _tasaDetraccion = Value
        End Set
    End Property

    Public Property TasaPercepcion As String
        Get
            Return _tasaPercepcion
        End Get
        Set
            _tasaPercepcion = Value
        End Set
    End Property

    Public Property DobleTributacion As String
        Get
            Return _dobleTributacion
        End Get
        Set
            _dobleTributacion = Value
        End Set
    End Property

    Public Property TipoPension As String
        Get
            Return _tipoPension
        End Get
        Set
            _tipoPension = Value
        End Set
    End Property

    Public Property Bd As String
        Get
            Return _bd
        End Get
        Set(value As String)
            _bd = value
        End Set
    End Property


#End Region

#Region "Constructors"
    Public Sub New()

    End Sub
    Public Sub New(ByVal idTipoAnexo As String, ByVal idAnexo As String, ByVal apellidoPaterno As String, ByVal apellidoMaterno As String, ByVal nombre As String, ByVal formuSuspen As String, ByVal telefono As String, ByVal idTipoEntidad As String, ByVal fechaRegistro As System.DateTime, ByVal usuario As String, ByVal direccion As String, ByVal provincia As String, ByVal departamento As String, ByVal pais As String, ByVal zonaPostal As String, ByVal idDocumIden As String, ByVal numDocumIden As String, ByVal tipoProcedencia As String, ByVal tasaDetraccion As String, ByVal tasaPercepcion As String, ByVal dobleTributacion As String, ByVal tipoPension As String)
        Me.New()
    End Sub
#End Region
#Region "Metodos"
    Public Sub Agregar(d As NAnexoComplementario_Con)
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo", "@apellidoPaterno", "@apellidoMaterno", "@nombre", "@formuSuspen", "@telefono", "@idTipoEntidad", "@fechaRegistro", "@usuario", "@direccion", "@provincia", "@departamento", "@pais", "@zonaPostal", "@idDocumIden", "@numDocumIden", "@tipoProcedencia", "@tasaDetraccion", "@tasaPercepcion", "@dobleTributacion", "@tipoPension"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdTipoAnexo, d.IdAnexo, d.ApellidoPaterno, d.ApellidoMaterno, d.Nombre, d.FormuSuspen, d.Telefono, d.IdTipoEntidad, d.FechaRegistro, d.Usuario, d.Direccion, d.Provincia, d.Departamento, d.Pais, d.ZonaPostal, d.IdDocumIden, d.NumDocumIden, d.TipoProcedencia, d.TasaDetraccion, d.TasaPercepcion, d.DobleTributacion, d.TipoPension}
        sql.EjecutarProcedure(Bd & ".dbo.Str_AnexoComplementario_I", parametros, valores, tipoParametro, 22)
    End Sub
    Public Sub Actualizar(d As NAnexoComplementario_Con)
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo", "@apellidoPaterno", "@apellidoMaterno", "@nombre", "@formuSuspen", "@telefono", "@idTipoEntidad", "@fechaRegistro", "@usuario", "@direccion", "@provincia", "@departamento", "@pais", "@zonaPostal", "@idDocumIden", "@numDocumIden", "@tipoProcedencia", "@tasaDetraccion", "@tasaPercepcion", "@dobleTributacion", "@tipoPension"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdTipoAnexo, d.IdAnexo, d.ApellidoPaterno, d.ApellidoMaterno, d.Nombre, d.FormuSuspen, d.Telefono, d.IdTipoEntidad, d.FechaRegistro, d.Usuario, d.Direccion, d.Provincia, d.Departamento, d.Pais, d.ZonaPostal, d.IdDocumIden, d.NumDocumIden, d.TipoProcedencia, d.TasaDetraccion, d.TasaPercepcion, d.DobleTributacion, d.TipoPension}
        sql.EjecutarProcedure(Bd & ".dbo.Str_AnexoComplementario_U", parametros, valores, tipoParametro, 22)
    End Sub
    Public Sub Eliminar(d As NAnexoComplementario_Con)
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdTipoAnexo, d.IdAnexo}
        sql.EjecutarProcedure(Bd & ".dbo.Str_AnexoComplementario_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_AnexoComplementario_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NAnexoComplementario_Con) As DataTable
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdTipoAnexo, d.IdAnexo}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_AnexoComplementario_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NAnexoComplementario_Con) As NAnexoComplementario_Con
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdTipoAnexo, d.IdAnexo}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_AnexoComplementario_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdTipoAnexo = IIf(dt.Rows(0).Item("idTipoAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoAnexo"))
            d.IdAnexo = IIf(dt.Rows(0).Item("idAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAnexo"))
            d.ApellidoPaterno = IIf(dt.Rows(0).Item("apellidoPaterno") Is DBNull.Value, Nothing, dt.Rows(0).Item("apellidoPaterno"))
            d.ApellidoMaterno = IIf(dt.Rows(0).Item("apellidoMaterno") Is DBNull.Value, Nothing, dt.Rows(0).Item("apellidoMaterno"))
            d.Nombre = IIf(dt.Rows(0).Item("nombre") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombre"))
            d.FormuSuspen = IIf(dt.Rows(0).Item("formuSuspen") Is DBNull.Value, Nothing, dt.Rows(0).Item("formuSuspen"))
            d.Telefono = IIf(dt.Rows(0).Item("telefono") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono"))
            d.IdTipoEntidad = IIf(dt.Rows(0).Item("idTipoEntidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoEntidad"))
            d.FechaRegistro = IIf(dt.Rows(0).Item("fechaRegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaRegistro"))
            d.Usuario = IIf(dt.Rows(0).Item("usuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuario"))
            d.Direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.Provincia = IIf(dt.Rows(0).Item("provincia") Is DBNull.Value, Nothing, dt.Rows(0).Item("provincia"))
            d.Departamento = IIf(dt.Rows(0).Item("departamento") Is DBNull.Value, Nothing, dt.Rows(0).Item("departamento"))
            d.Pais = IIf(dt.Rows(0).Item("pais") Is DBNull.Value, Nothing, dt.Rows(0).Item("pais"))
            d.ZonaPostal = IIf(dt.Rows(0).Item("zonaPostal") Is DBNull.Value, Nothing, dt.Rows(0).Item("zonaPostal"))
            d.IdDocumIden = IIf(dt.Rows(0).Item("idDocumIden") Is DBNull.Value, Nothing, dt.Rows(0).Item("idDocumIden"))
            d.NumDocumIden = IIf(dt.Rows(0).Item("numDocumIden") Is DBNull.Value, Nothing, dt.Rows(0).Item("numDocumIden"))
            d.TipoProcedencia = IIf(dt.Rows(0).Item("tipoProcedencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoProcedencia"))
            d.TasaDetraccion = IIf(dt.Rows(0).Item("tasaDetraccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("tasaDetraccion"))
            d.TasaPercepcion = IIf(dt.Rows(0).Item("tasaPercepcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("tasaPercepcion"))
            d.DobleTributacion = IIf(dt.Rows(0).Item("dobleTributacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("dobleTributacion"))
            d.TipoPension = IIf(dt.Rows(0).Item("tipoPension") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoPension"))
        Else
            d.ApellidoPaterno = Nothing
            d.ApellidoMaterno = Nothing
            d.Nombre = Nothing
            d.FormuSuspen = Nothing
            d.Telefono = Nothing
            d.IdTipoEntidad = Nothing
            d.FechaRegistro = Nothing
            d.Usuario = Nothing
            d.Direccion = Nothing
            d.Provincia = Nothing
            d.Departamento = Nothing
            d.Pais = Nothing
            d.ZonaPostal = Nothing
            d.IdDocumIden = Nothing
            d.NumDocumIden = Nothing
            d.TipoProcedencia = Nothing
            d.TasaDetraccion = Nothing
            d.TasaPercepcion = Nothing
            d.DobleTributacion = Nothing
            d.TipoPension = Nothing
        End If
        Return d
    End Function
#End Region
End Class
