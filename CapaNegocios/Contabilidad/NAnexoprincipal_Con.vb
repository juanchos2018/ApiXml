Imports CapaDatos
Public Class NAnexoprincipal_Con
    Dim sql As New ClsConexion
#Region "Declarations"
    Private _idTipoAnexo As String
    Private _idAnexo As String
    Private _descripcion As String
    Private _referencia As String
    Private _rUC As String
    Private _idMoneda As String
    Private _estado As String
    Private _fechaRegistro As System.DateTime
    Private _horaRegistro As String
    Private _avrete As String
    Private _aporre As Decimal
    Private _sw_Rus As Boolean
    Private _idPais As String
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

    Public Property Descripcion As String
        Get
            Return _descripcion
        End Get
        Set
            _descripcion = Value
        End Set
    End Property

    Public Property Referencia As String
        Get
            Return _referencia
        End Get
        Set
            _referencia = Value
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

    Public Property IdMoneda As String
        Get
            Return _idMoneda
        End Get
        Set
            _idMoneda = Value
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

    Public Property FechaRegistro As System.DateTime
        Get
            Return _fechaRegistro
        End Get
        Set
            _fechaRegistro = Value
        End Set
    End Property

    Public Property HoraRegistro As String
        Get
            Return _horaRegistro
        End Get
        Set
            _horaRegistro = Value
        End Set
    End Property

    Public Property avrete As String
        Get
            Return _avrete
        End Get
        Set
            _avrete = Value
        End Set
    End Property

    Public Property aporre As Decimal
        Get
            Return _aporre
        End Get
        Set
            _aporre = Value
        End Set
    End Property

    Public Property Sw_Rus As Boolean
        Get
            Return _sw_Rus
        End Get
        Set
            _sw_Rus = Value
        End Set
    End Property

    Public Property IdPais As String
        Get
            Return _idPais
        End Get
        Set
            _idPais = Value
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

    Public Sub New(ByVal idTipoAnexo As String, ByVal idAnexo As String, ByVal descripcion As String, ByVal referencia As String, ByVal rUC As String, ByVal idMoneda As String, ByVal estado As String, ByVal fechaRegistro As System.DateTime, ByVal horaRegistro As String, ByVal avrete As String, ByVal aporre As Decimal, ByVal sw_Rus As Boolean, ByVal idPais As String)
        Me.New()
    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NAnexoprincipal_Con)
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo", "@descripcion", "@referencia", "@rUC", "@idMoneda", "@estado", "@fechaRegistro", "@horaRegistro", "@avrete", "@aporre", "@sw_Rus", "@idPais"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Bit, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdTipoAnexo, d.IdAnexo, d.Descripcion, d.Referencia, d.RUC, d.IdMoneda, d.Estado, d.FechaRegistro, d.HoraRegistro, d.avrete, d.aporre, d.Sw_Rus, d.IdPais}
        sql.EjecutarProcedure(Bd & ".dbo.Str_AnexoPrincipal_I", parametros, valores, tipoParametro, 13)
    End Sub
    Public Sub Actualizar(d As NAnexoprincipal_Con)
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo", "@descripcion", "@referencia", "@rUC", "@idMoneda", "@estado", "@fechaRegistro", "@horaRegistro", "@avrete", "@aporre", "@sw_Rus", "@idPais"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Bit, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdTipoAnexo, d.IdAnexo, d.Descripcion, d.Referencia, d.RUC, d.IdMoneda, d.Estado, d.FechaRegistro, d.HoraRegistro, d.avrete, d.aporre, d.Sw_Rus, d.IdPais}
        sql.EjecutarProcedure(Bd & ".dbo.Str_AnexoPrincipal_U", parametros, valores, tipoParametro, 13)
    End Sub
    Public Sub Eliminar(d As NAnexoprincipal_Con)
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdTipoAnexo, d.IdAnexo}
        sql.EjecutarProcedure(Bd & ".dbo.Str_AnexoPrincipal_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_AnexoPrincipal_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NAnexoprincipal_Con) As DataTable
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdTipoAnexo, d.IdAnexo}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_AnexoPrincipal_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function

    Public Function Registro(d As NAnexoprincipal_Con) As NAnexoprincipal_Con
        Dim parametros() As Object = {"@idTipoAnexo", "@idAnexo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdTipoAnexo, d.IdAnexo}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_AnexoPrincipal_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdTipoAnexo = IIf(dt.Rows(0).Item("idTipoAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoAnexo"))
            d.IdAnexo = IIf(dt.Rows(0).Item("idAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAnexo"))
            d.Descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.Referencia = IIf(dt.Rows(0).Item("referencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("referencia"))
            d.RUC = IIf(dt.Rows(0).Item("rUC") Is DBNull.Value, Nothing, dt.Rows(0).Item("rUC"))
            d.IdMoneda = IIf(dt.Rows(0).Item("idMoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMoneda"))
            d.Estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.FechaRegistro = IIf(dt.Rows(0).Item("fechaRegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaRegistro"))
            d.HoraRegistro = IIf(dt.Rows(0).Item("horaRegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("horaRegistro"))
            d.avrete = IIf(dt.Rows(0).Item("avrete") Is DBNull.Value, Nothing, dt.Rows(0).Item("avrete"))
            d.aporre = IIf(dt.Rows(0).Item("aporre") Is DBNull.Value, Nothing, dt.Rows(0).Item("aporre"))
            d.Sw_Rus = IIf(dt.Rows(0).Item("sw_Rus") Is DBNull.Value, Nothing, dt.Rows(0).Item("sw_Rus"))
            d.IdPais = IIf(dt.Rows(0).Item("idPais") Is DBNull.Value, Nothing, dt.Rows(0).Item("idPais"))
        Else
            d.Descripcion = Nothing
            d.Referencia = Nothing
            d.RUC = Nothing
            d.IdMoneda = Nothing
            d.Estado = Nothing
            d.FechaRegistro = Nothing
            d.HoraRegistro = Nothing
            d.avrete = Nothing
            d.aporre = Nothing
            d.Sw_Rus = Nothing
            d.IdPais = Nothing
        End If
        Return d
    End Function
#End Region
End Class
