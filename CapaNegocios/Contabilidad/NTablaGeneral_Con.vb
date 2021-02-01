Imports CapaDatos

Public Class NTablaGeneral_Con
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _idGeneral As String
    Private _idCodigo As String
    Private _descripcion As String
    Private _fechaRegistro As System.DateTime
    Private _horaRegistro As String
    Private _checkLibros As Boolean
    Private _alias As String
    Private _bd As String

#End Region

#Region "Properties"

    Public Property IdGeneral As String
        Get
            Return _idGeneral
        End Get
        Set
            _idGeneral = Value
        End Set
    End Property

    Public Property IdCodigo As String
        Get
            Return _idCodigo
        End Get
        Set
            _idCodigo = Value
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

    Public Property FechaRegistro As System.DateTime
        Get
            Return _fechaRegistro
        End Get
        Set
            _fechaRegistro = Value
        End Set
    End Property

    Public Property horaRegistro As String
        Get
            Return _horaRegistro
        End Get
        Set
            _horaRegistro = Value
        End Set
    End Property

    Public Property CheckLibros As Boolean
        Get
            Return _checkLibros
        End Get
        Set
            _checkLibros = Value
        End Set
    End Property

    Public Property [alias] As String
        Get
            Return _alias
        End Get
        Set
            _alias = Value
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

    Public Sub New(ByVal idGeneral As String, ByVal idCodigo As String, ByVal descripcion As String, ByVal fechaRegistro As System.DateTime, ByVal horaRegistro As String, ByVal checkLibros As Boolean, ByVal [alias] As String)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NTablaGeneral_Con)
        Dim parametros() As Object = {"@idGeneral", "@idCodigo", "@descripcion", "@fechaRegistro", "@horaRegistro", "@checkLibros", "@alias"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdGeneral, d.IdCodigo, d.Descripcion, d.FechaRegistro, d.horaRegistro, d.CheckLibros, d.alias}
        sql.EjecutarProcedure(Bd & ".dbo.Str_TablaGeneral_I", parametros, valores, tipoParametro, 7)
    End Sub
    Public Sub Actualizar(d As NTablaGeneral_Con)
        Dim parametros() As Object = {"@idGeneral", "@idCodigo", "@descripcion", "@fechaRegistro", "@horaRegistro", "@checkLibros", "@alias"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdGeneral, d.IdCodigo, d.Descripcion, d.FechaRegistro, d.horaRegistro, d.CheckLibros, d.alias}
        sql.EjecutarProcedure(Bd & ".dbo.Str_TablaGeneral_U", parametros, valores, tipoParametro, 7)
    End Sub
    Public Sub Eliminar(d As NTablaGeneral_Con)
        Dim parametros() As Object = {"@idGeneral", "@idCodigo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdGeneral, d.IdCodigo}
        sql.EjecutarProcedure(Bd & ".dbo.Str_TablaGeneral_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idGeneral", "@idCodigo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_TablaGeneral_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTablaGeneral_Con) As DataTable
        Dim parametros() As Object = {"@idGeneral", "@idCodigo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdGeneral, d.IdCodigo}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_TablaGeneral_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function areaFlujo(d As NTablaGeneral_Con) As DataTable
        Return sql.EjecutarConsulta("d", "select IdArea,Area from " & Bd & "..VArea").Tables(0)
    End Function
    Public Function MedioPago(d As NTablaGeneral_Con) As DataTable
        Return sql.EjecutarConsulta("d", "SELECT IdCodigo,IdCodigo+' ' +Descripcion as Descripcion from " & Bd & "..TablaGeneral where idgeneral='S1'").Tables(0)
    End Function
    Public Function Registro(d As NTablaGeneral_Con) As NTablaGeneral_Con
        Dim parametros() As Object = {"@idGeneral", "@idCodigo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdGeneral, d.IdCodigo}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_TablaGeneral_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdGeneral = IIf(dt.Rows(0).Item("idGeneral") Is DBNull.Value, Nothing, dt.Rows(0).Item("idGeneral"))
            d.IdCodigo = IIf(dt.Rows(0).Item("idCodigo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCodigo"))
            d.Descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.FechaRegistro = IIf(dt.Rows(0).Item("fechaRegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaRegistro"))
            d.horaRegistro = IIf(dt.Rows(0).Item("horaRegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("horaRegistro"))
            d.CheckLibros = IIf(dt.Rows(0).Item("checkLibros") Is DBNull.Value, Nothing, dt.Rows(0).Item("checkLibros"))
            d.alias = IIf(dt.Rows(0).Item("alias") Is DBNull.Value, Nothing, dt.Rows(0).Item("alias"))
        Else
            d.Descripcion = Nothing
            d.FechaRegistro = Nothing
            d.horaRegistro = Nothing
            d.CheckLibros = Nothing
            d.alias = Nothing
        End If
        Return d
    End Function
#End Region

End Class
