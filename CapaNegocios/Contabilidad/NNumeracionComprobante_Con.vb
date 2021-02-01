Imports CapaDatos
Public Class NNumeracionComprobante_Con
    Dim sql As New ClsConexion
#Region "Declarations"
    Private _idSubdiario As String
    Private _anio As String
    Private _mes As String
    Private _numero As Decimal
    Private _fechaCrea As System.DateTime
    Private _fechaMod As System.DateTime
    Private _bd As String

#End Region

#Region "Properties"

    Public Property idSubdiario As String
        Get
            Return _idSubdiario
        End Get
        Set
            _idSubdiario = Value
        End Set
    End Property

    Public Property Anio As String
        Get
            Return _anio
        End Get
        Set
            _anio = Value
        End Set
    End Property

    Public Property Mes As String
        Get
            Return _mes
        End Get
        Set
            _mes = Value
        End Set
    End Property

    Public Property Numero As Decimal
        Get
            Return _numero
        End Get
        Set
            _numero = Value
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

    Public Property FechaMod As System.DateTime
        Get
            Return _fechaMod
        End Get
        Set
            _fechaMod = Value
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

    Public Sub New(ByVal idSubdiario As String, ByVal anio As String, ByVal mes As String, ByVal numero As Decimal, ByVal fechaCrea As System.DateTime, ByVal fechaMod As System.DateTime)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NNumeracionComprobante_Con)
        Dim parametros() As Object = {"@idSubdiario", "@anio", "@mes", "@numero", "@fechaCrea", "@fechaMod"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.DateTime}
        Dim valores() As Object = {d.idSubdiario, d.Anio, d.Mes, d.Numero, d.FechaCrea, d.FechaMod}
        sql.EjecutarProcedure(Bd & ".dbo.Str_NumeracionComprobante_I", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Actualizar(d As NNumeracionComprobante_Con)
        Dim parametros() As Object = {"@idSubdiario", "@anio", "@mes", "@numero", "@fechaCrea", "@fechaMod"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.DateTime}
        Dim valores() As Object = {d.idSubdiario, d.Anio, d.Mes, d.Numero, d.FechaCrea, d.FechaMod}
        sql.EjecutarProcedure(Bd & ".dbo.Str_NumeracionComprobante_U", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Eliminar(d As NNumeracionComprobante_Con)
        Dim parametros() As Object = {"@idSubdiario", "@anio", "@mes"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idSubdiario, d.Anio, d.Mes}
        sql.EjecutarProcedure(Bd & ".dbo.Str_NumeracionComprobante_D", parametros, valores, tipoParametro, 3)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idSubdiario", "@anio", "@mes"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_NumeracionComprobante_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NNumeracionComprobante_Con) As DataTable
        Dim parametros() As Object = {"@idSubdiario", "@anio", "@mes"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idSubdiario, d.Anio, d.Mes}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_NumeracionComprobante_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NNumeracionComprobante_Con) As NNumeracionComprobante_Con
        Dim parametros() As Object = {"@idSubdiario", "@anio", "@mes"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idSubdiario, d.Anio, d.Mes}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_NumeracionComprobante_S", parametros, valores, tipoParametro, 3).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idSubdiario = IIf(dt.Rows(0).Item("idSubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idSubdiario"))
            d.Anio = IIf(dt.Rows(0).Item("anio") Is DBNull.Value, Nothing, dt.Rows(0).Item("anio"))
            d.Mes = IIf(dt.Rows(0).Item("mes") Is DBNull.Value, Nothing, dt.Rows(0).Item("mes"))
            d.Numero = IIf(dt.Rows(0).Item("numero") Is DBNull.Value, Nothing, dt.Rows(0).Item("numero"))
            d.FechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.FechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
        Else
            d.Numero = Nothing
            d.FechaCrea = Nothing
            d.FechaMod = Nothing
        End If
        Return d
    End Function
#End Region
End Class
