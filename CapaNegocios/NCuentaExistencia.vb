Imports CapaDatos
Public Class NCuentaExistencia
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _idcuenta As String
    Private _idventa As String
    Private _iddevolucion As String
    Private _idanexo As String
    Private _idflete As String
    Private _iddscto As String
    Private _fechaCrea As System.DateTime
    Private _fechaMod As System.DateTime
    Private _idusuario As String
    Private _idpromocion As String
    Private _idcostoven As String
    Private _descripcioncuenta As String
    Private _idcompra As String
    Private _idconsumo As String
    Private _iddiferido As String
    Private _idcentrocosto As String
    Private _idexport As String
    Private _idCtaVariacion As String

#End Region

#Region "Properties"

    Public Property idcuenta As String
        Get
            Return _idcuenta
        End Get
        Set
            _idcuenta = Value
        End Set
    End Property

    Public Property idventa As String
        Get
            Return _idventa
        End Get
        Set
            _idventa = Value
        End Set
    End Property

    Public Property iddevolucion As String
        Get
            Return _iddevolucion
        End Get
        Set
            _iddevolucion = Value
        End Set
    End Property

    Public Property idanexo As String
        Get
            Return _idanexo
        End Get
        Set
            _idanexo = Value
        End Set
    End Property

    Public Property idflete As String
        Get
            Return _idflete
        End Get
        Set
            _idflete = Value
        End Set
    End Property

    Public Property iddscto As String
        Get
            Return _iddscto
        End Get
        Set
            _iddscto = Value
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

    Public Property idusuario As String
        Get
            Return _idusuario
        End Get
        Set
            _idusuario = Value
        End Set
    End Property

    Public Property idpromocion As String
        Get
            Return _idpromocion
        End Get
        Set
            _idpromocion = Value
        End Set
    End Property

    Public Property idcostoven As String
        Get
            Return _idcostoven
        End Get
        Set
            _idcostoven = Value
        End Set
    End Property

    Public Property descripcioncuenta As String
        Get
            Return _descripcioncuenta
        End Get
        Set
            _descripcioncuenta = Value
        End Set
    End Property

    Public Property idcompra As String
        Get
            Return _idcompra
        End Get
        Set
            _idcompra = Value
        End Set
    End Property

    Public Property idconsumo As String
        Get
            Return _idconsumo
        End Get
        Set
            _idconsumo = Value
        End Set
    End Property

    Public Property iddiferido As String
        Get
            Return _iddiferido
        End Get
        Set
            _iddiferido = Value
        End Set
    End Property

    Public Property idcentrocosto As String
        Get
            Return _idcentrocosto
        End Get
        Set
            _idcentrocosto = Value
        End Set
    End Property

    Public Property idexport As String
        Get
            Return _idexport
        End Get
        Set
            _idexport = Value
        End Set
    End Property

    Public Property IdCtaVariacion As String
        Get
            Return _idCtaVariacion
        End Get
        Set
            _idCtaVariacion = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idcuenta As String, ByVal idventa As String, ByVal iddevolucion As String, ByVal idanexo As String, ByVal idflete As String, ByVal iddscto As String, ByVal fechaCrea As System.DateTime, ByVal fechaMod As System.DateTime, ByVal idusuario As String, ByVal idpromocion As String, ByVal idcostoven As String, ByVal descripcioncuenta As String, ByVal idcompra As String, ByVal idconsumo As String, ByVal iddiferido As String, ByVal idcentrocosto As String, ByVal idexport As String, ByVal idCtaVariacion As String)
        Me.New()
    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NCuentaExistencia)
        Dim parametros() As Object = {"@idcuenta", "@idventa", "@iddevolucion", "@idanexo", "@idflete", "@iddscto", "@fechaCrea", "@fechaMod", "@idusuario", "@idpromocion", "@idcostoven", "@descripcioncuenta", "@idcompra", "@idconsumo", "@iddiferido", "@idcentrocosto", "@idexport", "@idCtaVariacion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idcuenta, d.idventa, d.iddevolucion, d.idanexo, d.idflete, d.iddscto, d.FechaCrea, d.FechaMod, d.idusuario, d.idpromocion, d.idcostoven, d.descripcioncuenta, d.idcompra, d.idconsumo, d.iddiferido, d.idcentrocosto, d.idexport, d.IdCtaVariacion}
        sql.EjecutarProcedure("Str_cuentaexistencia_I", parametros, valores, tipoParametro, 18)
    End Sub
    Public Sub Actualizar(d As NCuentaExistencia)
        Dim parametros() As Object = {"@idcuenta", "@idventa", "@iddevolucion", "@idanexo", "@idflete", "@iddscto", "@fechaCrea", "@fechaMod", "@idusuario", "@idpromocion", "@idcostoven", "@descripcioncuenta", "@idcompra", "@idconsumo", "@iddiferido", "@idcentrocosto", "@idexport", "@idCtaVariacion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idcuenta, d.idventa, d.iddevolucion, d.idanexo, d.idflete, d.iddscto, d.FechaCrea, d.FechaMod, d.idusuario, d.idpromocion, d.idcostoven, d.descripcioncuenta, d.idcompra, d.idconsumo, d.iddiferido, d.idcentrocosto, d.idexport, d.IdCtaVariacion}
        sql.EjecutarProcedure("Str_cuentaexistencia_U", parametros, valores, tipoParametro, 18)
    End Sub
    Public Sub Eliminar(d As NCuentaExistencia)
        Dim parametros() As Object = {"@idcuenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idcuenta}
        sql.EjecutarProcedure("Str_cuentaexistencia_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idcuenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_cuentaexistencia_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Existe(d As NCuentaExistencia) As Boolean
        Dim parametros() As Object = {"@idcuenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idcuenta}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_cuentaexistencia_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Registro(d As NCuentaExistencia) As NCuentaExistencia
        Dim parametros() As Object = {"@idcuenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idcuenta}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_cuentaexistencia_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idcuenta = IIf(dt.Rows(0).Item("idcuenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcuenta"))
            d.idventa = IIf(dt.Rows(0).Item("idventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idventa"))
            d.iddevolucion = IIf(dt.Rows(0).Item("iddevolucion") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddevolucion"))
            d.idanexo = IIf(dt.Rows(0).Item("idanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idanexo"))
            d.idflete = IIf(dt.Rows(0).Item("idflete") Is DBNull.Value, Nothing, dt.Rows(0).Item("idflete"))
            d.iddscto = IIf(dt.Rows(0).Item("iddscto") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddscto"))
            d.FechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.FechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.idusuario = IIf(dt.Rows(0).Item("idusuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuario"))
            d.idpromocion = IIf(dt.Rows(0).Item("idpromocion") Is DBNull.Value, Nothing, dt.Rows(0).Item("idpromocion"))
            d.idcostoven = IIf(dt.Rows(0).Item("idcostoven") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcostoven"))
            d.descripcioncuenta = IIf(dt.Rows(0).Item("descripcioncuenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcioncuenta"))
            d.idcompra = IIf(dt.Rows(0).Item("idcompra") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcompra"))
            d.idconsumo = IIf(dt.Rows(0).Item("idconsumo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idconsumo"))
            d.iddiferido = IIf(dt.Rows(0).Item("iddiferido") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddiferido"))
            d.idcentrocosto = IIf(dt.Rows(0).Item("idcentrocosto") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcentrocosto"))
            d.idexport = IIf(dt.Rows(0).Item("idexport") Is DBNull.Value, Nothing, dt.Rows(0).Item("idexport"))
            d.IdCtaVariacion = IIf(dt.Rows(0).Item("idCtaVariacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCtaVariacion"))
        Else
            d.idventa = Nothing
            d.iddevolucion = Nothing
            d.idanexo = Nothing
            d.idflete = Nothing
            d.iddscto = Nothing
            d.FechaCrea = Nothing
            d.FechaMod = Nothing
            d.idusuario = Nothing
            d.idpromocion = Nothing
            d.idcostoven = Nothing
            d.descripcioncuenta = Nothing
            d.idcompra = Nothing
            d.idconsumo = Nothing
            d.iddiferido = Nothing
            d.idcentrocosto = Nothing
            d.idexport = Nothing
            d.IdCtaVariacion = Nothing
        End If
        Return d
    End Function
#End Region
End Class
