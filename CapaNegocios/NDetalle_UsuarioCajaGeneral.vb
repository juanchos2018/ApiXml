Imports CapaDatos
Public Class NDetalle_UsuarioCajaGeneral
    Dim sql As New ClsConexion
#Region "Declarations"
    Private _idDetUsuarioCaja As Integer
    Private _idAgencia As String
    Private _idCaja As String
    Private _idCuenta As String
    Private _idAnexo As String
    Private _idMoneda As String
    Private _descripcion As String
    Private _subDiarioIngreso As String
    Private _subDiarioEgreso As String
    Private _fechaProceso As System.DateTime
    Private _saldoAnteriorMN As Decimal
    Private _saldoAnteriorUS As Decimal
    Private _fechaSaldoAnteror As System.DateTime
    Private _ingresoDiarioMN As Decimal
    Private _ingresoDiarioUS As Decimal
    Private _egresoDiarioMN As Decimal
    Private _egresoDiarioUS As Decimal
    Private _saldoFinalMN As Decimal
    Private _saldoFinalUS As Decimal
    Private _estado As String
    Private _usuarioCrea As String
    Private _usuarioMod As String
    Private _fechaCrea As System.DateTime
    Private _fechaMod As System.DateTime
    Private _usuarioCaja As String
    Private _cJ_NSANCMN As Decimal
    Private _cJ_NSANCUS As Decimal
    Private _cJ_NSACCMN As Decimal
    Private _cJ_NSACCUS As Decimal
    Private _idTipoOperacion As String

#End Region

#Region "Properties"

    Public Property IdDetUsuarioCaja As Integer
        Get
            Return _idDetUsuarioCaja
        End Get
        Set
            _idDetUsuarioCaja = Value
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

    Public Property IdCaja As String
        Get
            Return _idCaja
        End Get
        Set
            _idCaja = Value
        End Set
    End Property

    Public Property IdCuenta As String
        Get
            Return _idCuenta
        End Get
        Set
            _idCuenta = Value
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

    Public Property IdMoneda As String
        Get
            Return _idMoneda
        End Get
        Set
            _idMoneda = Value
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

    Public Property SubDiarioIngreso As String
        Get
            Return _subDiarioIngreso
        End Get
        Set
            _subDiarioIngreso = Value
        End Set
    End Property

    Public Property SubDiarioEgreso As String
        Get
            Return _subDiarioEgreso
        End Get
        Set
            _subDiarioEgreso = Value
        End Set
    End Property

    Public Property FechaProceso As System.DateTime
        Get
            Return _fechaProceso
        End Get
        Set
            _fechaProceso = Value
        End Set
    End Property

    Public Property SaldoAnteriorMN As Decimal
        Get
            Return _saldoAnteriorMN
        End Get
        Set
            _saldoAnteriorMN = Value
        End Set
    End Property

    Public Property SaldoAnteriorUS As Decimal
        Get
            Return _saldoAnteriorUS
        End Get
        Set
            _saldoAnteriorUS = Value
        End Set
    End Property

    Public Property FechaSaldoAnteror As System.DateTime
        Get
            Return _fechaSaldoAnteror
        End Get
        Set
            _fechaSaldoAnteror = Value
        End Set
    End Property

    Public Property IngresoDiarioMN As Decimal
        Get
            Return _ingresoDiarioMN
        End Get
        Set
            _ingresoDiarioMN = Value
        End Set
    End Property

    Public Property IngresoDiarioUS As Decimal
        Get
            Return _ingresoDiarioUS
        End Get
        Set
            _ingresoDiarioUS = Value
        End Set
    End Property

    Public Property EgresoDiarioMN As Decimal
        Get
            Return _egresoDiarioMN
        End Get
        Set
            _egresoDiarioMN = Value
        End Set
    End Property

    Public Property EgresoDiarioUS As Decimal
        Get
            Return _egresoDiarioUS
        End Get
        Set
            _egresoDiarioUS = Value
        End Set
    End Property

    Public Property SaldoFinalMN As Decimal
        Get
            Return _saldoFinalMN
        End Get
        Set
            _saldoFinalMN = Value
        End Set
    End Property

    Public Property SaldoFinalUS As Decimal
        Get
            Return _saldoFinalUS
        End Get
        Set
            _saldoFinalUS = Value
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

    Public Property UsuarioCrea As String
        Get
            Return _usuarioCrea
        End Get
        Set
            _usuarioCrea = Value
        End Set
    End Property

    Public Property UsuarioMod As String
        Get
            Return _usuarioMod
        End Get
        Set
            _usuarioMod = Value
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

    Public Property UsuarioCaja As String
        Get
            Return _usuarioCaja
        End Get
        Set
            _usuarioCaja = Value
        End Set
    End Property

    Public Property CJ_NSANCMN As Decimal
        Get
            Return _cJ_NSANCMN
        End Get
        Set
            _cJ_NSANCMN = Value
        End Set
    End Property

    Public Property CJ_NSANCUS As Decimal
        Get
            Return _cJ_NSANCUS
        End Get
        Set
            _cJ_NSANCUS = Value
        End Set
    End Property

    Public Property CJ_NSACCMN As Decimal
        Get
            Return _cJ_NSACCMN
        End Get
        Set
            _cJ_NSACCMN = Value
        End Set
    End Property

    Public Property CJ_NSACCUS As Decimal
        Get
            Return _cJ_NSACCUS
        End Get
        Set
            _cJ_NSACCUS = Value
        End Set
    End Property

    Public Property IdTipoOperacion As String
        Get
            Return _idTipoOperacion
        End Get
        Set
            _idTipoOperacion = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idDetUsuarioCaja As Integer, ByVal idAgencia As String, ByVal idCaja As String, ByVal idCuenta As String, ByVal idAnexo As String, ByVal idMoneda As String, ByVal descripcion As String, ByVal subDiarioIngreso As String, ByVal subDiarioEgreso As String, ByVal fechaProceso As System.DateTime, ByVal saldoAnteriorMN As Decimal, ByVal saldoAnteriorUS As Decimal, ByVal fechaSaldoAnteror As System.DateTime, ByVal ingresoDiarioMN As Decimal, ByVal ingresoDiarioUS As Decimal, ByVal egresoDiarioMN As Decimal, ByVal egresoDiarioUS As Decimal, ByVal saldoFinalMN As Decimal, ByVal saldoFinalUS As Decimal, ByVal estado As String, ByVal usuarioCrea As String, ByVal usuarioMod As String, ByVal fechaCrea As System.DateTime, ByVal fechaMod As System.DateTime, ByVal usuarioCaja As String, ByVal cJ_NSANCMN As Decimal, ByVal cJ_NSANCUS As Decimal, ByVal cJ_NSACCMN As Decimal, ByVal cJ_NSACCUS As Decimal, ByVal idTipoOperacion As String)
        Me.New()
    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NDetalle_UsuarioCajaGeneral)
        Dim parametros() As Object = {"@idAgencia", "@idCaja", "@idCuenta", "@idAnexo", "@idMoneda", "@descripcion", "@subDiarioIngreso", "@subDiarioEgreso", "@fechaProceso", "@saldoAnteriorMN", "@saldoAnteriorUS", "@fechaSaldoAnteror", "@ingresoDiarioMN", "@ingresoDiarioUS", "@egresoDiarioMN", "@egresoDiarioUS", "@saldoFinalMN", "@saldoFinalUS", "@estado", "@usuarioCrea", "@usuarioMod", "@fechaCrea", "@fechaMod", "@usuarioCaja", "@cJ_NSANCMN", "@cJ_NSANCUS", "@cJ_NSACCMN", "@cJ_NSACCUS", "@idTipoOperacion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdDetUsuarioCaja, d.IdAgencia, d.IdCaja, d.IdCuenta, d.IdAnexo, d.IdMoneda, d.Descripcion, d.SubDiarioIngreso, d.SubDiarioEgreso, d.FechaProceso, d.SaldoAnteriorMN, d.SaldoAnteriorUS, d.FechaSaldoAnteror, d.IngresoDiarioMN, d.IngresoDiarioUS, d.EgresoDiarioMN, d.EgresoDiarioUS, d.SaldoFinalMN, d.SaldoFinalUS, d.Estado, d.UsuarioCrea, d.UsuarioMod, d.FechaCrea, d.FechaMod, d.UsuarioCaja, d.CJ_NSANCMN, d.CJ_NSANCUS, d.CJ_NSACCMN, d.CJ_NSACCUS, d.IdTipoOperacion}
        sql.EjecutarProcedure("Str_AsignacionCaja_I", parametros, valores, tipoParametro, 29)
    End Sub
    Public Sub Actualizar(d As NDetalle_UsuarioCajaGeneral)
        Dim parametros() As Object = {"@idDetUsuarioCaja", "@idAgencia", "@idCaja", "@idCuenta", "@idAnexo", "@idMoneda", "@descripcion", "@subDiarioIngreso", "@subDiarioEgreso", "@fechaProceso", "@saldoAnteriorMN", "@saldoAnteriorUS", "@fechaSaldoAnteror", "@ingresoDiarioMN", "@ingresoDiarioUS", "@egresoDiarioMN", "@egresoDiarioUS", "@saldoFinalMN", "@saldoFinalUS", "@estado", "@usuarioCrea", "@usuarioMod", "@fechaCrea", "@fechaMod", "@usuarioCaja", "@cJ_NSANCMN", "@cJ_NSANCUS", "@cJ_NSACCMN", "@cJ_NSACCUS", "@idTipoOperacion"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdDetUsuarioCaja, d.IdAgencia, d.IdCaja, d.IdCuenta, d.IdAnexo, d.IdMoneda, d.Descripcion, d.SubDiarioIngreso, d.SubDiarioEgreso, d.FechaProceso, d.SaldoAnteriorMN, d.SaldoAnteriorUS, d.FechaSaldoAnteror, d.IngresoDiarioMN, d.IngresoDiarioUS, d.EgresoDiarioMN, d.EgresoDiarioUS, d.SaldoFinalMN, d.SaldoFinalUS, d.Estado, d.UsuarioCrea, d.UsuarioMod, d.FechaCrea, d.FechaMod, d.UsuarioCaja, d.CJ_NSANCMN, d.CJ_NSANCUS, d.CJ_NSACCMN, d.CJ_NSACCUS, d.IdTipoOperacion}
        sql.EjecutarProcedure("Str_AsignacionCaja_U", parametros, valores, tipoParametro, 30)
    End Sub
    Public Sub Eliminar(d As NDetalle_UsuarioCajaGeneral)
        Dim parametros() As Object = {"@idDetUsuarioCaja"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.IdDetUsuarioCaja}
        sql.EjecutarProcedure("Str_AsignacionCaja_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idDetUsuarioCaja"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_AsignacionCaja_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDetalle_UsuarioCajaGeneral) As NDetalle_UsuarioCajaGeneral
        Dim parametros() As Object = {"@idDetUsuarioCaja"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.IdDetUsuarioCaja}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_AsignacionCaja_S", parametros, valores, tipoParametro, 30).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdDetUsuarioCaja = dt.Rows(0).Item("idDetUsuarioCaja")
            d.IdAgencia = dt.Rows(0).Item("idAgencia")
            d.IdCaja = dt.Rows(0).Item("idCaja")
            d.IdCuenta = dt.Rows(0).Item("idCuenta")
            d.IdAnexo = dt.Rows(0).Item("idAnexo")
            d.IdMoneda = dt.Rows(0).Item("idMoneda")
            d.Descripcion = dt.Rows(0).Item("descripcion")
            d.SubDiarioIngreso = dt.Rows(0).Item("subDiarioIngreso")
            d.SubDiarioEgreso = dt.Rows(0).Item("subDiarioEgreso")
            d.FechaProceso = dt.Rows(0).Item("fechaProceso")
            d.SaldoAnteriorMN = dt.Rows(0).Item("saldoAnteriorMN")
            d.SaldoAnteriorUS = dt.Rows(0).Item("saldoAnteriorUS")
            d.FechaSaldoAnteror = dt.Rows(0).Item("fechaSaldoAnteror")
            d.IngresoDiarioMN = dt.Rows(0).Item("ingresoDiarioMN")
            d.IngresoDiarioUS = dt.Rows(0).Item("ingresoDiarioUS")
            d.EgresoDiarioMN = dt.Rows(0).Item("egresoDiarioMN")
            d.EgresoDiarioUS = dt.Rows(0).Item("egresoDiarioUS")
            d.SaldoFinalMN = dt.Rows(0).Item("saldoFinalMN")
            d.SaldoFinalUS = dt.Rows(0).Item("saldoFinalUS")
            d.Estado = dt.Rows(0).Item("estado")
            d.UsuarioCrea = dt.Rows(0).Item("usuarioCrea")
            d.UsuarioMod = dt.Rows(0).Item("usuarioMod")
            d.FechaCrea = dt.Rows(0).Item("fechaCrea")
            d.FechaMod = dt.Rows(0).Item("fechaMod")
            d.UsuarioCaja = dt.Rows(0).Item("usuarioCaja")
            d.CJ_NSANCMN = dt.Rows(0).Item("cJ_NSANCMN")
            d.CJ_NSANCUS = dt.Rows(0).Item("cJ_NSANCUS")
            d.CJ_NSACCMN = dt.Rows(0).Item("cJ_NSACCMN")
            d.CJ_NSACCUS = dt.Rows(0).Item("cJ_NSACCUS")
            d.IdTipoOperacion = dt.Rows(0).Item("idTipoOperacion")
        Else
            d.IdDetUsuarioCaja = 0
            d.IdAgencia = 0
            d.IdCaja = 0
            d.IdCuenta = 0
            d.IdAnexo = 0
            d.IdMoneda = 0
            d.Descripcion = 0
            d.SubDiarioIngreso = 0
            d.SubDiarioEgreso = 0
            d.SaldoAnteriorMN = 0
            d.SaldoAnteriorUS = 0
            d.IngresoDiarioMN = 0
            d.IngresoDiarioUS = 0
            d.EgresoDiarioMN = 0
            d.EgresoDiarioUS = 0
            d.SaldoFinalMN = 0
            d.SaldoFinalUS = 0
            d.Estado = 0
            d.UsuarioCrea = 0
            d.UsuarioMod = 0
            d.UsuarioCaja = 0
            d.CJ_NSANCMN = 0
            d.CJ_NSANCUS = 0
            d.CJ_NSACCMN = 0
            d.CJ_NSACCUS = 0
            d.IdTipoOperacion = 0
        End If
        Return d
    End Function
#End Region

End Class
