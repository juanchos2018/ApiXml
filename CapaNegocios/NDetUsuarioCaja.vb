Imports CapaDatos
Public Class NDetUsuarioCaja
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _idDetUsuarioCaja As Integer
    Private _idCaja As String
    Private _idCuenta As String
    Private _idAnexo As String
    Private _idMoneda As String
    Private _tipoCambio As Decimal
    Private _descripcion As String
    Private _subDiarioIngreso As String
    Private _subDiarioEgreso As String
    Private _saldoAnterior As Decimal
    Private _saldoAnteriorMN As Decimal
    Private _saldoAnteriorUS As Decimal
    Private _estado As String
    Private _usuarioCrea As String
    Private _usuarioMod As String
    Private _idTipoOperacion As String
    Private _fechaCrea As System.DateTime
    Private _fechaMod As System.DateTime
    Private _fechaProceso As System.DateTime
    Private _fechaSaldoAnteror As System.DateTime

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

    Public Property TipoCambio As Decimal
        Get
            Return _tipoCambio
        End Get
        Set
            _tipoCambio = Value
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

    Public Property SaldoAnterior As Decimal
        Get
            Return _saldoAnterior
        End Get
        Set
            _saldoAnterior = Value
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

    Public Property IdTipoOperacion As String
        Get
            Return _idTipoOperacion
        End Get
        Set
            _idTipoOperacion = Value
        End Set
    End Property

    Public Property chkprimero As Boolean
#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idDetUsuarioCaja As Integer, ByVal idCaja As String, ByVal idCuenta As String, ByVal idAnexo As String, ByVal idMoneda As String, ByVal tipoCambio As Decimal, ByVal descripcion As String, ByVal subDiarioIngreso As String, ByVal subDiarioEgreso As String, ByVal fechaProceso As System.DateTime, ByVal saldoAnterior As Decimal, ByVal saldoAnteriorMN As Decimal, ByVal saldoAnteriorUS As Decimal, ByVal fechaSaldoAnteror As System.DateTime, ByVal estado As String, ByVal usuarioCrea As String, ByVal usuarioMod As String, ByVal fechaCrea As System.DateTime, ByVal fechaMod As System.DateTime, ByVal idTipoOperacion As String)
        Me.New()
    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NDetUsuarioCaja)
        Dim parametros() As Object = {"@idCaja", "@idCuenta", "@idAnexo", "@idMoneda", "@tipoCambio", "@descripcion", "@subDiarioIngreso", "@subDiarioEgreso", "@fechaProceso", "@saldoAnterior", "@saldoAnteriorMN", "@saldoAnteriorUS", "@fechaSaldoAnteror", "@estado", "@usuarioCrea", "@usuarioMod", "@fechaCrea", "@fechaMod", "@idTipoOperacion", "@chkprimero"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.IdCaja, d.IdCuenta, d.IdAnexo, d.IdMoneda, d.TipoCambio, d.Descripcion, d.SubDiarioIngreso, d.SubDiarioEgreso, d.FechaProceso, d.SaldoAnterior, d.SaldoAnteriorMN, d.SaldoAnteriorUS, d.FechaSaldoAnteror, d.Estado, d.UsuarioCrea, d.UsuarioMod, d.FechaCrea, d.FechaMod, d.IdTipoOperacion, d.chkprimero}
        sql.EjecutarProcedure("Str_Tbl_DetalleUsuarioCaja_I", parametros, valores, tipoParametro, 20)
    End Sub
    Public Sub Actualizar(d As NDetUsuarioCaja)
        Dim parametros() As Object = {"@idDetUsuarioCaja", "@idCaja", "@idCuenta", "@idAnexo", "@idMoneda", "@tipoCambio", "@descripcion", "@subDiarioIngreso", "@subDiarioEgreso", "@fechaProceso", "@saldoAnterior", "@saldoAnteriorMN", "@saldoAnteriorUS", "@fechaSaldoAnteror", "@estado", "@usuarioCrea", "@usuarioMod", "@fechaCrea", "@fechaMod", "@idTipoOperacion", "@chkprimero"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.IdDetUsuarioCaja, d.IdCaja, d.IdCuenta, d.IdAnexo, d.IdMoneda, d.TipoCambio, d.Descripcion, d.SubDiarioIngreso, d.SubDiarioEgreso, d.FechaProceso, d.SaldoAnterior, d.SaldoAnteriorMN, d.SaldoAnteriorUS, d.FechaSaldoAnteror, d.Estado, d.UsuarioCrea, d.UsuarioMod, d.FechaCrea, d.FechaMod, d.IdTipoOperacion, d.chkprimero}
        sql.EjecutarProcedure("Str_Tbl_DetalleUsuarioCaja_U", parametros, valores, tipoParametro, 21)
    End Sub
    Public Sub Eliminar(d As NDetUsuarioCaja)
        Dim parametros() As Object = {"@idDetUsuarioCaja"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.IdDetUsuarioCaja}
        sql.EjecutarProcedure("Str_Tbl_DetalleUsuarioCaja_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idDetUsuarioCaja"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_DetalleUsuarioCaja_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function DetalleCajaLista(IdCaja As String) As DataTable
        Dim parametros() As Object = {"@IdCaja", "@IdDetUsuarioCaja"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.Int}
        Dim valores() As Object = {IdCaja, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_DetalleUsuarioCaja_M", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function DetalleCajaLista(IdCaja As String, iddetusuariocaja As Integer) As DataTable
        Dim parametros() As Object = {"@IdCaja", "@IdDetUsuarioCaja"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.Int}
        Dim valores() As Object = {IdCaja, iddetusuariocaja}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_DetalleUsuarioCaja_M", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function


    Public Function Registro(d As NDetUsuarioCaja) As NDetUsuarioCaja
        Dim parametros() As Object = {"@idDetUsuarioCaja"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.IdDetUsuarioCaja}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_DetalleUsuarioCaja_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdDetUsuarioCaja = dt.Rows(0).Item("idDetUsuarioCaja")
            d.IdCaja = dt.Rows(0).Item("idCaja").ToString
            d.IdCuenta = dt.Rows(0).Item("idCuenta").ToString
            d.IdAnexo = dt.Rows(0).Item("idAnexo").ToString
            d.IdMoneda = dt.Rows(0).Item("idMoneda").ToString
            d.TipoCambio = dt.Rows(0).Item("tipoCambio")
            d.Descripcion = dt.Rows(0).Item("descripcion").ToString
            d.SubDiarioIngreso = dt.Rows(0).Item("subDiarioIngreso").ToString
            d.SubDiarioEgreso = dt.Rows(0).Item("subDiarioEgreso").ToString
            d.FechaProceso = dt.Rows(0).Item("fechaProceso")
            d.SaldoAnterior = dt.Rows(0).Item("saldoAnterior")
            d.SaldoAnteriorMN = dt.Rows(0).Item("saldoAnteriorMN")
            d.SaldoAnteriorUS = dt.Rows(0).Item("saldoAnteriorUS")
            d.FechaSaldoAnteror = IIf(dt.Rows(0).Item("fechaSaldoAnteror") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaSaldoAnteror"))
            d.Estado = dt.Rows(0).Item("estado").ToString
            d.UsuarioCrea = dt.Rows(0).Item("usuarioCrea").ToString
            d.UsuarioMod = dt.Rows(0).Item("usuarioMod").ToString
            d.FechaCrea = dt.Rows(0).Item("fechaCrea")
            d.FechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.IdTipoOperacion = dt.Rows(0).Item("idTipoOperacion").ToString
            d.chkprimero = IIf(dt.Rows(0).Item("chkprimero") Is DBNull.Value, Nothing, dt.Rows(0).Item("chkprimero"))
        Else
            d.IdDetUsuarioCaja = 0
            d.IdCaja = 0
            d.IdCuenta = 0
            d.IdAnexo = 0
            d.IdMoneda = 0
            d.TipoCambio = 0
            d.Descripcion = 0
            d.SubDiarioIngreso = 0
            d.SubDiarioEgreso = 0
            d.SaldoAnterior = 0
            d.SaldoAnteriorMN = 0
            d.SaldoAnteriorUS = 0
            d.Estado = 0
            d.UsuarioCrea = 0
            d.UsuarioMod = 0
            d.IdTipoOperacion = 0
        End If
        Return d
    End Function
#End Region


End Class
