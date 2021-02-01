Imports CapaDatos
Public Class NPlanCuenta_Con
    Dim sql As New ClsConexion
#Region "Declarations"
    Private _idcuenta As String
    Private _descripcion As String
    Private _idTipoAnexo As String
    Private _idNivelSaldo As String
    Private _documentoRef As String
    Private _pldoc As Decimal
    Private _fechaVencimiento As String
    Private _idMonedaReferencial As String
    Private _idCuentaCargo As String
    Private _idCuentaAbono As String
    Private _centroCosto As String
    Private _idRegistroCuenta As String
    Private _idTipoCuenta As String
    Private _cuentaAjusteACM As String
    Private _area As String
    Private _conciliacionBanco As String
    Private _idBalanceGeneral As String
    Private _idGastPerdfuncion As String
    Private _plingyp As String
    Private _idGastPerdNaturale As String
    Private _idCentroCosto As String
    Private _estado As String
    Private _fechaRegistro As System.DateTime
    Private _horaRegistro As String
    Private _usuario As String
    Private _idAlFormgreGasto As String
    Private _idAlFormCosto As String
    Private _idAlFormBalanGen As String
    Private _idAlFormGasPerFun As String
    Private _idAlFormGasPerNat As String
    Private _pvactfij As String
    Private _pvglodet As String
    Private _idTipoAnexoRef As String
    Private _documentoref2 As String
    Private _tasa As String
    Private _medioPago As Boolean
    Private _bd As String
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

    Public Property Descripcion As String
        Get
            Return _descripcion
        End Get
        Set
            _descripcion = Value
        End Set
    End Property

    Public Property IdTipoAnexo As String
        Get
            Return _idTipoAnexo
        End Get
        Set
            _idTipoAnexo = Value
        End Set
    End Property

    Public Property IdNivelSaldo As String
        Get
            Return _idNivelSaldo
        End Get
        Set
            _idNivelSaldo = Value
        End Set
    End Property

    Public Property DocumentoRef As String
        Get
            Return _documentoRef
        End Get
        Set
            _documentoRef = Value
        End Set
    End Property

    Public Property pldoc As Decimal
        Get
            Return _pldoc
        End Get
        Set
            _pldoc = Value
        End Set
    End Property

    Public Property FechaVencimiento As String
        Get
            Return _fechaVencimiento
        End Get
        Set
            _fechaVencimiento = Value
        End Set
    End Property

    Public Property idMonedaReferencial As String
        Get
            Return _idMonedaReferencial
        End Get
        Set
            _idMonedaReferencial = Value
        End Set
    End Property

    Public Property IdCuentaCargo As String
        Get
            Return _idCuentaCargo
        End Get
        Set
            _idCuentaCargo = Value
        End Set
    End Property

    Public Property IdCuentaAbono As String
        Get
            Return _idCuentaAbono
        End Get
        Set
            _idCuentaAbono = Value
        End Set
    End Property

    Public Property CentroCosto As String
        Get
            Return _centroCosto
        End Get
        Set
            _centroCosto = Value
        End Set
    End Property

    Public Property idRegistroCuenta As String
        Get
            Return _idRegistroCuenta
        End Get
        Set
            _idRegistroCuenta = Value
        End Set
    End Property

    Public Property IdTipoCuenta As String
        Get
            Return _idTipoCuenta
        End Get
        Set
            _idTipoCuenta = Value
        End Set
    End Property

    Public Property CuentaAjusteACM As String
        Get
            Return _cuentaAjusteACM
        End Get
        Set
            _cuentaAjusteACM = Value
        End Set
    End Property

    Public Property Area As String
        Get
            Return _area
        End Get
        Set
            _area = Value
        End Set
    End Property

    Public Property conciliacionBanco As String
        Get
            Return _conciliacionBanco
        End Get
        Set
            _conciliacionBanco = Value
        End Set
    End Property

    Public Property IdBalanceGeneral As String
        Get
            Return _idBalanceGeneral
        End Get
        Set
            _idBalanceGeneral = Value
        End Set
    End Property

    Public Property IdGastPerdfuncion As String
        Get
            Return _idGastPerdfuncion
        End Get
        Set
            _idGastPerdfuncion = Value
        End Set
    End Property

    Public Property plingyp As String
        Get
            Return _plingyp
        End Get
        Set
            _plingyp = Value
        End Set
    End Property

    Public Property IdGastPerdNaturale As String
        Get
            Return _idGastPerdNaturale
        End Get
        Set
            _idGastPerdNaturale = Value
        End Set
    End Property

    Public Property IdCentroCosto As String
        Get
            Return _idCentroCosto
        End Get
        Set
            _idCentroCosto = Value
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

    Public Property Usuario As String
        Get
            Return _usuario
        End Get
        Set
            _usuario = Value
        End Set
    End Property

    Public Property IdAlFormgreGasto As String
        Get
            Return _idAlFormgreGasto
        End Get
        Set
            _idAlFormgreGasto = Value
        End Set
    End Property

    Public Property IdAlFormCosto As String
        Get
            Return _idAlFormCosto
        End Get
        Set
            _idAlFormCosto = Value
        End Set
    End Property

    Public Property IdAlFormBalanGen As String
        Get
            Return _idAlFormBalanGen
        End Get
        Set
            _idAlFormBalanGen = Value
        End Set
    End Property

    Public Property IdAlFormGasPerFun As String
        Get
            Return _idAlFormGasPerFun
        End Get
        Set
            _idAlFormGasPerFun = Value
        End Set
    End Property

    Public Property IdAlFormGasPerNat As String
        Get
            Return _idAlFormGasPerNat
        End Get
        Set
            _idAlFormGasPerNat = Value
        End Set
    End Property

    Public Property pvactfij As String
        Get
            Return _pvactfij
        End Get
        Set
            _pvactfij = Value
        End Set
    End Property

    Public Property pvglodet As String
        Get
            Return _pvglodet
        End Get
        Set
            _pvglodet = Value
        End Set
    End Property

    Public Property IdTipoAnexoRef As String
        Get
            Return _idTipoAnexoRef
        End Get
        Set
            _idTipoAnexoRef = Value
        End Set
    End Property

    Public Property Documentoref2 As String
        Get
            Return _documentoref2
        End Get
        Set
            _documentoref2 = Value
        End Set
    End Property

    Public Property Tasa As String
        Get
            Return _tasa
        End Get
        Set
            _tasa = Value
        End Set
    End Property

    Public Property MedioPago As Boolean
        Get
            Return _medioPago
        End Get
        Set
            _medioPago = Value
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

    Public Sub New(ByVal idcuenta As String, ByVal descripcion As String, ByVal idTipoAnexo As String, ByVal idNivelSaldo As String, ByVal documentoRef As String, ByVal pldoc As Decimal, ByVal fechaVencimiento As String, ByVal idMonedaReferencial As String, ByVal idCuentaCargo As String, ByVal idCuentaAbono As String, ByVal centroCosto As String, ByVal idRegistroCuenta As String, ByVal idTipoCuenta As String, ByVal cuentaAjusteACM As String, ByVal area As String, ByVal conciliacionBanco As String, ByVal idBalanceGeneral As String, ByVal idGastPerdfuncion As String, ByVal plingyp As String, ByVal idGastPerdNaturale As String, ByVal idCentroCosto As String, ByVal estado As String, ByVal fechaRegistro As System.DateTime, ByVal horaRegistro As String, ByVal usuario As String, ByVal idAlFormgreGasto As String, ByVal idAlFormCosto As String, ByVal idAlFormBalanGen As String, ByVal idAlFormGasPerFun As String, ByVal idAlFormGasPerNat As String, ByVal pvactfij As String, ByVal pvglodet As String, ByVal idTipoAnexoRef As String, ByVal documentoref2 As String, ByVal tasa As String, ByVal medioPago As Boolean)
        Me.New()
    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NPlanCuenta_Con)
        Dim parametros() As Object = {"@idcuenta", "@descripcion", "@idTipoAnexo", "@idNivelSaldo", "@documentoRef", "@pldoc", "@fechaVencimiento", "@idMonedaReferencial", "@idCuentaCargo", "@idCuentaAbono", "@centroCosto", "@idRegistroCuenta", "@idTipoCuenta", "@cuentaAjusteACM", "@area", "@conciliacionBanco", "@idBalanceGeneral", "@idGastPerdfuncion", "@plingyp", "@idGastPerdNaturale", "@idCentroCosto", "@estado", "@fechaRegistro", "@horaRegistro", "@usuario", "@idAlFormgreGasto", "@idAlFormCosto", "@idAlFormBalanGen", "@idAlFormGasPerFun", "@idAlFormGasPerNat", "@pvactfij", "@pvglodet", "@idTipoAnexoRef", "@documentoref2", "@tasa", "@medioPago"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.idcuenta, d.Descripcion, d.IdTipoAnexo, d.IdNivelSaldo, d.DocumentoRef, d.pldoc, d.FechaVencimiento, d.idMonedaReferencial, d.IdCuentaCargo, d.IdCuentaAbono, d.CentroCosto, d.idRegistroCuenta, d.IdTipoCuenta, d.CuentaAjusteACM, d.Area, d.conciliacionBanco, d.IdBalanceGeneral, d.IdGastPerdfuncion, d.plingyp, d.IdGastPerdNaturale, d.IdCentroCosto, d.Estado, d.FechaRegistro, d.HoraRegistro, d.Usuario, d.IdAlFormgreGasto, d.IdAlFormCosto, d.IdAlFormBalanGen, d.IdAlFormGasPerFun, d.IdAlFormGasPerNat, d.pvactfij, d.pvglodet, d.IdTipoAnexoRef, d.Documentoref2, d.Tasa, d.MedioPago}
        sql.EjecutarProcedure(Bd & ".dbo.Str_PlanCuenta_I", parametros, valores, tipoParametro, 36)
    End Sub
    Public Sub Actualizar(d As NPlanCuenta_Con)
        Dim parametros() As Object = {"@idcuenta", "@descripcion", "@idTipoAnexo", "@idNivelSaldo", "@documentoRef", "@pldoc", "@fechaVencimiento", "@idMonedaReferencial", "@idCuentaCargo", "@idCuentaAbono", "@centroCosto", "@idRegistroCuenta", "@idTipoCuenta", "@cuentaAjusteACM", "@area", "@conciliacionBanco", "@idBalanceGeneral", "@idGastPerdfuncion", "@plingyp", "@idGastPerdNaturale", "@idCentroCosto", "@estado", "@fechaRegistro", "@horaRegistro", "@usuario", "@idAlFormgreGasto", "@idAlFormCosto", "@idAlFormBalanGen", "@idAlFormGasPerFun", "@idAlFormGasPerNat", "@pvactfij", "@pvglodet", "@idTipoAnexoRef", "@documentoref2", "@tasa", "@medioPago"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.idcuenta, d.Descripcion, d.IdTipoAnexo, d.IdNivelSaldo, d.DocumentoRef, d.pldoc, d.FechaVencimiento, d.idMonedaReferencial, d.IdCuentaCargo, d.IdCuentaAbono, d.CentroCosto, d.idRegistroCuenta, d.IdTipoCuenta, d.CuentaAjusteACM, d.Area, d.conciliacionBanco, d.IdBalanceGeneral, d.IdGastPerdfuncion, d.plingyp, d.IdGastPerdNaturale, d.IdCentroCosto, d.Estado, d.FechaRegistro, d.HoraRegistro, d.Usuario, d.IdAlFormgreGasto, d.IdAlFormCosto, d.IdAlFormBalanGen, d.IdAlFormGasPerFun, d.IdAlFormGasPerNat, d.pvactfij, d.pvglodet, d.IdTipoAnexoRef, d.Documentoref2, d.Tasa, d.MedioPago}
        sql.EjecutarProcedure(Bd & ".dbo.Str_PlanCuenta_U", parametros, valores, tipoParametro, 36)
    End Sub
    Public Sub Eliminar(d As NPlanCuenta_Con)
        Dim parametros() As Object = {"@idcuenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idcuenta}
        sql.EjecutarProcedure(Bd & ".dbo.Str_PlanCuenta_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idcuenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd & ".dbo.Str_PlanCuenta_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NPlanCuenta_Con) As NPlanCuenta_Con
        Dim parametros() As Object = {"@idcuenta"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idcuenta}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL(Bd &".dbo.Str_PlanCuenta_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idcuenta = IIf(dt.Rows(0).Item("idcuenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcuenta"))
            d.Descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.IdTipoAnexo = IIf(dt.Rows(0).Item("idTipoAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoAnexo"))
            d.IdNivelSaldo = IIf(dt.Rows(0).Item("idNivelSaldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idNivelSaldo"))
            d.DocumentoRef = IIf(dt.Rows(0).Item("documentoRef") Is DBNull.Value, Nothing, dt.Rows(0).Item("documentoRef"))
            d.pldoc = IIf(dt.Rows(0).Item("pldoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("pldoc"))
            d.FechaVencimiento = IIf(dt.Rows(0).Item("fechaVencimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaVencimiento"))
            d.idMonedaReferencial = IIf(dt.Rows(0).Item("idMonedaReferencial") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMonedaReferencial"))
            d.IdCuentaCargo = IIf(dt.Rows(0).Item("idCuentaCargo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCuentaCargo"))
            d.IdCuentaAbono = IIf(dt.Rows(0).Item("idCuentaAbono") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCuentaAbono"))
            d.CentroCosto = IIf(dt.Rows(0).Item("centroCosto") Is DBNull.Value, Nothing, dt.Rows(0).Item("centroCosto"))
            d.idRegistroCuenta = IIf(dt.Rows(0).Item("idRegistroCuenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idRegistroCuenta"))
            d.IdTipoCuenta = IIf(dt.Rows(0).Item("idTipoCuenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoCuenta"))
            d.CuentaAjusteACM = IIf(dt.Rows(0).Item("cuentaAjusteACM") Is DBNull.Value, Nothing, dt.Rows(0).Item("cuentaAjusteACM"))
            d.Area = IIf(dt.Rows(0).Item("area") Is DBNull.Value, Nothing, dt.Rows(0).Item("area"))
            d.conciliacionBanco = IIf(dt.Rows(0).Item("conciliacionBanco") Is DBNull.Value, Nothing, dt.Rows(0).Item("conciliacionBanco"))
            d.IdBalanceGeneral = IIf(dt.Rows(0).Item("idBalanceGeneral") Is DBNull.Value, Nothing, dt.Rows(0).Item("idBalanceGeneral"))
            d.IdGastPerdfuncion = IIf(dt.Rows(0).Item("idGastPerdfuncion") Is DBNull.Value, Nothing, dt.Rows(0).Item("idGastPerdfuncion"))
            d.plingyp = IIf(dt.Rows(0).Item("plingyp") Is DBNull.Value, Nothing, dt.Rows(0).Item("plingyp"))
            d.IdGastPerdNaturale = IIf(dt.Rows(0).Item("idGastPerdNaturale") Is DBNull.Value, Nothing, dt.Rows(0).Item("idGastPerdNaturale"))
            d.IdCentroCosto = IIf(dt.Rows(0).Item("idCentroCosto") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCentroCosto"))
            d.Estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.FechaRegistro = IIf(dt.Rows(0).Item("fechaRegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaRegistro"))
            d.HoraRegistro = IIf(dt.Rows(0).Item("horaRegistro") Is DBNull.Value, Nothing, dt.Rows(0).Item("horaRegistro"))
            d.Usuario = IIf(dt.Rows(0).Item("usuario") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuario"))
            d.IdAlFormgreGasto = IIf(dt.Rows(0).Item("idAlFormgreGasto") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlFormgreGasto"))
            d.IdAlFormCosto = IIf(dt.Rows(0).Item("idAlFormCosto") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlFormCosto"))
            d.IdAlFormBalanGen = IIf(dt.Rows(0).Item("idAlFormBalanGen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlFormBalanGen"))
            d.IdAlFormGasPerFun = IIf(dt.Rows(0).Item("idAlFormGasPerFun") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlFormGasPerFun"))
            d.IdAlFormGasPerNat = IIf(dt.Rows(0).Item("idAlFormGasPerNat") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlFormGasPerNat"))
            d.pvactfij = IIf(dt.Rows(0).Item("pvactfij") Is DBNull.Value, Nothing, dt.Rows(0).Item("pvactfij"))
            d.pvglodet = IIf(dt.Rows(0).Item("pvglodet") Is DBNull.Value, Nothing, dt.Rows(0).Item("pvglodet"))
            d.IdTipoAnexoRef = IIf(dt.Rows(0).Item("idTipoAnexoRef") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoAnexoRef"))
            d.Documentoref2 = IIf(dt.Rows(0).Item("documentoref2") Is DBNull.Value, Nothing, dt.Rows(0).Item("documentoref2"))
            d.Tasa = IIf(dt.Rows(0).Item("tasa") Is DBNull.Value, Nothing, dt.Rows(0).Item("tasa"))
            d.MedioPago = IIf(dt.Rows(0).Item("medioPago") Is DBNull.Value, Nothing, dt.Rows(0).Item("medioPago"))
        Else
            d.Descripcion = Nothing
            d.IdTipoAnexo = Nothing
            d.IdNivelSaldo = Nothing
            d.DocumentoRef = Nothing
            d.pldoc = Nothing
            d.FechaVencimiento = Nothing
            d.idMonedaReferencial = Nothing
            d.IdCuentaCargo = Nothing
            d.IdCuentaAbono = Nothing
            d.CentroCosto = Nothing
            d.idRegistroCuenta = Nothing
            d.IdTipoCuenta = Nothing
            d.CuentaAjusteACM = Nothing
            d.Area = Nothing
            d.conciliacionBanco = Nothing
            d.IdBalanceGeneral = Nothing
            d.IdGastPerdfuncion = Nothing
            d.plingyp = Nothing
            d.IdGastPerdNaturale = Nothing
            d.IdCentroCosto = Nothing
            d.Estado = Nothing
            d.FechaRegistro = Nothing
            d.HoraRegistro = Nothing
            d.Usuario = Nothing
            d.IdAlFormgreGasto = Nothing
            d.IdAlFormCosto = Nothing
            d.IdAlFormBalanGen = Nothing
            d.IdAlFormGasPerFun = Nothing
            d.IdAlFormGasPerNat = Nothing
            d.pvactfij = Nothing
            d.pvglodet = Nothing
            d.IdTipoAnexoRef = Nothing
            d.Documentoref2 = Nothing
            d.Tasa = Nothing
            d.MedioPago = Nothing
        End If
        Return d
    End Function
#End Region

End Class
