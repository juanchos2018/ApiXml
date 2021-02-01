Imports CapaDatos


Public Class NFormaVenta
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _idFormaVenta As String
    Private _descripcion As String
    Private _diasVencidos As Decimal
    Private _numeroLetra As Decimal
    Private _diaVencimientoL01 As Decimal
    Private _diaVencimientoL02 As Decimal
    Private _diaVencimientoL03 As Decimal
    Private _diaVencimientoL04 As Decimal
    Private _diaVencimientoL05 As Decimal
    Private _diaVencimientoL06 As Decimal
    Private _diaVencimientoL07 As Decimal
    Private _diaVencimientoL08 As Decimal
    Private _diaVencimientoL09 As Decimal
    Private _diaVencimientoL10 As Decimal
    Private _diaVencimientoL11 As Decimal
    Private _fDiaVencimientoL12 As Decimal
    Private _diaVencimientoL13 As Decimal
    Private _diaVencimientoL14 As Decimal
    Private _diaVencimientoL15 As Decimal
    Private _diaVencimientoL16 As Decimal
    Private _usuarioCrea As String
    Private _usuarioMod As String
    Private _fechaCrea As System.DateTime
    Private _fechaMod As System.DateTime
    Private _fechaVenc As System.DateTime

#End Region

#Region "Properties"

    Public Property IdFormaVenta As String
        Get
            Return _idFormaVenta
        End Get
        Set
            _idFormaVenta = Value
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

    Public Property DiasVencidos As Decimal
        Get
            Return _diasVencidos
        End Get
        Set
            _diasVencidos = Value
        End Set
    End Property

    Public Property NumeroLetra As Decimal
        Get
            Return _numeroLetra
        End Get
        Set
            _numeroLetra = Value
        End Set
    End Property

    Public Property DiaVencimientoL01 As Decimal
        Get
            Return _diaVencimientoL01
        End Get
        Set
            _diaVencimientoL01 = Value
        End Set
    End Property

    Public Property DiaVencimientoL02 As Decimal
        Get
            Return _diaVencimientoL02
        End Get
        Set
            _diaVencimientoL02 = Value
        End Set
    End Property

    Public Property DiaVencimientoL03 As Decimal
        Get
            Return _diaVencimientoL03
        End Get
        Set
            _diaVencimientoL03 = Value
        End Set
    End Property

    Public Property DiaVencimientoL04 As Decimal
        Get
            Return _diaVencimientoL04
        End Get
        Set
            _diaVencimientoL04 = Value
        End Set
    End Property

    Public Property DiaVencimientoL05 As Decimal
        Get
            Return _diaVencimientoL05
        End Get
        Set
            _diaVencimientoL05 = Value
        End Set
    End Property

    Public Property DiaVencimientoL06 As Decimal
        Get
            Return _diaVencimientoL06
        End Get
        Set
            _diaVencimientoL06 = Value
        End Set
    End Property

    Public Property DiaVencimientoL07 As Decimal
        Get
            Return _diaVencimientoL07
        End Get
        Set
            _diaVencimientoL07 = Value
        End Set
    End Property

    Public Property DiaVencimientoL08 As Decimal
        Get
            Return _diaVencimientoL08
        End Get
        Set
            _diaVencimientoL08 = Value
        End Set
    End Property

    Public Property DiaVencimientoL09 As Decimal
        Get
            Return _diaVencimientoL09
        End Get
        Set
            _diaVencimientoL09 = Value
        End Set
    End Property

    Public Property DiaVencimientoL10 As Decimal
        Get
            Return _diaVencimientoL10
        End Get
        Set
            _diaVencimientoL10 = Value
        End Set
    End Property

    Public Property DiaVencimientoL11 As Decimal
        Get
            Return _diaVencimientoL11
        End Get
        Set
            _diaVencimientoL11 = Value
        End Set
    End Property

    Public Property FDiaVencimientoL12 As Decimal
        Get
            Return _fDiaVencimientoL12
        End Get
        Set
            _fDiaVencimientoL12 = Value
        End Set
    End Property

    Public Property DiaVencimientoL13 As Decimal
        Get
            Return _diaVencimientoL13
        End Get
        Set
            _diaVencimientoL13 = Value
        End Set
    End Property

    Public Property DiaVencimientoL14 As Decimal
        Get
            Return _diaVencimientoL14
        End Get
        Set
            _diaVencimientoL14 = Value
        End Set
    End Property

    Public Property DiaVencimientoL15 As Decimal
        Get
            Return _diaVencimientoL15
        End Get
        Set
            _diaVencimientoL15 = Value
        End Set
    End Property

    Public Property DiaVencimientoL16 As Decimal
        Get
            Return _diaVencimientoL16
        End Get
        Set
            _diaVencimientoL16 = Value
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

    Public Property FechaVenc As System.DateTime
        Get
            Return _fechaVenc
        End Get
        Set
            _fechaVenc = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal idFormaVenta As String, ByVal descripcion As String, ByVal diasVencidos As Decimal, ByVal numeroLetra As Decimal, ByVal diaVencimientoL01 As Decimal, ByVal diaVencimientoL02 As Decimal, ByVal diaVencimientoL03 As Decimal, ByVal diaVencimientoL04 As Decimal, ByVal diaVencimientoL05 As Decimal, ByVal diaVencimientoL06 As Decimal, ByVal diaVencimientoL07 As Decimal, ByVal diaVencimientoL08 As Decimal, ByVal diaVencimientoL09 As Decimal, ByVal diaVencimientoL10 As Decimal, ByVal diaVencimientoL11 As Decimal, ByVal fDiaVencimientoL12 As Decimal, ByVal diaVencimientoL13 As Decimal, ByVal diaVencimientoL14 As Decimal, ByVal diaVencimientoL15 As Decimal, ByVal diaVencimientoL16 As Decimal, ByVal usuarioCrea As String, ByVal usuarioMod As String, ByVal fechaCrea As System.DateTime, ByVal fechaMod As System.DateTime, ByVal fechaVenc As System.DateTime)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    Public Function lista() As DataTable
        Return sql.EjecutarConsulta("F", "select idformaventa, Descripcion as Descripcion,DiasVencidos from formaventa order by idFormaVenta").Tables(0)
    End Function
    Public Function Item(t As NFormaVenta) As NFormaVenta
        Dim ca As String = " select idformaventa, Descripcion,DiasVencidos from formaventa where IdFormaVenta='" & t.IdFormaVenta & "' order by idFormaVenta "
        Dim dt As New DataTable
        dt = sql.EjecutarConsulta("ca", ca).Tables(0)
        If dt.Rows.Count > 0 Then
            t.IdFormaVenta = dt.Rows(0).Item("IdFormaVenta")
            t.Descripcion = dt.Rows(0).Item("Descripcion")
            t.DiasVencidos = dt.Rows(0).Item("DiasVencidos")
        End If
        Return t
    End Function

    Public Sub Agregar(d As NFormaVenta)

        Dim parametros() As Object = {"@idFormaVenta", "@descripcion", "@diasVencidos", "@numeroLetra", "@diaVencimientoL01", "@diaVencimientoL02", "@diaVencimientoL03", "@diaVencimientoL04", "@diaVencimientoL05", "@diaVencimientoL06", "@diaVencimientoL07", "@diaVencimientoL08", "@diaVencimientoL09", "@diaVencimientoL10", "@diaVencimientoL11", "@fDiaVencimientoL12", "@diaVencimientoL13", "@diaVencimientoL14", "@diaVencimientoL15", "@diaVencimientoL16", "@usuarioCrea", "@usuarioMod", "@fechaCrea", "@fechaMod", "@fechaVenc"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime}
        Dim valores() As Object = {d.IdFormaVenta, d.Descripcion, d.DiasVencidos, d.NumeroLetra, d.DiaVencimientoL01, d.DiaVencimientoL02, d.DiaVencimientoL03, d.DiaVencimientoL04, d.DiaVencimientoL05, d.DiaVencimientoL06, d.DiaVencimientoL07, d.DiaVencimientoL08, d.DiaVencimientoL09, d.DiaVencimientoL10, d.DiaVencimientoL11, d.FDiaVencimientoL12, d.DiaVencimientoL13, d.DiaVencimientoL14, d.DiaVencimientoL15, d.DiaVencimientoL16, d.UsuarioCrea, d.UsuarioMod, d.FechaCrea, d.FechaMod, d.FechaVenc}
        sql.EjecutarProcedure("Str_FormaVenta_I", parametros, valores, tipoParametro, 25)
    End Sub
    Public Sub Actualizar(d As NFormaVenta)
        Dim parametros() As Object = {"@idFormaVenta", "@descripcion", "@diasVencidos", "@numeroLetra", "@diaVencimientoL01", "@diaVencimientoL02", "@diaVencimientoL03", "@diaVencimientoL04", "@diaVencimientoL05", "@diaVencimientoL06", "@diaVencimientoL07", "@diaVencimientoL08", "@diaVencimientoL09", "@diaVencimientoL10", "@diaVencimientoL11", "@fDiaVencimientoL12", "@diaVencimientoL13", "@diaVencimientoL14", "@diaVencimientoL15", "@diaVencimientoL16", "@usuarioCrea", "@usuarioMod", "@fechaCrea", "@fechaMod", "@fechaVenc"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime}
        Dim valores() As Object = {d.IdFormaVenta, d.Descripcion, d.DiasVencidos, d.NumeroLetra, d.DiaVencimientoL01, d.DiaVencimientoL02, d.DiaVencimientoL03, d.DiaVencimientoL04, d.DiaVencimientoL05, d.DiaVencimientoL06, d.DiaVencimientoL07, d.DiaVencimientoL08, d.DiaVencimientoL09, d.DiaVencimientoL10, d.DiaVencimientoL11, d.FDiaVencimientoL12, d.DiaVencimientoL13, d.DiaVencimientoL14, d.DiaVencimientoL15, d.DiaVencimientoL16, d.UsuarioCrea, d.UsuarioMod, d.FechaCrea, d.FechaMod, d.FechaVenc}
        sql.EjecutarProcedure("Str_FormaVenta_U", parametros, valores, tipoParametro, 25)
    End Sub
    Public Sub Eliminar(d As NFormaVenta)
        Dim parametros() As Object = {"@idFormaVenta"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.IdFormaVenta}
        sql.EjecutarProcedure("Str_FormaVenta_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function ListaG() As DataTable
        Dim parametros() As Object = {"@idFormaVenta"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_FormaVenta_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function ListaG(d As NFormaVenta) As DataTable
        Dim parametros() As Object = {"@idFormaVenta"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.IdFormaVenta}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_FormaVenta_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function

    Public Function Registro(d As NFormaVenta) As NFormaVenta
        Dim parametros() As Object = {"@idFormaVenta"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.IdFormaVenta}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_FormaVenta_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdFormaVenta = IIf(dt.Rows(0).Item("idFormaVenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idFormaVenta"))
            d.Descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.DiasVencidos = IIf(dt.Rows(0).Item("diasVencidos") Is DBNull.Value, Nothing, dt.Rows(0).Item("diasVencidos"))
            d.NumeroLetra = IIf(dt.Rows(0).Item("numeroLetra") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroLetra"))
            d.DiaVencimientoL01 = IIf(dt.Rows(0).Item("diaVencimientoL01") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL01"))
            d.DiaVencimientoL02 = IIf(dt.Rows(0).Item("diaVencimientoL02") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL02"))
            d.DiaVencimientoL03 = IIf(dt.Rows(0).Item("diaVencimientoL03") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL03"))
            d.DiaVencimientoL04 = IIf(dt.Rows(0).Item("diaVencimientoL04") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL04"))
            d.DiaVencimientoL05 = IIf(dt.Rows(0).Item("diaVencimientoL05") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL05"))
            d.DiaVencimientoL06 = IIf(dt.Rows(0).Item("diaVencimientoL06") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL06"))
            d.DiaVencimientoL07 = IIf(dt.Rows(0).Item("diaVencimientoL07") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL07"))
            d.DiaVencimientoL08 = IIf(dt.Rows(0).Item("diaVencimientoL08") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL08"))
            d.DiaVencimientoL09 = IIf(dt.Rows(0).Item("diaVencimientoL09") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL09"))
            d.DiaVencimientoL10 = IIf(dt.Rows(0).Item("diaVencimientoL10") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL10"))
            d.DiaVencimientoL11 = IIf(dt.Rows(0).Item("diaVencimientoL11") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL11"))
            d.FDiaVencimientoL12 = IIf(dt.Rows(0).Item("fDiaVencimientoL12") Is DBNull.Value, Nothing, dt.Rows(0).Item("fDiaVencimientoL12"))
            d.DiaVencimientoL13 = IIf(dt.Rows(0).Item("diaVencimientoL13") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL13"))
            d.DiaVencimientoL14 = IIf(dt.Rows(0).Item("diaVencimientoL14") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL14"))
            d.DiaVencimientoL15 = IIf(dt.Rows(0).Item("diaVencimientoL15") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL15"))
            d.DiaVencimientoL16 = IIf(dt.Rows(0).Item("diaVencimientoL16") Is DBNull.Value, Nothing, dt.Rows(0).Item("diaVencimientoL16"))
            d.UsuarioCrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.UsuarioMod = IIf(dt.Rows(0).Item("usuarioMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioMod"))
            d.FechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.FechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.FechaVenc = IIf(dt.Rows(0).Item("fechaVenc") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaVenc"))
        Else
            d.IdFormaVenta = Nothing
            d.Descripcion = Nothing
            d.DiasVencidos = Nothing
            d.NumeroLetra = Nothing
            d.DiaVencimientoL01 = Nothing
            d.DiaVencimientoL02 = Nothing
            d.DiaVencimientoL03 = Nothing
            d.DiaVencimientoL04 = Nothing
            d.DiaVencimientoL05 = Nothing
            d.DiaVencimientoL06 = Nothing
            d.DiaVencimientoL07 = Nothing
            d.DiaVencimientoL08 = Nothing
            d.DiaVencimientoL09 = Nothing
            d.DiaVencimientoL10 = Nothing
            d.DiaVencimientoL11 = Nothing
            d.FDiaVencimientoL12 = Nothing
            d.DiaVencimientoL13 = Nothing
            d.DiaVencimientoL14 = Nothing
            d.DiaVencimientoL15 = Nothing
            d.DiaVencimientoL16 = Nothing
            d.UsuarioCrea = Nothing
            d.UsuarioMod = Nothing
            d.FechaCrea = Nothing
            d.FechaMod = Nothing
            d.FechaVenc = Nothing
        End If
        Return d
    End Function

#End Region
End Class
