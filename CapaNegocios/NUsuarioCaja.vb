Imports CapaDatos
Public Class NUsuarioCaja
    Dim sql As New ClsConexion

#Region "Declarations"

    Private _id As Integer
    Private _idCaja As String
    Private _idTurno As String
    Private _idUsuario As String
    Private _estado As String
    Private _habilitar As Boolean

#End Region

#Region "Properties"

    Public Property Id As Integer
        Get
            Return _id
        End Get
        Set
            _id = Value
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

    Public Property IdTurno As String
        Get
            Return _idTurno
        End Get
        Set
            _idTurno = Value
        End Set
    End Property

    Public Property IdUsuario As String
        Get
            Return _idUsuario
        End Get
        Set
            _idUsuario = Value
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

    Public Property Habilitar As Boolean
        Get
            Return _habilitar
        End Get
        Set
            _habilitar = Value
        End Set
    End Property


#End Region

#Region "Constructors"
    Public Sub New()
    End Sub

    Public Sub New(ByVal id As Integer, ByVal idCaja As String, ByVal idTurno As String, ByVal idUsuario As String, ByVal estado As String)
        Me.New()
    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NUsuarioCaja)
        Dim parametros() As Object = {"@idCaja", "@idTurno", "@idUsuario", "@estado", "@habilitar"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.IdCaja, d.IdTurno, d.IdUsuario, d.Estado, d.Habilitar}
        sql.EjecutarProcedure("Str_tbl_Usuario_Caja_I", parametros, valores, tipoParametro, 5)
    End Sub
    Public Sub Actualizar(d As NUsuarioCaja)
        Dim parametros() As Object = {"@id", "@idCaja", "@idTurno", "@idUsuario", "@estado", "@habilitar"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.Id, d.IdCaja, d.IdTurno, d.IdUsuario, d.Estado, d.Habilitar}
        sql.EjecutarProcedure("Str_tbl_Usuario_Caja_U", parametros, valores, tipoParametro, 6)
    End Sub
    Public Sub Eliminar(d As NUsuarioCaja)
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.Id}
        sql.EjecutarProcedure("Str_tbl_Usuario_Caja_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Registro(d As NUsuarioCaja) As NUsuarioCaja
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.Id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Usuario_Caja_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.Id = dt.Rows(0).Item("Id")
            d.IdCaja = dt.Rows(0).Item("IdCaja")
            d.IdTurno = dt.Rows(0).Item("IdTurno")
            d.IdUsuario = dt.Rows(0).Item("IdUsuario")
            d.Estado = dt.Rows(0).Item("Estado")
            d.Habilitar = dt.Rows(0).Item("habilitar")
        Else
            d.IdCaja = ""
            d.IdTurno = ""
            d.IdUsuario = ""
            d.Habilitar = True
        End If
        Return d
    End Function
    Public Function Registro(IdUsuario) As NUsuarioCaja
        Dim dt As New DataTable
        Dim d As New NUsuarioCaja
        dt = sql.EjecutarConsulta("d", "select * from  tbl_Usuario_Caja where IdUsuario='" & IdUsuario & "'").Tables(0)
        If dt.Rows.Count > 0 Then
            d.Id = dt.Rows(0).Item("Id")
            d.IdCaja = dt.Rows(0).Item("IdCaja")
            d.IdTurno = dt.Rows(0).Item("IdTurno")
            d.IdUsuario = dt.Rows(0).Item("IdUsuario")
            d.Estado = dt.Rows(0).Item("Estado")
            d.Habilitar = dt.Rows(0).Item("habilitar")
        Else
            d.IdCaja = ""
            d.IdTurno = ""
            d.IdUsuario = ""
            d.Habilitar = True
        End If
        Return d
    End Function

    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Usuario_Caja_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
#End Region

End Class
