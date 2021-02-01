Imports CapaDatos
Public Class NDetTipoSurtidor
    Dim sql As New ClsConexion
#Region "Declarations"
    Private _item_D As Integer
    Private _item As Integer
    Private _idArticulo As String
    Private _lado As String
#End Region
#Region "Properties"
    Public Property Item_D As Integer
        Get
            Return _item_D
        End Get
        Set
            _item_D = Value
        End Set
    End Property

    Public Property Item As Integer
        Get
            Return _item
        End Get
        Set
            _item = Value
        End Set
    End Property

    Public Property IdArticulo As String
        Get
            Return _idArticulo
        End Get
        Set
            _idArticulo = Value
        End Set
    End Property

    Public Property Lado As String
        Get
            Return _lado
        End Get
        Set
            _lado = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal item_D As Integer, ByVal item As Integer, ByVal idArticulo As String, ByVal lado As String)
        Me.New()
    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NDetTipoSurtidor)
        Dim parametros() As Object = {"@item", "@idArticulo", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.Item, d.IdArticulo, d.Lado}
        sql.EjecutarProcedure("Str_tbl_Det_Tipo_Surtidor_I", parametros, valores, tipoParametro, 3)
    End Sub
    Public Sub Actualizar(d As NDetTipoSurtidor)
        Dim parametros() As Object = {"@item_D", "@item", "@idArticulo", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Int, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.Item_D, d.Item, d.IdArticulo, d.Lado}
        sql.EjecutarProcedure("Str_tbl_Det_Tipo_Surtidor_U", parametros, valores, tipoParametro, 4)
    End Sub
    Public Sub Eliminar(d As NDetTipoSurtidor)
        Dim parametros() As Object = {"@item_D"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.Item_D}
        sql.ProcedureSQL("Str_tbl_Det_Tipo_Surtidor_D", parametros, valores, tipoParametro, 1)
    End Sub
    ''' <summary>
    ''' Obtiene un registro en funcion a la clave
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    Public Function Registro(d As NDetTipoSurtidor) As NDetTipoSurtidor
        Dim parametros() As Object = {"@item_D"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.Item_D}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Det_Tipo_Surtidor_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.IdArticulo = dt.Rows(0).Item("IdArticulo")
            d.Lado = dt.Rows(0).Item("lado")
            d.Item = dt.Rows(0).Item("Item")
            d.Item_D = dt.Rows(0).Item("Item_D")
        End If
        Return d
    End Function
    ''' <summary>
    ''' Listado de todo los registros de la table
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    Public Function Lista(d As NDetTipoSurtidor) As DataTable
        Dim parametros() As Object = {"@item_D"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Det_Tipo_Surtidor_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    ''' <summary>
    ''' Retorna los registros detalles de una item maestro
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    Public Function Detalle_maestro(d As NDetTipoSurtidor) As DataTable
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.Item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Det_Tipo_Surtidor_M", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function

#End Region

End Class
