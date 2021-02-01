Imports CapaDatos
Public Class NControl_Impresion
    Dim sql As New ClsConexion
#Region "Declarations"

    Private _item As Integer
    Private _idAlmacen As String
    Private _tipoDocumento As String
    Private _numeroDocumento As String
    Private _esImpreso As Boolean

#End Region

#Region "Properties"

    Public Property item As Integer
        Get
            Return _item
        End Get
        Set
            _item = Value
        End Set
    End Property

    Public Property IdAlmacen As String
        Get
            Return _idAlmacen
        End Get
        Set
            _idAlmacen = Value
        End Set
    End Property

    Public Property TipoDocumento As String
        Get
            Return _tipoDocumento
        End Get
        Set
            _tipoDocumento = Value
        End Set
    End Property

    Public Property NumeroDocumento As String
        Get
            Return _numeroDocumento
        End Get
        Set
            _numeroDocumento = Value
        End Set
    End Property

    Public Property EsImpreso As Boolean
        Get
            Return _esImpreso
        End Get
        Set
            _esImpreso = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal item As Integer, ByVal idAlmacen As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String, ByVal esImpreso As Boolean)
        Me.New()
    End Sub

#End Region
#Region "Metodos"

    ''' <summary>
    ''' Agrega un registro a la tabla control_imrpresion utiliza el store procedure
    ''' Str_control_impresion_I
    ''' </summary>
    Public Sub Agregar(d As NControl_Impresion)
        Dim parametros() As Object = {"@idAlmacen", "@tipoDocumento", "@numeroDocumento", "@esImpreso"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.IdAlmacen, d.TipoDocumento, d.NumeroDocumento, d.EsImpreso}
        sql.EjecutarProcedure("Str_control_impresion_I", parametros, valores, tipoParametro, 4)
    End Sub
    ''' <summary>
    ''' Elimina un registro de la tabla control_impresion Str_Control_Impresion_D
    ''' </summary>
    ''' <param name="d"></param>
    Public Sub Eliminar(d As NControl_Impresion)
        Dim parametros() As Object = {"@idAlmacen", "@tipoDocumento", "@numeroDocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAlmacen, d.TipoDocumento, d.NumeroDocumento}
        sql.EjecutarProcedure("Str_Control_Impresion_D", parametros, valores, tipoParametro, 3)
    End Sub
    ''' <summary>
    ''' Obtiene un registro de la clase Ncontrol_impresion 
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns>d</returns>
    Public Function Fila(d As NControl_Impresion) As NControl_Impresion
        Dim parametros() As Object = {"@idAlmacen", "@tipoDocumento", "@numeroDocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAlmacen, d.TipoDocumento, d.NumeroDocumento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Control_Impresion_S", parametros, valores, tipoParametro, 3).Tables(0)
        If dt.Rows.Count > 0 Then
            d.TipoDocumento = dt.Rows(0).Item("TipoDocumento")
            d.NumeroDocumento = dt.Rows(0).Item("NumeroDocumento")
            d.IdAlmacen = dt.Rows(0).Item("IdAlmacen")
            d.EsImpreso = dt.Rows(0).Item("EsImpreso")
            d.item = dt.Rows(0).Item("Item")
        Else

            d.EsImpreso = False
        End If
        Return d
    End Function
    Public Function filaDt(d As NControl_Impresion) As DataTable
        Dim parametros() As Object = {"@idAlmacen", "@tipoDocumento", "@numeroDocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAlmacen, d.TipoDocumento, d.NumeroDocumento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Control_Impresion_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function

    ''' <summary>
    ''' Actualiza la tabla control impresion
    ''' </summary>
    ''' <param name="d"></param>
    Public Sub Actualizar(d As NControl_Impresion)
        Dim parametros() As Object = {"@idAlmacen", "@tipoDocumento", "@numeroDocumento", "@Item", "@Esimpreso"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Int, SqlDbType.Bit}
        Dim valores() As Object = {d.IdAlmacen, d.TipoDocumento, d.NumeroDocumento, d.item, d.EsImpreso}
        sql.EjecutarProcedure("Str_Control_Impresion_U", parametros, valores, tipoParametro, 5)
    End Sub

    Public Function lista() As DataTable
        Dim cadena As String = " select item,doc.*,cp.EsImpreso from "
        cadena += " (select fechadocumento,IdTipoDocumento,(Serie+NumeroDocumento) as Numero,IdCliente,Nombrecliente,importetotal,estado from comprobante "
        'cadena += " where isnull(Estado,'')<>'A' "
        cadena += " union all "
        cadena += " select FechaDocumento,TipoDocumento as IdTipoDocumento,NumeroDocumento as Numero,IdCliente,NombreCliente,ImporteTotalVenta,situacion as estado from movimiento where tipodocumento='GR'"
        'cadena += " AND isnull(situacion,'')<>'A' "
        cadena += " ) as doc inner join  "
        cadena += " control_impresion cp on doc.IdTipoDocumento=cp.TipoDocumento and doc.Numero=cp.NumeroDocumento "
        Return sql.EjecutarConsulta("d", cadena).Tables(0)
    End Function
#End Region
End Class

