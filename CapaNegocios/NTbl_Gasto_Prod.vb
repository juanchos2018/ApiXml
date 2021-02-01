Imports CapaDatos
Public Class NTbl_Gasto_Prod
    Dim sql As New ClsConexion

#Region "Declarations"

    Private _item As Integer
    Private _idgencia As String
    Private _idAlmacen As String
    Private _idTipoDocumento As String
    Private _serie As String
    Private _numeroDocumento As String
    Private _idProveedor As String
    Private _idAlmacen_ref As String
    Private _tipoDocumento_Ref As String
    Private _numeroDocumento_Ref As String
    Private _idArticulo As String
    Private _idMoneda As String
    Private _importe As Decimal
    Private _tipoCambio As Decimal
    Private _importeMN As Decimal
    Private _importeUS As Decimal
    Private _tipoProrrateo As String
    Private _factor As Decimal
    Private _t_Prorrateo As Decimal
    Private _t_ProrrateoMN As Decimal
    Private _t_ProrrateoUS As Decimal
    Private _u_Prorrateo As Decimal
    Private _cantidad As Decimal
    Private _u_ProrrateoMN As Decimal
    Private _u_ProrrateoUS As Decimal
    Private _cuenta_Gasto As String

#End Region

#Region "Properties"

    Public Property Item As Integer
        Get
            Return _item
        End Get
        Set
            _item = Value
        End Set
    End Property

    Public Property Idgencia As String
        Get
            Return _idgencia
        End Get
        Set
            _idgencia = Value
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

    Public Property IdTipoDocumento As String
        Get
            Return _idTipoDocumento
        End Get
        Set
            _idTipoDocumento = Value
        End Set
    End Property

    Public Property serie As String
        Get
            Return _serie
        End Get
        Set
            _serie = Value
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

    Public Property IdProveedor As String
        Get
            Return _idProveedor
        End Get
        Set
            _idProveedor = Value
        End Set
    End Property

    Public Property IdAlmacen_ref As String
        Get
            Return _idAlmacen_ref
        End Get
        Set
            _idAlmacen_ref = Value
        End Set
    End Property

    Public Property TipoDocumento_Ref As String
        Get
            Return _tipoDocumento_Ref
        End Get
        Set
            _tipoDocumento_Ref = Value
        End Set
    End Property

    Public Property NumeroDocumento_Ref As String
        Get
            Return _numeroDocumento_Ref
        End Get
        Set
            _numeroDocumento_Ref = Value
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

    Public Property IdMoneda As String
        Get
            Return _idMoneda
        End Get
        Set
            _idMoneda = Value
        End Set
    End Property

    Public Property Importe As Decimal
        Get
            Return _importe
        End Get
        Set
            _importe = Value
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

    Public Property ImporteMN As Decimal
        Get
            Return _importeMN
        End Get
        Set
            _importeMN = Value
        End Set
    End Property

    Public Property ImporteUS As Decimal
        Get
            Return _importeUS
        End Get
        Set
            _importeUS = Value
        End Set
    End Property

    Public Property TipoProrrateo As String
        Get
            Return _tipoProrrateo
        End Get
        Set
            _tipoProrrateo = Value
        End Set
    End Property

    Public Property Factor As Decimal
        Get
            Return _factor
        End Get
        Set
            _factor = Value
        End Set
    End Property

    Public Property T_Prorrateo As Decimal
        Get
            Return _t_Prorrateo
        End Get
        Set
            _t_Prorrateo = Value
        End Set
    End Property

    Public Property T_ProrrateoMN As Decimal
        Get
            Return _t_ProrrateoMN
        End Get
        Set
            _t_ProrrateoMN = Value
        End Set
    End Property

    Public Property T_ProrrateoUS As Decimal
        Get
            Return _t_ProrrateoUS
        End Get
        Set
            _t_ProrrateoUS = Value
        End Set
    End Property

    Public Property U_Prorrateo As Decimal
        Get
            Return _u_Prorrateo
        End Get
        Set
            _u_Prorrateo = Value
        End Set
    End Property

    Public Property Cantidad As Decimal
        Get
            Return _cantidad
        End Get
        Set
            _cantidad = Value
        End Set
    End Property

    Public Property U_ProrrateoMN As Decimal
        Get
            Return _u_ProrrateoMN
        End Get
        Set
            _u_ProrrateoMN = Value
        End Set
    End Property

    Public Property U_ProrrateoUS As Decimal
        Get
            Return _u_ProrrateoUS
        End Get
        Set
            _u_ProrrateoUS = Value
        End Set
    End Property

    Public Property Cuenta_Gasto As String
        Get
            Return _cuenta_Gasto
        End Get
        Set
            _cuenta_Gasto = Value
        End Set
    End Property


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal item As Integer, ByVal idgencia As String, ByVal idAlmacen As String, ByVal idTipoDocumento As String, ByVal serie As String, ByVal numeroDocumento As String, ByVal idProveedor As String, ByVal idAlmacen_ref As String, ByVal tipoDocumento_Ref As String, ByVal numeroDocumento_Ref As String, ByVal idArticulo As String, ByVal idMoneda As String, ByVal importe As Decimal, ByVal tipoCambio As Decimal, ByVal importeMN As Decimal, ByVal importeUS As Decimal, ByVal tipoProrrateo As String, ByVal factor As Decimal, ByVal t_Prorrateo As Decimal, ByVal t_ProrrateoMN As Decimal, ByVal t_ProrrateoUS As Decimal, ByVal u_Prorrateo As Decimal, ByVal cantidad As Decimal, ByVal u_ProrrateoMN As Decimal, ByVal u_ProrrateoUS As Decimal, ByVal cuenta_Gasto As String)
        Me.New()
    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NTbl_Gasto_Prod)
        Dim parametros() As Object = {"@idgencia", "@idAlmacen", "@idTipoDocumento", "@serie", "@numeroDocumento", "@idProveedor", "@idAlmacen_ref", "@tipoDocumento_Ref", "@numeroDocumento_Ref", "@idArticulo", "@idMoneda", "@importe", "@tipoCambio", "@importeMN", "@importeUS", "@tipoProrrateo", "@factor", "@t_Prorrateo", "@t_ProrrateoMN", "@t_ProrrateoUS", "@u_Prorrateo", "@cantidad", "@u_ProrrateoMN", "@u_ProrrateoUS", "@cuenta_Gasto"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.Idgencia, d.IdAlmacen, d.IdTipoDocumento, d.serie, d.NumeroDocumento, d.IdProveedor, d.IdAlmacen_ref, d.TipoDocumento_Ref, d.NumeroDocumento_Ref, d.IdArticulo, d.IdMoneda, d.Importe, d.TipoCambio, d.ImporteMN, d.ImporteUS, d.TipoProrrateo, d.Factor, d.T_Prorrateo, d.T_ProrrateoMN, d.T_ProrrateoUS, d.U_Prorrateo, d.Cantidad, d.U_ProrrateoMN, d.U_ProrrateoUS, d.Cuenta_Gasto}
        sql.EjecutarProcedure("Str_Tbl_Gasto_Prod_I", parametros, valores, tipoParametro, 25)
    End Sub
    Public Sub Actualizar(d As NTbl_Gasto_Prod)
        Dim parametros() As Object = {"@item", "@idgencia", "@idAlmacen", "@idTipoDocumento", "@serie", "@numeroDocumento", "@idProveedor", "@idAlmacen_ref", "@tipoDocumento_Ref", "@numeroDocumento_Ref", "@idArticulo", "@idMoneda", "@importe", "@tipoCambio", "@importeMN", "@importeUS", "@tipoProrrateo", "@factor", "@t_Prorrateo", "@t_ProrrateoMN", "@t_ProrrateoUS", "@u_Prorrateo", "@cantidad", "@u_ProrrateoMN", "@u_ProrrateoUS", "@cuenta_Gasto"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.Item, d.Idgencia, d.IdAlmacen, d.IdTipoDocumento, d.serie, d.NumeroDocumento, d.IdProveedor, d.IdAlmacen_ref, d.TipoDocumento_Ref, d.NumeroDocumento_Ref, d.IdArticulo, d.IdMoneda, d.Importe, d.TipoCambio, d.ImporteMN, d.ImporteUS, d.TipoProrrateo, d.Factor, d.T_Prorrateo, d.T_ProrrateoMN, d.T_ProrrateoUS, d.U_Prorrateo, d.Cantidad, d.U_ProrrateoMN, d.U_ProrrateoUS, d.Cuenta_Gasto}
        sql.EjecutarProcedure("Str_Tbl_Gasto_Prod_U", parametros, valores, tipoParametro, 26)
    End Sub
    Public Sub Eliminar(d As NTbl_Gasto_Prod)
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.Item}
        sql.EjecutarProcedure("Str_Tbl_Gasto_Prod_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Gasto_Prod_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTbl_Gasto_Prod) As DataTable
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Gasto_Prod_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTbl_Gasto_Prod) As NTbl_Gasto_Prod
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.Int}
        Dim valores() As Object = {d.Item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Gasto_Prod_S", parametros, valores, tipoParametro, 26).Tables(0)
        If dt.Rows.Count > 0 Then
            d.Item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.Idgencia = IIf(dt.Rows(0).Item("idgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idgencia"))
            d.IdAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.IdTipoDocumento = IIf(dt.Rows(0).Item("idTipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoDocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.NumeroDocumento = IIf(dt.Rows(0).Item("numeroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento"))
            d.IdProveedor = IIf(dt.Rows(0).Item("idProveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idProveedor"))
            d.IdAlmacen_ref = IIf(dt.Rows(0).Item("idAlmacen_ref") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen_ref"))
            d.TipoDocumento_Ref = IIf(dt.Rows(0).Item("tipoDocumento_Ref") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoDocumento_Ref"))
            d.NumeroDocumento_Ref = IIf(dt.Rows(0).Item("numeroDocumento_Ref") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento_Ref"))
            d.IdArticulo = IIf(dt.Rows(0).Item("idArticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idArticulo"))
            d.IdMoneda = IIf(dt.Rows(0).Item("idMoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMoneda"))
            d.Importe = IIf(dt.Rows(0).Item("importe") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe"))
            d.TipoCambio = IIf(dt.Rows(0).Item("tipoCambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoCambio"))
            d.ImporteMN = IIf(dt.Rows(0).Item("importeMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeMN"))
            d.ImporteUS = IIf(dt.Rows(0).Item("importeUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeUS"))
            d.TipoProrrateo = IIf(dt.Rows(0).Item("tipoProrrateo") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoProrrateo"))
            d.Factor = IIf(dt.Rows(0).Item("factor") Is DBNull.Value, Nothing, dt.Rows(0).Item("factor"))
            d.T_Prorrateo = IIf(dt.Rows(0).Item("t_Prorrateo") Is DBNull.Value, Nothing, dt.Rows(0).Item("t_Prorrateo"))
            d.T_ProrrateoMN = IIf(dt.Rows(0).Item("t_ProrrateoMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("t_ProrrateoMN"))
            d.T_ProrrateoUS = IIf(dt.Rows(0).Item("t_ProrrateoUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("t_ProrrateoUS"))
            d.U_Prorrateo = IIf(dt.Rows(0).Item("u_Prorrateo") Is DBNull.Value, Nothing, dt.Rows(0).Item("u_Prorrateo"))
            d.Cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.U_ProrrateoMN = IIf(dt.Rows(0).Item("u_ProrrateoMN") Is DBNull.Value, Nothing, dt.Rows(0).Item("u_ProrrateoMN"))
            d.U_ProrrateoUS = IIf(dt.Rows(0).Item("u_ProrrateoUS") Is DBNull.Value, Nothing, dt.Rows(0).Item("u_ProrrateoUS"))
            d.Cuenta_Gasto = IIf(dt.Rows(0).Item("cuenta_Gasto") Is DBNull.Value, Nothing, dt.Rows(0).Item("cuenta_Gasto"))
        Else
            d.Idgencia = Nothing
            d.IdAlmacen = Nothing
            d.IdTipoDocumento = Nothing
            d.serie = Nothing
            d.NumeroDocumento = Nothing
            d.IdProveedor = Nothing
            d.IdAlmacen_ref = Nothing
            d.TipoDocumento_Ref = Nothing
            d.NumeroDocumento_Ref = Nothing
            d.IdArticulo = Nothing
            d.IdMoneda = Nothing
            d.Importe = Nothing
            d.TipoCambio = Nothing
            d.ImporteMN = Nothing
            d.ImporteUS = Nothing
            d.TipoProrrateo = Nothing
            d.Factor = Nothing
            d.T_Prorrateo = Nothing
            d.T_ProrrateoMN = Nothing
            d.T_ProrrateoUS = Nothing
            d.U_Prorrateo = Nothing
            d.Cantidad = Nothing
            d.U_ProrrateoMN = Nothing
            d.U_ProrrateoUS = Nothing
            d.Cuenta_Gasto = Nothing
        End If
        Return d
    End Function

    Public Function AsientoProrrateo(d As NTbl_Gasto_Prod) As DataTable
        Dim cad_asientovtas As String = "select Item, Idgencia, IdAlmacen, IdTipoDocumento, serie, NumeroDocumento, IdProveedor, IdAlmacen_ref, TipoDocumento_Ref, NumeroDocumento_Ref, IdArticulo,  "
        cad_asientovtas += " IdMoneda, Importe, TipoCambio, ImporteMN, ImporteUS, TipoProrrateo, Factor, T_Prorrateo, T_ProrrateoMN, T_ProrrateoUS, U_Prorrateo, Cantidad, U_ProrrateoMN, "
        cad_asientovtas += " U_ProrrateoUS, Cuenta_Gasto FROM Tbl_Gasto_Prod d "
        cad_asientovtas += " where d.IdAlmacen='" & d.IdAlmacen & "' and d.IdTipoDocumento='" & d.IdTipoDocumento & "' and d.Serie='" & d.serie & "' and d.NumeroDocumento='" & d.NumeroDocumento & "' and IdProveedor='" & d.IdProveedor & "'"
        Return sql.EjecutarConsulta("d", cad_asientovtas).Tables(0)
    End Function
#End Region

End Class
