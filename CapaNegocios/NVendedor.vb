Imports CapaDatos
Public Class NVendedor
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idVendedor As String
    Public Property nombre As String
    Public Property dNI As String
    Public Property direccion As String
    Public Property telefono As String
    Public Property idTipoVendedor As String
    Public Property idLinea As String
    Public Property idZonaVenta As String
    Public Property email As String
    Public Property fechaIngreso As System.DateTime
    Public Property usuarioCrea As String
    Public Property fechaCrea As System.DateTime
    Public Property usuarioMod As String
    Public Property fechaMod As System.DateTime
    Public Property fechaCese As System.DateTime
    Public Property empresa As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NVendedor)

        Dim parametros() As Object = {"@direccion", "@dNI", "@email", "@empresa", "@fechaCese", "@fechaCrea", "@fechaIngreso", "@fechaMod", "@idLinea", "@idTipoVendedor", "@idVendedor", "@idZonaVenta", "@nombre", "@telefono", "@usuarioCrea", "@usuarioMod"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.direccion, d.dNI, d.email, d.empresa, d.fechaCese, d.fechaCrea, d.fechaIngreso, d.fechaMod, d.idLinea, d.idTipoVendedor, d.idVendedor, d.idZonaVenta, d.nombre, d.telefono, d.usuarioCrea, d.usuarioMod}
        sql.EjecutarProcedure("Str_Vendedor_I", parametros, valores, tipoParametro, 16)
    End Sub
    Public Sub Actualizar(d As NVendedor)
        Dim parametros() As Object = {"@direccion", "@dNI", "@email", "@empresa", "@fechaCese", "@fechaCrea", "@fechaIngreso", "@fechaMod", "@idLinea", "@idTipoVendedor", "@idVendedor", "@idZonaVenta", "@nombre", "@telefono", "@usuarioCrea", "@usuarioMod"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.direccion, d.dNI, d.email, d.empresa, d.fechaCese, d.fechaCrea, d.fechaIngreso, d.fechaMod, d.idLinea, d.idTipoVendedor, d.idVendedor, d.idZonaVenta, d.nombre, d.telefono, d.usuarioCrea, d.usuarioMod}
        sql.EjecutarProcedure("Str_Vendedor_U", parametros, valores, tipoParametro, 16)
    End Sub
    Public Sub Eliminar(d As NVendedor)
        Dim parametros() As Object = {"@idVendedor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idVendedor}
        sql.EjecutarProcedure("Str_Vendedor_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idVendedor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Vendedor_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NVendedor) As DataTable
        Dim parametros() As Object = {"@idVendedor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idVendedor}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Vendedor_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NVendedor) As NVendedor
        Dim parametros() As Object = {"@idVendedor"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idVendedor}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Vendedor_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.dNI = IIf(dt.Rows(0).Item("dNI") Is DBNull.Value, Nothing, dt.Rows(0).Item("dNI"))
            d.email = IIf(dt.Rows(0).Item("email") Is DBNull.Value, Nothing, dt.Rows(0).Item("email"))
            d.empresa = IIf(dt.Rows(0).Item("empresa") Is DBNull.Value, Nothing, dt.Rows(0).Item("empresa"))
            d.fechaCese = IIf(dt.Rows(0).Item("fechaCese") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCese"))
            d.fechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.fechaIngreso = IIf(dt.Rows(0).Item("fechaIngreso") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaIngreso"))
            d.fechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.idLinea = IIf(dt.Rows(0).Item("idLinea") Is DBNull.Value, Nothing, dt.Rows(0).Item("idLinea"))
            d.idTipoVendedor = IIf(dt.Rows(0).Item("idTipoVendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoVendedor"))
            d.idVendedor = IIf(dt.Rows(0).Item("idVendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idVendedor"))
            d.idZonaVenta = IIf(dt.Rows(0).Item("idZonaVenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("idZonaVenta"))
            d.nombre = IIf(dt.Rows(0).Item("nombre") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombre"))
            d.telefono = IIf(dt.Rows(0).Item("telefono") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono"))
            d.usuarioCrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.usuarioMod = IIf(dt.Rows(0).Item("usuarioMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioMod"))
        Else
            d.direccion = Nothing
            d.dNI = Nothing
            d.email = Nothing
            d.empresa = Nothing
            d.fechaCese = Nothing
            d.fechaCrea = Nothing
            d.fechaIngreso = Nothing
            d.fechaMod = Nothing
            d.idLinea = Nothing
            d.idTipoVendedor = Nothing
            d.idVendedor = Nothing
            d.idZonaVenta = Nothing
            d.nombre = Nothing
            d.telefono = Nothing
            d.usuarioCrea = Nothing
            d.usuarioMod = Nothing
        End If
        Return d
    End Function
    ''' <summary>
    ''' Retorna un resumen de ventas por vendedor y por familia
    ''' </summary>
    ''' <param name="fechai"></param>
    ''' <param name="fechaf"></param>
    ''' <returns></returns>
    Public Function VentaVendedorResumen(fechai As DateTime, fechaf As DateTime, Optional ve As DataTable = Nothing, Optional plan As DataTable = Nothing) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF", "@ven", "@idfamilia"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Structured, SqlDbType.Structured}
        Dim valores() As Object = {fechai, fechaf, ve, plan}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_VentaVendedorResumen", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function

    ''' <summary>
    ''' Retorna las ventas por comprobante y producto
    ''' </summary>
    ''' <param name="fechai"></param>
    ''' <param name="fechaf"></param>
    ''' <returns></returns>
    Public Function VentaVendedor(fechai As DateTime, fechaf As DateTime, Optional ve As DataTable = Nothing, Optional plan As DataTable = Nothing) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF", "@ven", "@idfamilia"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Structured, SqlDbType.Structured}
        Dim valores() As Object = {fechai, fechaf, ve, plan}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_VentaVendedor", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function VentaVendedor(fechai As DateTime, fechaf As DateTime) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime}
        Dim valores() As Object = {fechai, fechaf}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_VentaVendedor", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    ''' <summary>
    ''' retorna las ventas por comprobante 
    ''' </summary>
    ''' <param name="fechai"></param>
    ''' <param name="fechaf"></param>
    ''' <returns></returns>
    Public Function VentaxVendedor(fechai As DateTime, fechaf As DateTime, Optional ve As DataTable = Nothing, Optional plan As DataTable = Nothing) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF", "@ven", "@plan"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Structured, SqlDbType.Structured}
        Dim valores() As Object = {fechai, fechaf, ve, plan}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_VentasxVendedor", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function lista_vendedor() As DataTable
        Dim s As String = " select cast(0 as bit) as ok,IdVendedor,DNI,Nombre from vendedor "
        s += " order by Nombre "
        Return sql.EjecutarConsulta("D", s).Tables(0)
    End Function
#End Region
End Class
