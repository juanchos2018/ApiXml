Imports CapaDatos
Public Class NComprobante_Aduana
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numerodocumento As String
    Public Property idregimen_aduana As String
    Public Property iddespacho_aduana As String
    Public Property anio_dam As String
    Public Property numero_dam As String
    Public Property importe_fob As Decimal
    Public Property importe_seguro As Decimal
    Public Property importe_flete As Decimal
    Public Property importe_cif As Decimal
    Public Property importe_advalorem As Decimal
    Public Property importe_igv As Decimal
    Public Property importe_isc As Decimal
    Public Property importe_ipm As Decimal
    Public Property cantidad_bultos As Decimal
    Public Property peso_bruto As Decimal
    Public Property fecha_llegada As System.DateTime
    Public Property awb_bl As String
    Public Property nave As String
    Public Property mercancia As String
    Public Property IdAlmacen As String
    Public Property IdAgencia As String


#End Region

#Region "Constructors"
    Public Sub New()

    End Sub
#End Region
#Region "Metodos"
    Public Sub Agregar(d As NComprobante_Aduana)

        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@idregimen_aduana", "@iddespacho_aduana", "@anio_dam", "@numero_dam", "@importe_fob", "@importe_seguro", "@importe_flete", "@importe_cif", "@importe_advalorem", "@importe_igv", "@importe_isc", "@importe_ipm", "@cantidad_bultos", "@peso_bruto", "@fecha_llegada", "@awb_bl", "@nave", "@mercancia", "@idalmacen", "@idagencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.idregimen_aduana, d.iddespacho_aduana, d.anio_dam, d.numero_dam, d.importe_fob, d.importe_seguro, d.importe_flete, d.importe_cif, d.importe_advalorem, d.importe_igv, d.importe_isc, d.importe_ipm, d.cantidad_bultos, d.peso_bruto, d.fecha_llegada, d.awb_bl, d.nave, d.mercancia, d.idalmacen, d.idagencia}
        sql.EjecutarProcedure("Str_Comprobante_Aduana_I", parametros, valores, tipoParametro, 23)
    End Sub
    Public Sub Actualizar(d As NComprobante_Aduana)
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@idregimen_aduana", "@iddespacho_aduana", "@anio_dam", "@numero_dam", "@importe_fob", "@importe_seguro", "@importe_flete", "@importe_cif", "@importe_advalorem", "@importe_igv", "@importe_isc", "@importe_ipm", "@cantidad_bultos", "@peso_bruto", "@fecha_llegada", "@awb_bl", "@nave", "@mercancia", "@idalmacen", "@idagencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.idregimen_aduana, d.iddespacho_aduana, d.anio_dam, d.numero_dam, d.importe_fob, d.importe_seguro, d.importe_flete, d.importe_cif, d.importe_advalorem, d.importe_igv, d.importe_isc, d.importe_ipm, d.cantidad_bultos, d.peso_bruto, d.fecha_llegada, d.awb_bl, d.nave, d.mercancia, d.idalmacen, d.idagencia}
        sql.EjecutarProcedure("Str_Comprobante_Aduana_U", parametros, valores, tipoParametro, 23)
    End Sub
    Public Function Agregar(d As NComprobante_Aduana, Retornatable As Boolean) As NComprobante_Aduana

        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@idregimen_aduana", "@iddespacho_aduana", "@anio_dam", "@numero_dam", "@importe_fob", "@importe_seguro", "@importe_flete", "@importe_cif", "@importe_advalorem", "@importe_igv", "@importe_isc", "@importe_ipm", "@cantidad_bultos", "@peso_bruto", "@fecha_llegada", "@awb_bl", "@nave", "@mercancia", "@idalmacen", "@idagencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.idregimen_aduana, d.iddespacho_aduana, d.anio_dam, d.numero_dam, d.importe_fob, d.importe_seguro, d.importe_flete, d.importe_cif, d.importe_advalorem, d.importe_igv, d.importe_isc, d.importe_ipm, d.cantidad_bultos, d.peso_bruto, d.fecha_llegada, d.awb_bl, d.nave, d.mercancia, d.idalmacen, d.idagencia}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Comprobante_Aduana_I_S", parametros, valores, tipoParametro, 23).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.idregimen_aduana = IIf(dt.Rows(0).Item("idregimen_aduana") Is DBNull.Value, Nothing, dt.Rows(0).Item("idregimen_aduana"))
            d.iddespacho_aduana = IIf(dt.Rows(0).Item("iddespacho_aduana") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddespacho_aduana"))
            d.anio_dam = IIf(dt.Rows(0).Item("anio_dam") Is DBNull.Value, Nothing, dt.Rows(0).Item("anio_dam"))
            d.numero_dam = IIf(dt.Rows(0).Item("numero_dam") Is DBNull.Value, Nothing, dt.Rows(0).Item("numero_dam"))
            d.importe_fob = IIf(dt.Rows(0).Item("importe_fob") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_fob"))
            d.importe_seguro = IIf(dt.Rows(0).Item("importe_seguro") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_seguro"))
            d.importe_flete = IIf(dt.Rows(0).Item("importe_flete") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_flete"))
            d.importe_cif = IIf(dt.Rows(0).Item("importe_cif") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_cif"))
            d.importe_advalorem = IIf(dt.Rows(0).Item("importe_advalorem") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_advalorem"))
            d.importe_igv = IIf(dt.Rows(0).Item("importe_igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_igv"))
            d.importe_isc = IIf(dt.Rows(0).Item("importe_isc") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_isc"))
            d.importe_ipm = IIf(dt.Rows(0).Item("importe_ipm") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_ipm"))
            d.cantidad_bultos = IIf(dt.Rows(0).Item("cantidad_bultos") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad_bultos"))
            d.peso_bruto = IIf(dt.Rows(0).Item("peso_bruto") Is DBNull.Value, Nothing, dt.Rows(0).Item("peso_bruto"))
            d.fecha_llegada = IIf(dt.Rows(0).Item("fecha_llegada") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecha_llegada"))
            d.awb_bl = IIf(dt.Rows(0).Item("awb_bl") Is DBNull.Value, Nothing, dt.Rows(0).Item("awb_bl"))
            d.nave = IIf(dt.Rows(0).Item("nave") Is DBNull.Value, Nothing, dt.Rows(0).Item("nave"))
            d.mercancia = IIf(dt.Rows(0).Item("mercancia") Is DBNull.Value, Nothing, dt.Rows(0).Item("mercancia"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
        Else
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.idregimen_aduana = Nothing
            d.iddespacho_aduana = Nothing
            d.anio_dam = Nothing
            d.numero_dam = Nothing
            d.importe_fob = Nothing
            d.importe_seguro = Nothing
            d.importe_flete = Nothing
            d.importe_cif = Nothing
            d.importe_advalorem = Nothing
            d.importe_igv = Nothing
            d.importe_isc = Nothing
            d.importe_ipm = Nothing
            d.cantidad_bultos = Nothing
            d.peso_bruto = Nothing
            d.fecha_llegada = Nothing
            d.awb_bl = Nothing
            d.nave = Nothing
            d.mercancia = Nothing
            d.idalmacen = Nothing
            d.idagencia = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NComprobante_Aduana, Retornatable As Boolean) As NComprobante_Aduana

        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@idregimen_aduana", "@iddespacho_aduana", "@anio_dam", "@numero_dam", "@importe_fob", "@importe_seguro", "@importe_flete", "@importe_cif", "@importe_advalorem", "@importe_igv", "@importe_isc", "@importe_ipm", "@cantidad_bultos", "@peso_bruto", "@fecha_llegada", "@awb_bl", "@nave", "@mercancia", "@idalmacen", "@idagencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Comprobante_Aduana_U_S", parametros, valores, tipoParametro, 69).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.idregimen_aduana = IIf(dt.Rows(0).Item("idregimen_aduana") Is DBNull.Value, Nothing, dt.Rows(0).Item("idregimen_aduana"))
            d.iddespacho_aduana = IIf(dt.Rows(0).Item("iddespacho_aduana") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddespacho_aduana"))
            d.anio_dam = IIf(dt.Rows(0).Item("anio_dam") Is DBNull.Value, Nothing, dt.Rows(0).Item("anio_dam"))
            d.numero_dam = IIf(dt.Rows(0).Item("numero_dam") Is DBNull.Value, Nothing, dt.Rows(0).Item("numero_dam"))
            d.importe_fob = IIf(dt.Rows(0).Item("importe_fob") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_fob"))
            d.importe_seguro = IIf(dt.Rows(0).Item("importe_seguro") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_seguro"))
            d.importe_flete = IIf(dt.Rows(0).Item("importe_flete") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_flete"))
            d.importe_cif = IIf(dt.Rows(0).Item("importe_cif") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_cif"))
            d.importe_advalorem = IIf(dt.Rows(0).Item("importe_advalorem") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_advalorem"))
            d.importe_igv = IIf(dt.Rows(0).Item("importe_igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_igv"))
            d.importe_isc = IIf(dt.Rows(0).Item("importe_isc") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_isc"))
            d.importe_ipm = IIf(dt.Rows(0).Item("importe_ipm") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_ipm"))
            d.cantidad_bultos = IIf(dt.Rows(0).Item("cantidad_bultos") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad_bultos"))
            d.peso_bruto = IIf(dt.Rows(0).Item("peso_bruto") Is DBNull.Value, Nothing, dt.Rows(0).Item("peso_bruto"))
            d.fecha_llegada = IIf(dt.Rows(0).Item("fecha_llegada") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecha_llegada"))
            d.awb_bl = IIf(dt.Rows(0).Item("awb_bl") Is DBNull.Value, Nothing, dt.Rows(0).Item("awb_bl"))
            d.nave = IIf(dt.Rows(0).Item("nave") Is DBNull.Value, Nothing, dt.Rows(0).Item("nave"))
            d.mercancia = IIf(dt.Rows(0).Item("mercancia") Is DBNull.Value, Nothing, dt.Rows(0).Item("mercancia"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
        Else
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.idregimen_aduana = Nothing
            d.iddespacho_aduana = Nothing
            d.anio_dam = Nothing
            d.numero_dam = Nothing
            d.importe_fob = Nothing
            d.importe_seguro = Nothing
            d.importe_flete = Nothing
            d.importe_cif = Nothing
            d.importe_advalorem = Nothing
            d.importe_igv = Nothing
            d.importe_isc = Nothing
            d.importe_ipm = Nothing
            d.cantidad_bultos = Nothing
            d.peso_bruto = Nothing
            d.fecha_llegada = Nothing
            d.awb_bl = Nothing
            d.nave = Nothing
            d.mercancia = Nothing
            d.idalmacen = Nothing
            d.idagencia = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NComprobante_Aduana)
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento}
        sql.EjecutarProcedure("Str_Comprobante_Aduana_D", parametros, valores, tipoParametro, 3)
    End Sub
    Public Function Existe_Comprobante_Aduana(d As NComprobante_Aduana) As Boolean
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_Comprobante_Aduana", parametros, valores, tipoParametro, 3)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Comprobante_Aduana_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NComprobante_Aduana) As DataTable
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Comprobante_Aduana_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NComprobante_Aduana) As NComprobante_Aduana
        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Comprobante_Aduana_S", parametros, valores, tipoParametro, 3).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.idregimen_aduana = IIf(dt.Rows(0).Item("idregimen_aduana") Is DBNull.Value, Nothing, dt.Rows(0).Item("idregimen_aduana"))
            d.iddespacho_aduana = IIf(dt.Rows(0).Item("iddespacho_aduana") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddespacho_aduana"))
            d.anio_dam = IIf(dt.Rows(0).Item("anio_dam") Is DBNull.Value, Nothing, dt.Rows(0).Item("anio_dam"))
            d.numero_dam = IIf(dt.Rows(0).Item("numero_dam") Is DBNull.Value, Nothing, dt.Rows(0).Item("numero_dam"))
            d.importe_fob = IIf(dt.Rows(0).Item("importe_fob") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_fob"))
            d.importe_seguro = IIf(dt.Rows(0).Item("importe_seguro") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_seguro"))
            d.importe_flete = IIf(dt.Rows(0).Item("importe_flete") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_flete"))
            d.importe_cif = IIf(dt.Rows(0).Item("importe_cif") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_cif"))
            d.importe_advalorem = IIf(dt.Rows(0).Item("importe_advalorem") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_advalorem"))
            d.importe_igv = IIf(dt.Rows(0).Item("importe_igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_igv"))
            d.importe_isc = IIf(dt.Rows(0).Item("importe_isc") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_isc"))
            d.importe_ipm = IIf(dt.Rows(0).Item("importe_ipm") Is DBNull.Value, Nothing, dt.Rows(0).Item("importe_ipm"))
            d.cantidad_bultos = IIf(dt.Rows(0).Item("cantidad_bultos") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad_bultos"))
            d.peso_bruto = IIf(dt.Rows(0).Item("peso_bruto") Is DBNull.Value, Nothing, dt.Rows(0).Item("peso_bruto"))
            d.fecha_llegada = IIf(dt.Rows(0).Item("fecha_llegada") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecha_llegada"))
            d.awb_bl = IIf(dt.Rows(0).Item("awb_bl") Is DBNull.Value, Nothing, dt.Rows(0).Item("awb_bl"))
            d.nave = IIf(dt.Rows(0).Item("nave") Is DBNull.Value, Nothing, dt.Rows(0).Item("nave"))
            d.mercancia = IIf(dt.Rows(0).Item("mercancia") Is DBNull.Value, Nothing, dt.Rows(0).Item("mercancia"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
        Else
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.idregimen_aduana = Nothing
            d.iddespacho_aduana = Nothing
            d.anio_dam = Nothing
            d.numero_dam = Nothing
            d.importe_fob = Nothing
            d.importe_seguro = Nothing
            d.importe_flete = Nothing
            d.importe_cif = Nothing
            d.importe_advalorem = Nothing
            d.importe_igv = Nothing
            d.importe_isc = Nothing
            d.importe_ipm = Nothing
            d.cantidad_bultos = Nothing
            d.peso_bruto = Nothing
            d.fecha_llegada = Nothing
            d.awb_bl = Nothing
            d.nave = Nothing
            d.mercancia = Nothing
            d.idalmacen = Nothing
            d.idagencia = Nothing
        End If
        Return d
    End Function
#End Region


End Class
