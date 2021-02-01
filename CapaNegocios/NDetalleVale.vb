Imports CapaDatos
Public Class NDetalleVale
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idagencia As String
    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numerodocumento As String
    Public Property item As String
    Public Property idarticulo As String
    Public Property descripcion As String
    Public Property texto As String
    Public Property cantidad As Decimal
    Public Property unidad As String
    Public Property serie1 As String
    Public Property cantidad1 As Decimal
    Public Property unidadenvase As String
    Public Property numeroenvase As Decimal
    Public Property saldoentrega As Decimal
    Public Property precioventa As Decimal
    Public Property precioventah As Decimal
    Public Property precioventaimportacion As Decimal
    Public Property precioventaimportacionh As Decimal
    Public Property preciosigv As Decimal
    Public Property importedescuento As Decimal
    Public Property descuentodocumento As Decimal
    Public Property cargodistribucion As Decimal
    Public Property igv As Decimal
    Public Property importeigv As Decimal
    Public Property importeus As Decimal
    Public Property importemn As Decimal
    Public Property idtipoitemdescuento As String
    Public Property descuento1 As Decimal
    Public Property importedescuento1 As Decimal
    Public Property descuento2 As Decimal
    Public Property importedescuento2 As Decimal
    Public Property descuento3 As Decimal
    Public Property importedescuento3 As Decimal
    Public Property descuento4 As Decimal
    Public Property importedescuento4 As Decimal
    Public Property descuento5 As Decimal
    Public Property importedescuento5 As Decimal
    Public Property descuento6 As Decimal
    Public Property estado As String
    Public Property vendedor As String
    Public Property idalmacen As String
    Public Property numerocaja As String
    Public Property stock As String
    Public Property fechasdocumento As System.DateTime
    Public Property idlinea As String
    Public Property idcampania As String
    Public Property numeropaquete As String
    Public Property nrodescuentofinaciero As String
    Public Property nrodescuentolaboratorio As String
    Public Property nrodescuentoadicional As String
    Public Property nrodescuentobonificacion As String
    Public Property nrodescuentoflag As String
    Public Property comision As Decimal
    Public Property importecomision As Decimal
    Public Property usuariocrea As String
    Public Property fechacrea As System.DateTime
    Public Property preciounitarioorigen As Decimal
    Public Property idvendedor2 As String
    Public Property identrada As String
    Public Property npfacturado As String
    Public Property idlista As Integer
    Public Property loteserie As String
    Public Property lado As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NDetalleVale)
        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@item", "@idarticulo", "@descripcion", "@texto", "@cantidad", "@unidad", "@serie1", "@cantidad1", "@unidadenvase", "@numeroenvase", "@saldoentrega", "@precioventa", "@precioventah", "@precioventaimportacion", "@precioventaimportacionh", "@preciosigv", "@importedescuento", "@descuentodocumento", "@cargodistribucion", "@igv", "@importeigv", "@importeus", "@importemn", "@idtipoitemdescuento", "@descuento1", "@importedescuento1", "@descuento2", "@importedescuento2", "@descuento3", "@importedescuento3", "@descuento4", "@importedescuento4", "@descuento5", "@importedescuento5", "@descuento6", "@estado", "@vendedor", "@idalmacen", "@numerocaja", "@stock", "@fechasdocumento", "@idlinea", "@idcampania", "@numeropaquete", "@nrodescuentofinaciero", "@nrodescuentolaboratorio", "@nrodescuentoadicional", "@nrodescuentobonificacion", "@nrodescuentoflag", "@comision", "@importecomision", "@usuariocrea", "@fechacrea", "@preciounitarioorigen", "@idvendedor2", "@identrada", "@npfacturado", "@idlista", "@loteserie", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.idarticulo, d.descripcion, d.texto, d.cantidad, d.unidad, d.serie1, d.cantidad1, d.unidadenvase, d.numeroenvase, d.saldoentrega, d.precioventa, d.precioventah, d.precioventaimportacion, d.precioventaimportacionh, d.preciosigv, d.importedescuento, d.descuentodocumento, d.cargodistribucion, d.igv, d.importeigv, d.importeus, d.importemn, d.idtipoitemdescuento, d.descuento1, d.importedescuento1, d.descuento2, d.importedescuento2, d.descuento3, d.importedescuento3, d.descuento4, d.importedescuento4, d.descuento5, d.importedescuento5, d.descuento6, d.estado, d.vendedor, d.idalmacen, d.numerocaja, d.stock, d.fechasdocumento, d.idlinea, d.idcampania, d.numeropaquete, d.nrodescuentofinaciero, d.nrodescuentolaboratorio, d.nrodescuentoadicional, d.nrodescuentobonificacion, d.nrodescuentoflag, d.comision, d.importecomision, d.usuariocrea, d.fechacrea, d.preciounitarioorigen, d.idvendedor2, d.identrada, d.npfacturado, d.idlista, d.loteserie, d.lado}
        sql.EjecutarProcedure("Str_DetalleVale_I", parametros, valores, tipoParametro, 64)
    End Sub
    Public Sub Actualizar(d As NDetalleVale)
        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@item", "@idarticulo", "@descripcion", "@texto", "@cantidad", "@unidad", "@serie1", "@cantidad1", "@unidadenvase", "@numeroenvase", "@saldoentrega", "@precioventa", "@precioventah", "@precioventaimportacion", "@precioventaimportacionh", "@preciosigv", "@importedescuento", "@descuentodocumento", "@cargodistribucion", "@igv", "@importeigv", "@importeus", "@importemn", "@idtipoitemdescuento", "@descuento1", "@importedescuento1", "@descuento2", "@importedescuento2", "@descuento3", "@importedescuento3", "@descuento4", "@importedescuento4", "@descuento5", "@importedescuento5", "@descuento6", "@estado", "@vendedor", "@idalmacen", "@numerocaja", "@stock", "@fechasdocumento", "@idlinea", "@idcampania", "@numeropaquete", "@nrodescuentofinaciero", "@nrodescuentolaboratorio", "@nrodescuentoadicional", "@nrodescuentobonificacion", "@nrodescuentoflag", "@comision", "@importecomision", "@usuariocrea", "@fechacrea", "@preciounitarioorigen", "@idvendedor2", "@identrada", "@npfacturado", "@idlista", "@loteserie", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.idarticulo, d.descripcion, d.texto, d.cantidad, d.unidad, d.serie1, d.cantidad1, d.unidadenvase, d.numeroenvase, d.saldoentrega, d.precioventa, d.precioventah, d.precioventaimportacion, d.precioventaimportacionh, d.preciosigv, d.importedescuento, d.descuentodocumento, d.cargodistribucion, d.igv, d.importeigv, d.importeus, d.importemn, d.idtipoitemdescuento, d.descuento1, d.importedescuento1, d.descuento2, d.importedescuento2, d.descuento3, d.importedescuento3, d.descuento4, d.importedescuento4, d.descuento5, d.importedescuento5, d.descuento6, d.estado, d.vendedor, d.idalmacen, d.numerocaja, d.stock, d.fechasdocumento, d.idlinea, d.idcampania, d.numeropaquete, d.nrodescuentofinaciero, d.nrodescuentolaboratorio, d.nrodescuentoadicional, d.nrodescuentobonificacion, d.nrodescuentoflag, d.comision, d.importecomision, d.usuariocrea, d.fechacrea, d.preciounitarioorigen, d.idvendedor2, d.identrada, d.npfacturado, d.idlista, d.loteserie, d.lado}
        sql.EjecutarProcedure("Str_DetalleVale_U", parametros, valores, tipoParametro, 64)
    End Sub

    Public Function Agregar(d As NDetalleVale, Retornatable As Boolean) As NDetalleVale

        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@item", "@idarticulo", "@descripcion", "@texto", "@cantidad", "@unidad", "@serie1", "@cantidad1", "@unidadenvase", "@numeroenvase", "@saldoentrega", "@precioventa", "@precioventah", "@precioventaimportacion", "@precioventaimportacionh", "@preciosigv", "@importedescuento", "@descuentodocumento", "@cargodistribucion", "@igv", "@importeigv", "@importeus", "@importemn", "@idtipoitemdescuento", "@descuento1", "@importedescuento1", "@descuento2", "@importedescuento2", "@descuento3", "@importedescuento3", "@descuento4", "@importedescuento4", "@descuento5", "@importedescuento5", "@descuento6", "@estado", "@vendedor", "@idalmacen", "@numerocaja", "@stock", "@fechasdocumento", "@idlinea", "@idcampania", "@numeropaquete", "@nrodescuentofinaciero", "@nrodescuentolaboratorio", "@nrodescuentoadicional", "@nrodescuentobonificacion", "@nrodescuentoflag", "@comision", "@importecomision", "@usuariocrea", "@fechacrea", "@preciounitarioorigen", "@idvendedor2", "@identrada", "@npfacturado", "@idlista", "@loteserie", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.idarticulo, d.descripcion, d.texto, d.cantidad, d.unidad, d.serie1, d.cantidad1, d.unidadenvase, d.numeroenvase, d.saldoentrega, d.precioventa, d.precioventah, d.precioventaimportacion, d.precioventaimportacionh, d.preciosigv, d.importedescuento, d.descuentodocumento, d.cargodistribucion, d.igv, d.importeigv, d.importeus, d.importemn, d.idtipoitemdescuento, d.descuento1, d.importedescuento1, d.descuento2, d.importedescuento2, d.descuento3, d.importedescuento3, d.descuento4, d.importedescuento4, d.descuento5, d.importedescuento5, d.descuento6, d.estado, d.vendedor, d.idalmacen, d.numerocaja, d.stock, d.fechasdocumento, d.idlinea, d.idcampania, d.numeropaquete, d.nrodescuentofinaciero, d.nrodescuentolaboratorio, d.nrodescuentoadicional, d.nrodescuentobonificacion, d.nrodescuentoflag, d.comision, d.importecomision, d.usuariocrea, d.fechacrea, d.preciounitarioorigen, d.idvendedor2, d.identrada, d.npfacturado, d.idlista, d.loteserie, d.lado}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleVale_I_S", parametros, valores, tipoParametro, 64).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.texto = IIf(dt.Rows(0).Item("texto") Is DBNull.Value, Nothing, dt.Rows(0).Item("texto"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.unidad = IIf(dt.Rows(0).Item("unidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidad"))
            d.serie1 = IIf(dt.Rows(0).Item("serie1") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie1"))
            d.cantidad1 = IIf(dt.Rows(0).Item("cantidad1") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad1"))
            d.unidadenvase = IIf(dt.Rows(0).Item("unidadenvase") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidadenvase"))
            d.numeroenvase = IIf(dt.Rows(0).Item("numeroenvase") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroenvase"))
            d.saldoentrega = IIf(dt.Rows(0).Item("saldoentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldoentrega"))
            d.precioventa = IIf(dt.Rows(0).Item("precioventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventa"))
            d.precioventah = IIf(dt.Rows(0).Item("precioventah") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventah"))
            d.precioventaimportacion = IIf(dt.Rows(0).Item("precioventaimportacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventaimportacion"))
            d.precioventaimportacionh = IIf(dt.Rows(0).Item("precioventaimportacionh") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventaimportacionh"))
            d.preciosigv = IIf(dt.Rows(0).Item("preciosigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciosigv"))
            d.importedescuento = IIf(dt.Rows(0).Item("importedescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento"))
            d.descuentodocumento = IIf(dt.Rows(0).Item("descuentodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuentodocumento"))
            d.cargodistribucion = IIf(dt.Rows(0).Item("cargodistribucion") Is DBNull.Value, Nothing, dt.Rows(0).Item("cargodistribucion"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.idtipoitemdescuento = IIf(dt.Rows(0).Item("idtipoitemdescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipoitemdescuento"))
            d.descuento1 = IIf(dt.Rows(0).Item("descuento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento1"))
            d.importedescuento1 = IIf(dt.Rows(0).Item("importedescuento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento1"))
            d.descuento2 = IIf(dt.Rows(0).Item("descuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento2"))
            d.importedescuento2 = IIf(dt.Rows(0).Item("importedescuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento2"))
            d.descuento3 = IIf(dt.Rows(0).Item("descuento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento3"))
            d.importedescuento3 = IIf(dt.Rows(0).Item("importedescuento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento3"))
            d.descuento4 = IIf(dt.Rows(0).Item("descuento4") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento4"))
            d.importedescuento4 = IIf(dt.Rows(0).Item("importedescuento4") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento4"))
            d.descuento5 = IIf(dt.Rows(0).Item("descuento5") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento5"))
            d.importedescuento5 = IIf(dt.Rows(0).Item("importedescuento5") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento5"))
            d.descuento6 = IIf(dt.Rows(0).Item("descuento6") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento6"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.vendedor = IIf(dt.Rows(0).Item("vendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("vendedor"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.numerocaja = IIf(dt.Rows(0).Item("numerocaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerocaja"))
            d.stock = IIf(dt.Rows(0).Item("stock") Is DBNull.Value, Nothing, dt.Rows(0).Item("stock"))
            d.fechasdocumento = IIf(dt.Rows(0).Item("fechasdocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechasdocumento"))
            d.idlinea = IIf(dt.Rows(0).Item("idlinea") Is DBNull.Value, Nothing, dt.Rows(0).Item("idlinea"))
            d.idcampania = IIf(dt.Rows(0).Item("idcampania") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcampania"))
            d.numeropaquete = IIf(dt.Rows(0).Item("numeropaquete") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeropaquete"))
            d.nrodescuentofinaciero = IIf(dt.Rows(0).Item("nrodescuentofinaciero") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentofinaciero"))
            d.nrodescuentolaboratorio = IIf(dt.Rows(0).Item("nrodescuentolaboratorio") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentolaboratorio"))
            d.nrodescuentoadicional = IIf(dt.Rows(0).Item("nrodescuentoadicional") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentoadicional"))
            d.nrodescuentobonificacion = IIf(dt.Rows(0).Item("nrodescuentobonificacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentobonificacion"))
            d.nrodescuentoflag = IIf(dt.Rows(0).Item("nrodescuentoflag") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentoflag"))
            d.comision = IIf(dt.Rows(0).Item("comision") Is DBNull.Value, Nothing, dt.Rows(0).Item("comision"))
            d.importecomision = IIf(dt.Rows(0).Item("importecomision") Is DBNull.Value, Nothing, dt.Rows(0).Item("importecomision"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.preciounitarioorigen = IIf(dt.Rows(0).Item("preciounitarioorigen") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitarioorigen"))
            d.idvendedor2 = IIf(dt.Rows(0).Item("idvendedor2") Is DBNull.Value, Nothing, dt.Rows(0).Item("idvendedor2"))
            d.identrada = IIf(dt.Rows(0).Item("identrada") Is DBNull.Value, Nothing, dt.Rows(0).Item("identrada"))
            d.npfacturado = IIf(dt.Rows(0).Item("npfacturado") Is DBNull.Value, Nothing, dt.Rows(0).Item("npfacturado"))
            d.idlista = IIf(dt.Rows(0).Item("idlista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idlista"))
            d.loteserie = IIf(dt.Rows(0).Item("loteserie") Is DBNull.Value, Nothing, dt.Rows(0).Item("loteserie"))
            d.lado = IIf(dt.Rows(0).Item("lado") Is DBNull.Value, Nothing, dt.Rows(0).Item("lado"))
        Else
            d.idagencia = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
            d.texto = Nothing
            d.cantidad = Nothing
            d.unidad = Nothing
            d.serie1 = Nothing
            d.cantidad1 = Nothing
            d.unidadenvase = Nothing
            d.numeroenvase = Nothing
            d.saldoentrega = Nothing
            d.precioventa = Nothing
            d.precioventah = Nothing
            d.precioventaimportacion = Nothing
            d.precioventaimportacionh = Nothing
            d.preciosigv = Nothing
            d.importedescuento = Nothing
            d.descuentodocumento = Nothing
            d.cargodistribucion = Nothing
            d.igv = Nothing
            d.importeigv = Nothing
            d.importeus = Nothing
            d.importemn = Nothing
            d.idtipoitemdescuento = Nothing
            d.descuento1 = Nothing
            d.importedescuento1 = Nothing
            d.descuento2 = Nothing
            d.importedescuento2 = Nothing
            d.descuento3 = Nothing
            d.importedescuento3 = Nothing
            d.descuento4 = Nothing
            d.importedescuento4 = Nothing
            d.descuento5 = Nothing
            d.importedescuento5 = Nothing
            d.descuento6 = Nothing
            d.estado = Nothing
            d.vendedor = Nothing
            d.idalmacen = Nothing
            d.numerocaja = Nothing
            d.stock = Nothing
            d.fechasdocumento = Nothing
            d.idlinea = Nothing
            d.idcampania = Nothing
            d.numeropaquete = Nothing
            d.nrodescuentofinaciero = Nothing
            d.nrodescuentolaboratorio = Nothing
            d.nrodescuentoadicional = Nothing
            d.nrodescuentobonificacion = Nothing
            d.nrodescuentoflag = Nothing
            d.comision = Nothing
            d.importecomision = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.preciounitarioorigen = Nothing
            d.idvendedor2 = Nothing
            d.identrada = Nothing
            d.npfacturado = Nothing
            d.idlista = Nothing
            d.loteserie = Nothing
            d.lado = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NDetalleVale, Retornatable As Boolean) As NDetalleVale

        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@item", "@idarticulo", "@descripcion", "@texto", "@cantidad", "@unidad", "@serie1", "@cantidad1", "@unidadenvase", "@numeroenvase", "@saldoentrega", "@precioventa", "@precioventah", "@precioventaimportacion", "@precioventaimportacionh", "@preciosigv", "@importedescuento", "@descuentodocumento", "@cargodistribucion", "@igv", "@importeigv", "@importeus", "@importemn", "@idtipoitemdescuento", "@descuento1", "@importedescuento1", "@descuento2", "@importedescuento2", "@descuento3", "@importedescuento3", "@descuento4", "@importedescuento4", "@descuento5", "@importedescuento5", "@descuento6", "@estado", "@vendedor", "@idalmacen", "@numerocaja", "@stock", "@fechasdocumento", "@idlinea", "@idcampania", "@numeropaquete", "@nrodescuentofinaciero", "@nrodescuentolaboratorio", "@nrodescuentoadicional", "@nrodescuentobonificacion", "@nrodescuentoflag", "@comision", "@importecomision", "@usuariocrea", "@fechacrea", "@preciounitarioorigen", "@idvendedor2", "@identrada", "@npfacturado", "@idlista", "@loteserie", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.lado = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleVale_U_S", parametros, valores, tipoParametro, 64).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.texto = IIf(dt.Rows(0).Item("texto") Is DBNull.Value, Nothing, dt.Rows(0).Item("texto"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.unidad = IIf(dt.Rows(0).Item("unidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidad"))
            d.serie1 = IIf(dt.Rows(0).Item("serie1") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie1"))
            d.cantidad1 = IIf(dt.Rows(0).Item("cantidad1") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad1"))
            d.unidadenvase = IIf(dt.Rows(0).Item("unidadenvase") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidadenvase"))
            d.numeroenvase = IIf(dt.Rows(0).Item("numeroenvase") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroenvase"))
            d.saldoentrega = IIf(dt.Rows(0).Item("saldoentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldoentrega"))
            d.precioventa = IIf(dt.Rows(0).Item("precioventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventa"))
            d.precioventah = IIf(dt.Rows(0).Item("precioventah") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventah"))
            d.precioventaimportacion = IIf(dt.Rows(0).Item("precioventaimportacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventaimportacion"))
            d.precioventaimportacionh = IIf(dt.Rows(0).Item("precioventaimportacionh") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventaimportacionh"))
            d.preciosigv = IIf(dt.Rows(0).Item("preciosigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciosigv"))
            d.importedescuento = IIf(dt.Rows(0).Item("importedescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento"))
            d.descuentodocumento = IIf(dt.Rows(0).Item("descuentodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuentodocumento"))
            d.cargodistribucion = IIf(dt.Rows(0).Item("cargodistribucion") Is DBNull.Value, Nothing, dt.Rows(0).Item("cargodistribucion"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.idtipoitemdescuento = IIf(dt.Rows(0).Item("idtipoitemdescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipoitemdescuento"))
            d.descuento1 = IIf(dt.Rows(0).Item("descuento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento1"))
            d.importedescuento1 = IIf(dt.Rows(0).Item("importedescuento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento1"))
            d.descuento2 = IIf(dt.Rows(0).Item("descuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento2"))
            d.importedescuento2 = IIf(dt.Rows(0).Item("importedescuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento2"))
            d.descuento3 = IIf(dt.Rows(0).Item("descuento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento3"))
            d.importedescuento3 = IIf(dt.Rows(0).Item("importedescuento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento3"))
            d.descuento4 = IIf(dt.Rows(0).Item("descuento4") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento4"))
            d.importedescuento4 = IIf(dt.Rows(0).Item("importedescuento4") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento4"))
            d.descuento5 = IIf(dt.Rows(0).Item("descuento5") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento5"))
            d.importedescuento5 = IIf(dt.Rows(0).Item("importedescuento5") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento5"))
            d.descuento6 = IIf(dt.Rows(0).Item("descuento6") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento6"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.vendedor = IIf(dt.Rows(0).Item("vendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("vendedor"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.numerocaja = IIf(dt.Rows(0).Item("numerocaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerocaja"))
            d.stock = IIf(dt.Rows(0).Item("stock") Is DBNull.Value, Nothing, dt.Rows(0).Item("stock"))
            d.fechasdocumento = IIf(dt.Rows(0).Item("fechasdocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechasdocumento"))
            d.idlinea = IIf(dt.Rows(0).Item("idlinea") Is DBNull.Value, Nothing, dt.Rows(0).Item("idlinea"))
            d.idcampania = IIf(dt.Rows(0).Item("idcampania") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcampania"))
            d.numeropaquete = IIf(dt.Rows(0).Item("numeropaquete") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeropaquete"))
            d.nrodescuentofinaciero = IIf(dt.Rows(0).Item("nrodescuentofinaciero") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentofinaciero"))
            d.nrodescuentolaboratorio = IIf(dt.Rows(0).Item("nrodescuentolaboratorio") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentolaboratorio"))
            d.nrodescuentoadicional = IIf(dt.Rows(0).Item("nrodescuentoadicional") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentoadicional"))
            d.nrodescuentobonificacion = IIf(dt.Rows(0).Item("nrodescuentobonificacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentobonificacion"))
            d.nrodescuentoflag = IIf(dt.Rows(0).Item("nrodescuentoflag") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentoflag"))
            d.comision = IIf(dt.Rows(0).Item("comision") Is DBNull.Value, Nothing, dt.Rows(0).Item("comision"))
            d.importecomision = IIf(dt.Rows(0).Item("importecomision") Is DBNull.Value, Nothing, dt.Rows(0).Item("importecomision"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.preciounitarioorigen = IIf(dt.Rows(0).Item("preciounitarioorigen") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitarioorigen"))
            d.idvendedor2 = IIf(dt.Rows(0).Item("idvendedor2") Is DBNull.Value, Nothing, dt.Rows(0).Item("idvendedor2"))
            d.identrada = IIf(dt.Rows(0).Item("identrada") Is DBNull.Value, Nothing, dt.Rows(0).Item("identrada"))
            d.npfacturado = IIf(dt.Rows(0).Item("npfacturado") Is DBNull.Value, Nothing, dt.Rows(0).Item("npfacturado"))
            d.idlista = IIf(dt.Rows(0).Item("idlista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idlista"))
            d.loteserie = IIf(dt.Rows(0).Item("loteserie") Is DBNull.Value, Nothing, dt.Rows(0).Item("loteserie"))
            d.lado = IIf(dt.Rows(0).Item("lado") Is DBNull.Value, Nothing, dt.Rows(0).Item("lado"))
        Else
            d.idagencia = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
            d.texto = Nothing
            d.cantidad = Nothing
            d.unidad = Nothing
            d.serie1 = Nothing
            d.cantidad1 = Nothing
            d.unidadenvase = Nothing
            d.numeroenvase = Nothing
            d.saldoentrega = Nothing
            d.precioventa = Nothing
            d.precioventah = Nothing
            d.precioventaimportacion = Nothing
            d.precioventaimportacionh = Nothing
            d.preciosigv = Nothing
            d.importedescuento = Nothing
            d.descuentodocumento = Nothing
            d.cargodistribucion = Nothing
            d.igv = Nothing
            d.importeigv = Nothing
            d.importeus = Nothing
            d.importemn = Nothing
            d.idtipoitemdescuento = Nothing
            d.descuento1 = Nothing
            d.importedescuento1 = Nothing
            d.descuento2 = Nothing
            d.importedescuento2 = Nothing
            d.descuento3 = Nothing
            d.importedescuento3 = Nothing
            d.descuento4 = Nothing
            d.importedescuento4 = Nothing
            d.descuento5 = Nothing
            d.importedescuento5 = Nothing
            d.descuento6 = Nothing
            d.estado = Nothing
            d.vendedor = Nothing
            d.idalmacen = Nothing
            d.numerocaja = Nothing
            d.stock = Nothing
            d.fechasdocumento = Nothing
            d.idlinea = Nothing
            d.idcampania = Nothing
            d.numeropaquete = Nothing
            d.nrodescuentofinaciero = Nothing
            d.nrodescuentolaboratorio = Nothing
            d.nrodescuentoadicional = Nothing
            d.nrodescuentobonificacion = Nothing
            d.nrodescuentoflag = Nothing
            d.comision = Nothing
            d.importecomision = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.preciounitarioorigen = Nothing
            d.idvendedor2 = Nothing
            d.identrada = Nothing
            d.npfacturado = Nothing
            d.idlista = Nothing
            d.loteserie = Nothing
            d.lado = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NDetalleVale)
        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@item", "@idarticulo", "@idalmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.idarticulo, d.idalmacen}
        sql.EjecutarProcedure("Str_DetalleVale_D", parametros, valores, tipoParametro, 7)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@item", "@idarticulo", "@idalmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleVale_S", parametros, valores, tipoParametro, 7).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NDetalleVale) As DataTable
        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@item", "@idarticulo", "@idalmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.idarticulo, d.idalmacen}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleVale_S", parametros, valores, tipoParametro, 7).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDetalleVale) As NDetalleVale
        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@item", "@idarticulo", "@idalmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.item, d.idarticulo, d.idalmacen}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_DetalleVale_S", parametros, valores, tipoParametro, 7).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.texto = IIf(dt.Rows(0).Item("texto") Is DBNull.Value, Nothing, dt.Rows(0).Item("texto"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.unidad = IIf(dt.Rows(0).Item("unidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidad"))
            d.serie1 = IIf(dt.Rows(0).Item("serie1") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie1"))
            d.cantidad1 = IIf(dt.Rows(0).Item("cantidad1") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad1"))
            d.unidadenvase = IIf(dt.Rows(0).Item("unidadenvase") Is DBNull.Value, Nothing, dt.Rows(0).Item("unidadenvase"))
            d.numeroenvase = IIf(dt.Rows(0).Item("numeroenvase") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroenvase"))
            d.saldoentrega = IIf(dt.Rows(0).Item("saldoentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldoentrega"))
            d.precioventa = IIf(dt.Rows(0).Item("precioventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventa"))
            d.precioventah = IIf(dt.Rows(0).Item("precioventah") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventah"))
            d.precioventaimportacion = IIf(dt.Rows(0).Item("precioventaimportacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventaimportacion"))
            d.precioventaimportacionh = IIf(dt.Rows(0).Item("precioventaimportacionh") Is DBNull.Value, Nothing, dt.Rows(0).Item("precioventaimportacionh"))
            d.preciosigv = IIf(dt.Rows(0).Item("preciosigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciosigv"))
            d.importedescuento = IIf(dt.Rows(0).Item("importedescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento"))
            d.descuentodocumento = IIf(dt.Rows(0).Item("descuentodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuentodocumento"))
            d.cargodistribucion = IIf(dt.Rows(0).Item("cargodistribucion") Is DBNull.Value, Nothing, dt.Rows(0).Item("cargodistribucion"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.importeus = IIf(dt.Rows(0).Item("importeus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeus"))
            d.importemn = IIf(dt.Rows(0).Item("importemn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importemn"))
            d.idtipoitemdescuento = IIf(dt.Rows(0).Item("idtipoitemdescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipoitemdescuento"))
            d.descuento1 = IIf(dt.Rows(0).Item("descuento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento1"))
            d.importedescuento1 = IIf(dt.Rows(0).Item("importedescuento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento1"))
            d.descuento2 = IIf(dt.Rows(0).Item("descuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento2"))
            d.importedescuento2 = IIf(dt.Rows(0).Item("importedescuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento2"))
            d.descuento3 = IIf(dt.Rows(0).Item("descuento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento3"))
            d.importedescuento3 = IIf(dt.Rows(0).Item("importedescuento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento3"))
            d.descuento4 = IIf(dt.Rows(0).Item("descuento4") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento4"))
            d.importedescuento4 = IIf(dt.Rows(0).Item("importedescuento4") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento4"))
            d.descuento5 = IIf(dt.Rows(0).Item("descuento5") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento5"))
            d.importedescuento5 = IIf(dt.Rows(0).Item("importedescuento5") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento5"))
            d.descuento6 = IIf(dt.Rows(0).Item("descuento6") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento6"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.vendedor = IIf(dt.Rows(0).Item("vendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("vendedor"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.numerocaja = IIf(dt.Rows(0).Item("numerocaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerocaja"))
            d.stock = IIf(dt.Rows(0).Item("stock") Is DBNull.Value, Nothing, dt.Rows(0).Item("stock"))
            d.fechasdocumento = IIf(dt.Rows(0).Item("fechasdocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechasdocumento"))
            d.idlinea = IIf(dt.Rows(0).Item("idlinea") Is DBNull.Value, Nothing, dt.Rows(0).Item("idlinea"))
            d.idcampania = IIf(dt.Rows(0).Item("idcampania") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcampania"))
            d.numeropaquete = IIf(dt.Rows(0).Item("numeropaquete") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeropaquete"))
            d.nrodescuentofinaciero = IIf(dt.Rows(0).Item("nrodescuentofinaciero") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentofinaciero"))
            d.nrodescuentolaboratorio = IIf(dt.Rows(0).Item("nrodescuentolaboratorio") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentolaboratorio"))
            d.nrodescuentoadicional = IIf(dt.Rows(0).Item("nrodescuentoadicional") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentoadicional"))
            d.nrodescuentobonificacion = IIf(dt.Rows(0).Item("nrodescuentobonificacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentobonificacion"))
            d.nrodescuentoflag = IIf(dt.Rows(0).Item("nrodescuentoflag") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodescuentoflag"))
            d.comision = IIf(dt.Rows(0).Item("comision") Is DBNull.Value, Nothing, dt.Rows(0).Item("comision"))
            d.importecomision = IIf(dt.Rows(0).Item("importecomision") Is DBNull.Value, Nothing, dt.Rows(0).Item("importecomision"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.preciounitarioorigen = IIf(dt.Rows(0).Item("preciounitarioorigen") Is DBNull.Value, Nothing, dt.Rows(0).Item("preciounitarioorigen"))
            d.idvendedor2 = IIf(dt.Rows(0).Item("idvendedor2") Is DBNull.Value, Nothing, dt.Rows(0).Item("idvendedor2"))
            d.identrada = IIf(dt.Rows(0).Item("identrada") Is DBNull.Value, Nothing, dt.Rows(0).Item("identrada"))
            d.npfacturado = IIf(dt.Rows(0).Item("npfacturado") Is DBNull.Value, Nothing, dt.Rows(0).Item("npfacturado"))
            d.idlista = IIf(dt.Rows(0).Item("idlista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idlista"))
            d.loteserie = IIf(dt.Rows(0).Item("loteserie") Is DBNull.Value, Nothing, dt.Rows(0).Item("loteserie"))
            d.lado = IIf(dt.Rows(0).Item("lado") Is DBNull.Value, Nothing, dt.Rows(0).Item("lado"))
        Else
            d.idagencia = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.descripcion = Nothing
            d.texto = Nothing
            d.cantidad = Nothing
            d.unidad = Nothing
            d.serie1 = Nothing
            d.cantidad1 = Nothing
            d.unidadenvase = Nothing
            d.numeroenvase = Nothing
            d.saldoentrega = Nothing
            d.precioventa = Nothing
            d.precioventah = Nothing
            d.precioventaimportacion = Nothing
            d.precioventaimportacionh = Nothing
            d.preciosigv = Nothing
            d.importedescuento = Nothing
            d.descuentodocumento = Nothing
            d.cargodistribucion = Nothing
            d.igv = Nothing
            d.importeigv = Nothing
            d.importeus = Nothing
            d.importemn = Nothing
            d.idtipoitemdescuento = Nothing
            d.descuento1 = Nothing
            d.importedescuento1 = Nothing
            d.descuento2 = Nothing
            d.importedescuento2 = Nothing
            d.descuento3 = Nothing
            d.importedescuento3 = Nothing
            d.descuento4 = Nothing
            d.importedescuento4 = Nothing
            d.descuento5 = Nothing
            d.importedescuento5 = Nothing
            d.descuento6 = Nothing
            d.estado = Nothing
            d.vendedor = Nothing
            d.idalmacen = Nothing
            d.numerocaja = Nothing
            d.stock = Nothing
            d.fechasdocumento = Nothing
            d.idlinea = Nothing
            d.idcampania = Nothing
            d.numeropaquete = Nothing
            d.nrodescuentofinaciero = Nothing
            d.nrodescuentolaboratorio = Nothing
            d.nrodescuentoadicional = Nothing
            d.nrodescuentobonificacion = Nothing
            d.nrodescuentoflag = Nothing
            d.comision = Nothing
            d.importecomision = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.preciounitarioorigen = Nothing
            d.idvendedor2 = Nothing
            d.identrada = Nothing
            d.npfacturado = Nothing
            d.idlista = Nothing
            d.loteserie = Nothing
            d.lado = Nothing
        End If
        Return d
    End Function
#End Region

End Class
