Imports CapaDatos

Public Class NVales
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idagencia As String
    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numerodocumento As String
    Public Property numeropedido As String
    Public Property fechadocumento As System.DateTime
    Public Property fechavencimineto As System.DateTime
    Public Property debehaber As String
    Public Property idvendedor As String
    Public Property idcaja As String
    Public Property idcliente As String
    Public Property nombrecliente As String
    Public Property direccion As String
    Public Property ruc As String
    Public Property idalmacen As String
    Public Property idformaventa As String
    Public Property idmoneda As String
    Public Property tipocambio As Decimal
    Public Property importetotal As Decimal
    Public Property importeigv As Decimal
    Public Property saldo As Decimal
    Public Property importedescuento As Decimal
    Public Property numeroorden As String
    Public Property idtipodocumento1 As String
    Public Property serie1 As String
    Public Property numerodocumento1 As String
    Public Property descripcion As String
    Public Property estado As String
    Public Property facturaguia As String
    Public Property idtransportista As String
    Public Property idcentrocosto As String
    Public Property idmaquina As String
    Public Property destino As String
    Public Property idtipofactura As String
    Public Property idtipoanexo As String
    Public Property idanexo As String
    Public Property descuneto1 As Decimal
    Public Property descuento2 As Decimal
    Public Property flete As Decimal
    Public Property embalaje As Decimal
    Public Property tasa As Decimal
    Public Property idusuariooperador As String
    Public Property idusuariosectorista As String
    Public Property idcadena As String
    Public Property idinternocadena As String
    Public Property idautorizacion As String
    Public Property reparto As String
    Public Property usuariocrea As String
    Public Property fechacrea As System.DateTime
    Public Property usuariomod As String
    Public Property fechamod As System.DateTime
    Public Property idtiponotacredito As String
    Public Property linea As String
    Public Property impreso As String
    Public Property anuladonc As String
    Public Property idvendedor1 As String
    Public Property igv As Decimal
    Public Property idchofer As String
    Public Property idzonaventa As String
    Public Property idtipodocumento2 As String
    Public Property numerodocumento2 As String
    Public Property idsubdiario As String
    Public Property nrocontable As String
    Public Property importetotalus As Decimal
    Public Property importetotalmn As Decimal
    Public Property importeigvus As Decimal
    Public Property importeigvmn As Decimal
    Public Property estadosunat As String
    Public Property codigohas As String
    Public Property barrapdf417 As Byte()
    Public Property signaturevalue As String
    Public Property tipooperacion As String
    Public Property tipoafecigv As String
    Public Property islote As String
    Public Property valorventa As Decimal
    Public Property idturno As String
    Public Property isliq As Boolean
    Public Property fechadocumento2 As System.DateTime

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NVales)

        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@numeropedido", "@fechadocumento", "@fechavencimineto", "@debehaber", "@idvendedor", "@idcaja", "@idcliente", "@nombrecliente", "@direccion", "@ruc", "@idalmacen", "@idformaventa", "@idmoneda", "@tipocambio", "@importetotal", "@importeigv", "@saldo", "@importedescuento", "@numeroorden", "@idtipodocumento1", "@serie1", "@numerodocumento1", "@descripcion", "@estado", "@facturaguia", "@idtransportista", "@idcentrocosto", "@idmaquina", "@destino", "@idtipofactura", "@idtipoanexo", "@idanexo", "@descuneto1", "@descuento2", "@flete", "@embalaje", "@tasa", "@idusuariooperador", "@idusuariosectorista", "@idcadena", "@idinternocadena", "@idautorizacion", "@reparto", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@idtiponotacredito", "@linea", "@impreso", "@anuladonc", "@idvendedor1", "@igv", "@idchofer", "@idzonaventa", "@idtipodocumento2", "@numerodocumento2", "@idsubdiario", "@nrocontable", "@importetotalus", "@importetotalmn", "@importeigvus", "@importeigvmn", "@estadosunat", "@codigohas", "@barrapdf417", "@signaturevalue", "@tipooperacion", "@tipoafecigv", "@islote", "@valorventa", "@idturno", "@isliq", "@fechadocumento2"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.numeropedido, d.fechadocumento, d.fechavencimineto, d.debehaber, d.idvendedor, d.idcaja, d.idcliente, d.nombrecliente, d.direccion, d.ruc, d.idalmacen, d.idformaventa, d.idmoneda, d.tipocambio, d.importetotal, d.importeigv, d.saldo, d.importedescuento, d.numeroorden, d.idtipodocumento1, d.serie1, d.numerodocumento1, d.descripcion, d.estado, d.facturaguia, d.idtransportista, d.idcentrocosto, d.idmaquina, d.destino, d.idtipofactura, d.idtipoanexo, d.idanexo, d.descuneto1, d.descuento2, d.flete, d.embalaje, d.tasa, d.idusuariooperador, d.idusuariosectorista, d.idcadena, d.idinternocadena, d.idautorizacion, d.reparto, d.usuariocrea, d.fechacrea, d.usuariomod, d.fechamod, d.idtiponotacredito, d.linea, d.impreso, d.anuladonc, d.idvendedor1, d.igv, d.idchofer, d.idzonaventa, d.idtipodocumento2, d.numerodocumento2, d.idsubdiario, d.nrocontable, d.importetotalus, d.importetotalmn, d.importeigvus, d.importeigvmn, d.estadosunat, d.codigohas, d.barrapdf417, d.signaturevalue, d.tipooperacion, d.tipoafecigv, d.islote, d.valorventa, d.idturno, d.isliq, d.fechadocumento2}
        sql.EjecutarProcedure("Str_Vales_I", parametros, valores, tipoParametro, 78)
    End Sub
    Public Sub Actualizar(d As NVales)
        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@numeropedido", "@fechadocumento", "@fechavencimineto", "@debehaber", "@idvendedor", "@idcaja", "@idcliente", "@nombrecliente", "@direccion", "@ruc", "@idalmacen", "@idformaventa", "@idmoneda", "@tipocambio", "@importetotal", "@importeigv", "@saldo", "@importedescuento", "@numeroorden", "@idtipodocumento1", "@serie1", "@numerodocumento1", "@descripcion", "@estado", "@facturaguia", "@idtransportista", "@idcentrocosto", "@idmaquina", "@destino", "@idtipofactura", "@idtipoanexo", "@idanexo", "@descuneto1", "@descuento2", "@flete", "@embalaje", "@tasa", "@idusuariooperador", "@idusuariosectorista", "@idcadena", "@idinternocadena", "@idautorizacion", "@reparto", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@idtiponotacredito", "@linea", "@impreso", "@anuladonc", "@idvendedor1", "@igv", "@idchofer", "@idzonaventa", "@idtipodocumento2", "@numerodocumento2", "@idsubdiario", "@nrocontable", "@importetotalus", "@importetotalmn", "@importeigvus", "@importeigvmn", "@estadosunat", "@codigohas", "@barrapdf417", "@signaturevalue", "@tipooperacion", "@tipoafecigv", "@islote", "@valorventa", "@idturno", "@isliq", "@fechadocumento2"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.numeropedido, d.fechadocumento, d.fechavencimineto, d.debehaber, d.idvendedor, d.idcaja, d.idcliente, d.nombrecliente, d.direccion, d.ruc, d.idalmacen, d.idformaventa, d.idmoneda, d.tipocambio, d.importetotal, d.importeigv, d.saldo, d.importedescuento, d.numeroorden, d.idtipodocumento1, d.serie1, d.numerodocumento1, d.descripcion, d.estado, d.facturaguia, d.idtransportista, d.idcentrocosto, d.idmaquina, d.destino, d.idtipofactura, d.idtipoanexo, d.idanexo, d.descuneto1, d.descuento2, d.flete, d.embalaje, d.tasa, d.idusuariooperador, d.idusuariosectorista, d.idcadena, d.idinternocadena, d.idautorizacion, d.reparto, d.usuariocrea, d.fechacrea, d.usuariomod, d.fechamod, d.idtiponotacredito, d.linea, d.impreso, d.anuladonc, d.idvendedor1, d.igv, d.idchofer, d.idzonaventa, d.idtipodocumento2, d.numerodocumento2, d.idsubdiario, d.nrocontable, d.importetotalus, d.importetotalmn, d.importeigvus, d.importeigvmn, d.estadosunat, d.codigohas, d.barrapdf417, d.signaturevalue, d.tipooperacion, d.tipoafecigv, d.islote, d.valorventa, d.idturno, d.isliq, d.fechadocumento2}
        sql.EjecutarProcedure("Str_Vales_U", parametros, valores, tipoParametro, 78)
    End Sub
    Public Function Agregar(d As NVales, Retornatable As Boolean) As NVales

        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@numeropedido", "@fechadocumento", "@fechavencimineto", "@debehaber", "@idvendedor", "@idcaja", "@idcliente", "@nombrecliente", "@direccion", "@ruc", "@idalmacen", "@idformaventa", "@idmoneda", "@tipocambio", "@importetotal", "@importeigv", "@saldo", "@importedescuento", "@numeroorden", "@idtipodocumento1", "@serie1", "@numerodocumento1", "@descripcion", "@estado", "@facturaguia", "@idtransportista", "@idcentrocosto", "@idmaquina", "@destino", "@idtipofactura", "@idtipoanexo", "@idanexo", "@descuneto1", "@descuento2", "@flete", "@embalaje", "@tasa", "@idusuariooperador", "@idusuariosectorista", "@idcadena", "@idinternocadena", "@idautorizacion", "@reparto", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@idtiponotacredito", "@linea", "@impreso", "@anuladonc", "@idvendedor1", "@igv", "@idchofer", "@idzonaventa", "@idtipodocumento2", "@numerodocumento2", "@idsubdiario", "@nrocontable", "@importetotalus", "@importetotalmn", "@importeigvus", "@importeigvmn", "@estadosunat", "@codigohas", "@barrapdf417", "@signaturevalue", "@tipooperacion", "@tipoafecigv", "@islote", "@valorventa", "@idturno", "@isliq", "@fechadocumento2"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.numeropedido, d.fechadocumento, d.fechavencimineto, d.debehaber, d.idvendedor, d.idcaja, d.idcliente, d.nombrecliente, d.direccion, d.ruc, d.idalmacen, d.idformaventa, d.idmoneda, d.tipocambio, d.importetotal, d.importeigv, d.saldo, d.importedescuento, d.numeroorden, d.idtipodocumento1, d.serie1, d.numerodocumento1, d.descripcion, d.estado, d.facturaguia, d.idtransportista, d.idcentrocosto, d.idmaquina, d.destino, d.idtipofactura, d.idtipoanexo, d.idanexo, d.descuneto1, d.descuento2, d.flete, d.embalaje, d.tasa, d.idusuariooperador, d.idusuariosectorista, d.idcadena, d.idinternocadena, d.idautorizacion, d.reparto, d.usuariocrea, d.fechacrea, d.usuariomod, d.fechamod, d.idtiponotacredito, d.linea, d.impreso, d.anuladonc, d.idvendedor1, d.igv, d.idchofer, d.idzonaventa, d.idtipodocumento2, d.numerodocumento2, d.idsubdiario, d.nrocontable, d.importetotalus, d.importetotalmn, d.importeigvus, d.importeigvmn, d.estadosunat, d.codigohas, d.barrapdf417, d.signaturevalue, d.tipooperacion, d.tipoafecigv, d.islote, d.valorventa, d.idturno, d.isliq, d.fechadocumento2}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Vales_I_S", parametros, valores, tipoParametro, 78).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.numeropedido = IIf(dt.Rows(0).Item("numeropedido") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeropedido"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.fechavencimineto = IIf(dt.Rows(0).Item("fechavencimineto") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechavencimineto"))
            d.debehaber = IIf(dt.Rows(0).Item("debehaber") Is DBNull.Value, Nothing, dt.Rows(0).Item("debehaber"))
            d.idvendedor = IIf(dt.Rows(0).Item("idvendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idvendedor"))
            d.idcaja = IIf(dt.Rows(0).Item("idcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcaja"))
            d.idcliente = IIf(dt.Rows(0).Item("idcliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcliente"))
            d.nombrecliente = IIf(dt.Rows(0).Item("nombrecliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecliente"))
            d.direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.ruc = IIf(dt.Rows(0).Item("ruc") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruc"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idformaventa = IIf(dt.Rows(0).Item("idformaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idformaventa"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.importetotal = IIf(dt.Rows(0).Item("importetotal") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotal"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.importedescuento = IIf(dt.Rows(0).Item("importedescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento"))
            d.numeroorden = IIf(dt.Rows(0).Item("numeroorden") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroorden"))
            d.idtipodocumento1 = IIf(dt.Rows(0).Item("idtipodocumento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento1"))
            d.serie1 = IIf(dt.Rows(0).Item("serie1") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie1"))
            d.numerodocumento1 = IIf(dt.Rows(0).Item("numerodocumento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento1"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.facturaguia = IIf(dt.Rows(0).Item("facturaguia") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaguia"))
            d.idtransportista = IIf(dt.Rows(0).Item("idtransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtransportista"))
            d.idcentrocosto = IIf(dt.Rows(0).Item("idcentrocosto") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcentrocosto"))
            d.idmaquina = IIf(dt.Rows(0).Item("idmaquina") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmaquina"))
            d.destino = IIf(dt.Rows(0).Item("destino") Is DBNull.Value, Nothing, dt.Rows(0).Item("destino"))
            d.idtipofactura = IIf(dt.Rows(0).Item("idtipofactura") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipofactura"))
            d.idtipoanexo = IIf(dt.Rows(0).Item("idtipoanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipoanexo"))
            d.idanexo = IIf(dt.Rows(0).Item("idanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idanexo"))
            d.descuneto1 = IIf(dt.Rows(0).Item("descuneto1") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuneto1"))
            d.descuento2 = IIf(dt.Rows(0).Item("descuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento2"))
            d.flete = IIf(dt.Rows(0).Item("flete") Is DBNull.Value, Nothing, dt.Rows(0).Item("flete"))
            d.embalaje = IIf(dt.Rows(0).Item("embalaje") Is DBNull.Value, Nothing, dt.Rows(0).Item("embalaje"))
            d.tasa = IIf(dt.Rows(0).Item("tasa") Is DBNull.Value, Nothing, dt.Rows(0).Item("tasa"))
            d.idusuariooperador = IIf(dt.Rows(0).Item("idusuariooperador") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariooperador"))
            d.idusuariosectorista = IIf(dt.Rows(0).Item("idusuariosectorista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariosectorista"))
            d.idcadena = IIf(dt.Rows(0).Item("idcadena") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcadena"))
            d.idinternocadena = IIf(dt.Rows(0).Item("idinternocadena") Is DBNull.Value, Nothing, dt.Rows(0).Item("idinternocadena"))
            d.idautorizacion = IIf(dt.Rows(0).Item("idautorizacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("idautorizacion"))
            d.reparto = IIf(dt.Rows(0).Item("reparto") Is DBNull.Value, Nothing, dt.Rows(0).Item("reparto"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.idtiponotacredito = IIf(dt.Rows(0).Item("idtiponotacredito") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtiponotacredito"))
            d.linea = IIf(dt.Rows(0).Item("linea") Is DBNull.Value, Nothing, dt.Rows(0).Item("linea"))
            d.impreso = IIf(dt.Rows(0).Item("impreso") Is DBNull.Value, Nothing, dt.Rows(0).Item("impreso"))
            d.anuladonc = IIf(dt.Rows(0).Item("anuladonc") Is DBNull.Value, Nothing, dt.Rows(0).Item("anuladonc"))
            d.idvendedor1 = IIf(dt.Rows(0).Item("idvendedor1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idvendedor1"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.idchofer = IIf(dt.Rows(0).Item("idchofer") Is DBNull.Value, Nothing, dt.Rows(0).Item("idchofer"))
            d.idzonaventa = IIf(dt.Rows(0).Item("idzonaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idzonaventa"))
            d.idtipodocumento2 = IIf(dt.Rows(0).Item("idtipodocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento2"))
            d.numerodocumento2 = IIf(dt.Rows(0).Item("numerodocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento2"))
            d.idsubdiario = IIf(dt.Rows(0).Item("idsubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsubdiario"))
            d.nrocontable = IIf(dt.Rows(0).Item("nrocontable") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrocontable"))
            d.importetotalus = IIf(dt.Rows(0).Item("importetotalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalus"))
            d.importetotalmn = IIf(dt.Rows(0).Item("importetotalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalmn"))
            d.importeigvus = IIf(dt.Rows(0).Item("importeigvus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvus"))
            d.importeigvmn = IIf(dt.Rows(0).Item("importeigvmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvmn"))
            d.estadosunat = IIf(dt.Rows(0).Item("estadosunat") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadosunat"))
            d.codigohas = IIf(dt.Rows(0).Item("codigohas") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigohas"))
            d.barrapdf417 = IIf(dt.Rows(0).Item("barrapdf417") Is DBNull.Value, Nothing, dt.Rows(0).Item("barrapdf417"))
            d.signaturevalue = IIf(dt.Rows(0).Item("signaturevalue") Is DBNull.Value, Nothing, dt.Rows(0).Item("signaturevalue"))
            d.tipooperacion = IIf(dt.Rows(0).Item("tipooperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipooperacion"))
            d.tipoafecigv = IIf(dt.Rows(0).Item("tipoafecigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoafecigv"))
            d.islote = IIf(dt.Rows(0).Item("islote") Is DBNull.Value, Nothing, dt.Rows(0).Item("islote"))
            d.valorventa = IIf(dt.Rows(0).Item("valorventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventa"))
            d.idturno = IIf(dt.Rows(0).Item("idturno") Is DBNull.Value, Nothing, dt.Rows(0).Item("idturno"))
            d.isliq = IIf(dt.Rows(0).Item("isliq") Is DBNull.Value, Nothing, dt.Rows(0).Item("isliq"))
            d.fechadocumento2 = IIf(dt.Rows(0).Item("fechadocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento2"))
        Else
            d.idagencia = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.numeropedido = Nothing
            d.fechadocumento = Nothing
            d.fechavencimineto = Nothing
            d.debehaber = Nothing
            d.idvendedor = Nothing
            d.idcaja = Nothing
            d.idcliente = Nothing
            d.nombrecliente = Nothing
            d.direccion = Nothing
            d.ruc = Nothing
            d.idalmacen = Nothing
            d.idformaventa = Nothing
            d.idmoneda = Nothing
            d.tipocambio = Nothing
            d.importetotal = Nothing
            d.importeigv = Nothing
            d.saldo = Nothing
            d.importedescuento = Nothing
            d.numeroorden = Nothing
            d.idtipodocumento1 = Nothing
            d.serie1 = Nothing
            d.numerodocumento1 = Nothing
            d.descripcion = Nothing
            d.estado = Nothing
            d.facturaguia = Nothing
            d.idtransportista = Nothing
            d.idcentrocosto = Nothing
            d.idmaquina = Nothing
            d.destino = Nothing
            d.idtipofactura = Nothing
            d.idtipoanexo = Nothing
            d.idanexo = Nothing
            d.descuneto1 = Nothing
            d.descuento2 = Nothing
            d.flete = Nothing
            d.embalaje = Nothing
            d.tasa = Nothing
            d.idusuariooperador = Nothing
            d.idusuariosectorista = Nothing
            d.idcadena = Nothing
            d.idinternocadena = Nothing
            d.idautorizacion = Nothing
            d.reparto = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.usuariomod = Nothing
            d.fechamod = Nothing
            d.idtiponotacredito = Nothing
            d.linea = Nothing
            d.impreso = Nothing
            d.anuladonc = Nothing
            d.idvendedor1 = Nothing
            d.igv = Nothing
            d.idchofer = Nothing
            d.idzonaventa = Nothing
            d.idtipodocumento2 = Nothing
            d.numerodocumento2 = Nothing
            d.idsubdiario = Nothing
            d.nrocontable = Nothing
            d.importetotalus = Nothing
            d.importetotalmn = Nothing
            d.importeigvus = Nothing
            d.importeigvmn = Nothing
            d.estadosunat = Nothing
            d.codigohas = Nothing
            d.barrapdf417 = Nothing
            d.signaturevalue = Nothing
            d.tipooperacion = Nothing
            d.tipoafecigv = Nothing
            d.islote = Nothing
            d.valorventa = Nothing
            d.idturno = Nothing
            d.isliq = Nothing
            d.fechadocumento2 = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NVales, Retornatable As Boolean) As NVales

        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@numeropedido", "@fechadocumento", "@fechavencimineto", "@debehaber", "@idvendedor", "@idcaja", "@idcliente", "@nombrecliente", "@direccion", "@ruc", "@idalmacen", "@idformaventa", "@idmoneda", "@tipocambio", "@importetotal", "@importeigv", "@saldo", "@importedescuento", "@numeroorden", "@idtipodocumento1", "@serie1", "@numerodocumento1", "@descripcion", "@estado", "@facturaguia", "@idtransportista", "@idcentrocosto", "@idmaquina", "@destino", "@idtipofactura", "@idtipoanexo", "@idanexo", "@descuneto1", "@descuento2", "@flete", "@embalaje", "@tasa", "@idusuariooperador", "@idusuariosectorista", "@idcadena", "@idinternocadena", "@idautorizacion", "@reparto", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@idtiponotacredito", "@linea", "@impreso", "@anuladonc", "@idvendedor1", "@igv", "@idchofer", "@idzonaventa", "@idtipodocumento2", "@numerodocumento2", "@idsubdiario", "@nrocontable", "@importetotalus", "@importetotalmn", "@importeigvus", "@importeigvmn", "@estadosunat", "@codigohas", "@barrapdf417", "@signaturevalue", "@tipooperacion", "@tipoafecigv", "@islote", "@valorventa", "@idturno", "@isliq", "@fechadocumento2"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarBinary, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime}
        Dim valores() As Object = {d.fechadocumento2 = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Vales_U_S", parametros, valores, tipoParametro, 78).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.numeropedido = IIf(dt.Rows(0).Item("numeropedido") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeropedido"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.fechavencimineto = IIf(dt.Rows(0).Item("fechavencimineto") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechavencimineto"))
            d.debehaber = IIf(dt.Rows(0).Item("debehaber") Is DBNull.Value, Nothing, dt.Rows(0).Item("debehaber"))
            d.idvendedor = IIf(dt.Rows(0).Item("idvendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idvendedor"))
            d.idcaja = IIf(dt.Rows(0).Item("idcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcaja"))
            d.idcliente = IIf(dt.Rows(0).Item("idcliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcliente"))
            d.nombrecliente = IIf(dt.Rows(0).Item("nombrecliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecliente"))
            d.direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.ruc = IIf(dt.Rows(0).Item("ruc") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruc"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idformaventa = IIf(dt.Rows(0).Item("idformaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idformaventa"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.importetotal = IIf(dt.Rows(0).Item("importetotal") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotal"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.importedescuento = IIf(dt.Rows(0).Item("importedescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento"))
            d.numeroorden = IIf(dt.Rows(0).Item("numeroorden") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroorden"))
            d.idtipodocumento1 = IIf(dt.Rows(0).Item("idtipodocumento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento1"))
            d.serie1 = IIf(dt.Rows(0).Item("serie1") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie1"))
            d.numerodocumento1 = IIf(dt.Rows(0).Item("numerodocumento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento1"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.facturaguia = IIf(dt.Rows(0).Item("facturaguia") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaguia"))
            d.idtransportista = IIf(dt.Rows(0).Item("idtransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtransportista"))
            d.idcentrocosto = IIf(dt.Rows(0).Item("idcentrocosto") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcentrocosto"))
            d.idmaquina = IIf(dt.Rows(0).Item("idmaquina") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmaquina"))
            d.destino = IIf(dt.Rows(0).Item("destino") Is DBNull.Value, Nothing, dt.Rows(0).Item("destino"))
            d.idtipofactura = IIf(dt.Rows(0).Item("idtipofactura") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipofactura"))
            d.idtipoanexo = IIf(dt.Rows(0).Item("idtipoanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipoanexo"))
            d.idanexo = IIf(dt.Rows(0).Item("idanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idanexo"))
            d.descuneto1 = IIf(dt.Rows(0).Item("descuneto1") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuneto1"))
            d.descuento2 = IIf(dt.Rows(0).Item("descuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento2"))
            d.flete = IIf(dt.Rows(0).Item("flete") Is DBNull.Value, Nothing, dt.Rows(0).Item("flete"))
            d.embalaje = IIf(dt.Rows(0).Item("embalaje") Is DBNull.Value, Nothing, dt.Rows(0).Item("embalaje"))
            d.tasa = IIf(dt.Rows(0).Item("tasa") Is DBNull.Value, Nothing, dt.Rows(0).Item("tasa"))
            d.idusuariooperador = IIf(dt.Rows(0).Item("idusuariooperador") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariooperador"))
            d.idusuariosectorista = IIf(dt.Rows(0).Item("idusuariosectorista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariosectorista"))
            d.idcadena = IIf(dt.Rows(0).Item("idcadena") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcadena"))
            d.idinternocadena = IIf(dt.Rows(0).Item("idinternocadena") Is DBNull.Value, Nothing, dt.Rows(0).Item("idinternocadena"))
            d.idautorizacion = IIf(dt.Rows(0).Item("idautorizacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("idautorizacion"))
            d.reparto = IIf(dt.Rows(0).Item("reparto") Is DBNull.Value, Nothing, dt.Rows(0).Item("reparto"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.idtiponotacredito = IIf(dt.Rows(0).Item("idtiponotacredito") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtiponotacredito"))
            d.linea = IIf(dt.Rows(0).Item("linea") Is DBNull.Value, Nothing, dt.Rows(0).Item("linea"))
            d.impreso = IIf(dt.Rows(0).Item("impreso") Is DBNull.Value, Nothing, dt.Rows(0).Item("impreso"))
            d.anuladonc = IIf(dt.Rows(0).Item("anuladonc") Is DBNull.Value, Nothing, dt.Rows(0).Item("anuladonc"))
            d.idvendedor1 = IIf(dt.Rows(0).Item("idvendedor1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idvendedor1"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.idchofer = IIf(dt.Rows(0).Item("idchofer") Is DBNull.Value, Nothing, dt.Rows(0).Item("idchofer"))
            d.idzonaventa = IIf(dt.Rows(0).Item("idzonaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idzonaventa"))
            d.idtipodocumento2 = IIf(dt.Rows(0).Item("idtipodocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento2"))
            d.numerodocumento2 = IIf(dt.Rows(0).Item("numerodocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento2"))
            d.idsubdiario = IIf(dt.Rows(0).Item("idsubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsubdiario"))
            d.nrocontable = IIf(dt.Rows(0).Item("nrocontable") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrocontable"))
            d.importetotalus = IIf(dt.Rows(0).Item("importetotalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalus"))
            d.importetotalmn = IIf(dt.Rows(0).Item("importetotalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalmn"))
            d.importeigvus = IIf(dt.Rows(0).Item("importeigvus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvus"))
            d.importeigvmn = IIf(dt.Rows(0).Item("importeigvmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvmn"))
            d.estadosunat = IIf(dt.Rows(0).Item("estadosunat") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadosunat"))
            d.codigohas = IIf(dt.Rows(0).Item("codigohas") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigohas"))
            d.barrapdf417 = IIf(dt.Rows(0).Item("barrapdf417") Is DBNull.Value, Nothing, dt.Rows(0).Item("barrapdf417"))
            d.signaturevalue = IIf(dt.Rows(0).Item("signaturevalue") Is DBNull.Value, Nothing, dt.Rows(0).Item("signaturevalue"))
            d.tipooperacion = IIf(dt.Rows(0).Item("tipooperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipooperacion"))
            d.tipoafecigv = IIf(dt.Rows(0).Item("tipoafecigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoafecigv"))
            d.islote = IIf(dt.Rows(0).Item("islote") Is DBNull.Value, Nothing, dt.Rows(0).Item("islote"))
            d.valorventa = IIf(dt.Rows(0).Item("valorventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventa"))
            d.idturno = IIf(dt.Rows(0).Item("idturno") Is DBNull.Value, Nothing, dt.Rows(0).Item("idturno"))
            d.isliq = IIf(dt.Rows(0).Item("isliq") Is DBNull.Value, Nothing, dt.Rows(0).Item("isliq"))
            d.fechadocumento2 = IIf(dt.Rows(0).Item("fechadocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento2"))
        Else
            d.idagencia = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.numeropedido = Nothing
            d.fechadocumento = Nothing
            d.fechavencimineto = Nothing
            d.debehaber = Nothing
            d.idvendedor = Nothing
            d.idcaja = Nothing
            d.idcliente = Nothing
            d.nombrecliente = Nothing
            d.direccion = Nothing
            d.ruc = Nothing
            d.idalmacen = Nothing
            d.idformaventa = Nothing
            d.idmoneda = Nothing
            d.tipocambio = Nothing
            d.importetotal = Nothing
            d.importeigv = Nothing
            d.saldo = Nothing
            d.importedescuento = Nothing
            d.numeroorden = Nothing
            d.idtipodocumento1 = Nothing
            d.serie1 = Nothing
            d.numerodocumento1 = Nothing
            d.descripcion = Nothing
            d.estado = Nothing
            d.facturaguia = Nothing
            d.idtransportista = Nothing
            d.idcentrocosto = Nothing
            d.idmaquina = Nothing
            d.destino = Nothing
            d.idtipofactura = Nothing
            d.idtipoanexo = Nothing
            d.idanexo = Nothing
            d.descuneto1 = Nothing
            d.descuento2 = Nothing
            d.flete = Nothing
            d.embalaje = Nothing
            d.tasa = Nothing
            d.idusuariooperador = Nothing
            d.idusuariosectorista = Nothing
            d.idcadena = Nothing
            d.idinternocadena = Nothing
            d.idautorizacion = Nothing
            d.reparto = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.usuariomod = Nothing
            d.fechamod = Nothing
            d.idtiponotacredito = Nothing
            d.linea = Nothing
            d.impreso = Nothing
            d.anuladonc = Nothing
            d.idvendedor1 = Nothing
            d.igv = Nothing
            d.idchofer = Nothing
            d.idzonaventa = Nothing
            d.idtipodocumento2 = Nothing
            d.numerodocumento2 = Nothing
            d.idsubdiario = Nothing
            d.nrocontable = Nothing
            d.importetotalus = Nothing
            d.importetotalmn = Nothing
            d.importeigvus = Nothing
            d.importeigvmn = Nothing
            d.estadosunat = Nothing
            d.codigohas = Nothing
            d.barrapdf417 = Nothing
            d.signaturevalue = Nothing
            d.tipooperacion = Nothing
            d.tipoafecigv = Nothing
            d.islote = Nothing
            d.valorventa = Nothing
            d.idturno = Nothing
            d.isliq = Nothing
            d.fechadocumento2 = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NVales)
        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@idalmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.idalmacen}
        sql.EjecutarProcedure("Str_Vales_D", parametros, valores, tipoParametro, 5)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@idalmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Vales_S", parametros, valores, tipoParametro, 5).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NVales) As DataTable
        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@idalmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.idalmacen}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Vales_S", parametros, valores, tipoParametro, 5).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NVales) As NVales
        Dim parametros() As Object = {"@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@idalmacen"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.idalmacen}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Vales_S", parametros, valores, tipoParametro, 5).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.numeropedido = IIf(dt.Rows(0).Item("numeropedido") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeropedido"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.fechavencimineto = IIf(dt.Rows(0).Item("fechavencimineto") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechavencimineto"))
            d.debehaber = IIf(dt.Rows(0).Item("debehaber") Is DBNull.Value, Nothing, dt.Rows(0).Item("debehaber"))
            d.idvendedor = IIf(dt.Rows(0).Item("idvendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idvendedor"))
            d.idcaja = IIf(dt.Rows(0).Item("idcaja") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcaja"))
            d.idcliente = IIf(dt.Rows(0).Item("idcliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcliente"))
            d.nombrecliente = IIf(dt.Rows(0).Item("nombrecliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecliente"))
            d.direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.ruc = IIf(dt.Rows(0).Item("ruc") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruc"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idformaventa = IIf(dt.Rows(0).Item("idformaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idformaventa"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.importetotal = IIf(dt.Rows(0).Item("importetotal") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotal"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.importedescuento = IIf(dt.Rows(0).Item("importedescuento") Is DBNull.Value, Nothing, dt.Rows(0).Item("importedescuento"))
            d.numeroorden = IIf(dt.Rows(0).Item("numeroorden") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroorden"))
            d.idtipodocumento1 = IIf(dt.Rows(0).Item("idtipodocumento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento1"))
            d.serie1 = IIf(dt.Rows(0).Item("serie1") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie1"))
            d.numerodocumento1 = IIf(dt.Rows(0).Item("numerodocumento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento1"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.facturaguia = IIf(dt.Rows(0).Item("facturaguia") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaguia"))
            d.idtransportista = IIf(dt.Rows(0).Item("idtransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtransportista"))
            d.idcentrocosto = IIf(dt.Rows(0).Item("idcentrocosto") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcentrocosto"))
            d.idmaquina = IIf(dt.Rows(0).Item("idmaquina") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmaquina"))
            d.destino = IIf(dt.Rows(0).Item("destino") Is DBNull.Value, Nothing, dt.Rows(0).Item("destino"))
            d.idtipofactura = IIf(dt.Rows(0).Item("idtipofactura") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipofactura"))
            d.idtipoanexo = IIf(dt.Rows(0).Item("idtipoanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipoanexo"))
            d.idanexo = IIf(dt.Rows(0).Item("idanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idanexo"))
            d.descuneto1 = IIf(dt.Rows(0).Item("descuneto1") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuneto1"))
            d.descuento2 = IIf(dt.Rows(0).Item("descuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento2"))
            d.flete = IIf(dt.Rows(0).Item("flete") Is DBNull.Value, Nothing, dt.Rows(0).Item("flete"))
            d.embalaje = IIf(dt.Rows(0).Item("embalaje") Is DBNull.Value, Nothing, dt.Rows(0).Item("embalaje"))
            d.tasa = IIf(dt.Rows(0).Item("tasa") Is DBNull.Value, Nothing, dt.Rows(0).Item("tasa"))
            d.idusuariooperador = IIf(dt.Rows(0).Item("idusuariooperador") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariooperador"))
            d.idusuariosectorista = IIf(dt.Rows(0).Item("idusuariosectorista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idusuariosectorista"))
            d.idcadena = IIf(dt.Rows(0).Item("idcadena") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcadena"))
            d.idinternocadena = IIf(dt.Rows(0).Item("idinternocadena") Is DBNull.Value, Nothing, dt.Rows(0).Item("idinternocadena"))
            d.idautorizacion = IIf(dt.Rows(0).Item("idautorizacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("idautorizacion"))
            d.reparto = IIf(dt.Rows(0).Item("reparto") Is DBNull.Value, Nothing, dt.Rows(0).Item("reparto"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.idtiponotacredito = IIf(dt.Rows(0).Item("idtiponotacredito") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtiponotacredito"))
            d.linea = IIf(dt.Rows(0).Item("linea") Is DBNull.Value, Nothing, dt.Rows(0).Item("linea"))
            d.impreso = IIf(dt.Rows(0).Item("impreso") Is DBNull.Value, Nothing, dt.Rows(0).Item("impreso"))
            d.anuladonc = IIf(dt.Rows(0).Item("anuladonc") Is DBNull.Value, Nothing, dt.Rows(0).Item("anuladonc"))
            d.idvendedor1 = IIf(dt.Rows(0).Item("idvendedor1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idvendedor1"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.idchofer = IIf(dt.Rows(0).Item("idchofer") Is DBNull.Value, Nothing, dt.Rows(0).Item("idchofer"))
            d.idzonaventa = IIf(dt.Rows(0).Item("idzonaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idzonaventa"))
            d.idtipodocumento2 = IIf(dt.Rows(0).Item("idtipodocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento2"))
            d.numerodocumento2 = IIf(dt.Rows(0).Item("numerodocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento2"))
            d.idsubdiario = IIf(dt.Rows(0).Item("idsubdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsubdiario"))
            d.nrocontable = IIf(dt.Rows(0).Item("nrocontable") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrocontable"))
            d.importetotalus = IIf(dt.Rows(0).Item("importetotalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalus"))
            d.importetotalmn = IIf(dt.Rows(0).Item("importetotalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalmn"))
            d.importeigvus = IIf(dt.Rows(0).Item("importeigvus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvus"))
            d.importeigvmn = IIf(dt.Rows(0).Item("importeigvmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvmn"))
            d.estadosunat = IIf(dt.Rows(0).Item("estadosunat") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadosunat"))
            d.codigohas = IIf(dt.Rows(0).Item("codigohas") Is DBNull.Value, Nothing, dt.Rows(0).Item("codigohas"))
            d.barrapdf417 = IIf(dt.Rows(0).Item("barrapdf417") Is DBNull.Value, Nothing, dt.Rows(0).Item("barrapdf417"))
            d.signaturevalue = IIf(dt.Rows(0).Item("signaturevalue") Is DBNull.Value, Nothing, dt.Rows(0).Item("signaturevalue"))
            d.tipooperacion = IIf(dt.Rows(0).Item("tipooperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipooperacion"))
            d.tipoafecigv = IIf(dt.Rows(0).Item("tipoafecigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoafecigv"))
            d.islote = IIf(dt.Rows(0).Item("islote") Is DBNull.Value, Nothing, dt.Rows(0).Item("islote"))
            d.valorventa = IIf(dt.Rows(0).Item("valorventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("valorventa"))
            d.idturno = IIf(dt.Rows(0).Item("idturno") Is DBNull.Value, Nothing, dt.Rows(0).Item("idturno"))
            d.isliq = IIf(dt.Rows(0).Item("isliq") Is DBNull.Value, Nothing, dt.Rows(0).Item("isliq"))
            d.fechadocumento2 = IIf(dt.Rows(0).Item("fechadocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento2"))
        Else
            d.idagencia = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.numeropedido = Nothing
            d.fechadocumento = Nothing
            d.fechavencimineto = Nothing
            d.debehaber = Nothing
            d.idvendedor = Nothing
            d.idcaja = Nothing
            d.idcliente = Nothing
            d.nombrecliente = Nothing
            d.direccion = Nothing
            d.ruc = Nothing
            d.idalmacen = Nothing
            d.idformaventa = Nothing
            d.idmoneda = Nothing
            d.tipocambio = Nothing
            d.importetotal = Nothing
            d.importeigv = Nothing
            d.saldo = Nothing
            d.importedescuento = Nothing
            d.numeroorden = Nothing
            d.idtipodocumento1 = Nothing
            d.serie1 = Nothing
            d.numerodocumento1 = Nothing
            d.descripcion = Nothing
            d.estado = Nothing
            d.facturaguia = Nothing
            d.idtransportista = Nothing
            d.idcentrocosto = Nothing
            d.idmaquina = Nothing
            d.destino = Nothing
            d.idtipofactura = Nothing
            d.idtipoanexo = Nothing
            d.idanexo = Nothing
            d.descuneto1 = Nothing
            d.descuento2 = Nothing
            d.flete = Nothing
            d.embalaje = Nothing
            d.tasa = Nothing
            d.idusuariooperador = Nothing
            d.idusuariosectorista = Nothing
            d.idcadena = Nothing
            d.idinternocadena = Nothing
            d.idautorizacion = Nothing
            d.reparto = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.usuariomod = Nothing
            d.fechamod = Nothing
            d.idtiponotacredito = Nothing
            d.linea = Nothing
            d.impreso = Nothing
            d.anuladonc = Nothing
            d.idvendedor1 = Nothing
            d.igv = Nothing
            d.idchofer = Nothing
            d.idzonaventa = Nothing
            d.idtipodocumento2 = Nothing
            d.numerodocumento2 = Nothing
            d.idsubdiario = Nothing
            d.nrocontable = Nothing
            d.importetotalus = Nothing
            d.importetotalmn = Nothing
            d.importeigvus = Nothing
            d.importeigvmn = Nothing
            d.estadosunat = Nothing
            d.codigohas = Nothing
            d.barrapdf417 = Nothing
            d.signaturevalue = Nothing
            d.tipooperacion = Nothing
            d.tipoafecigv = Nothing
            d.islote = Nothing
            d.valorventa = Nothing
            d.idturno = Nothing
            d.isliq = Nothing
            d.fechadocumento2 = Nothing
        End If
        Return d
    End Function
#End Region

End Class
