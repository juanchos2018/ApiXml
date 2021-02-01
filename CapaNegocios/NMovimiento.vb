Imports CapaDatos
Public Class NMovimiento
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idalmacen As String
    Public Property idalmacen1 As String
    Public Property tipodocumento As String
    Public Property numerodocumento As String
    Public Property idagencia As String
    Public Property fechadocumento As System.DateTime
    Public Property fechavencimiento As System.DateTime
    Public Property tipomovimiento As String
    Public Property idmovimiento As String
    Public Property situacion As String
    Public Property tipodocumento2 As String
    Public Property numerodocumento1 As String
    Public Property numerodocumento2 As String
    Public Property solicitante As String
    Public Property centrocosto As String
    Public Property idalmacen2 As String
    Public Property glosa1 As String
    Public Property glosa2 As String
    Public Property glosa3 As String
    Public Property idtipoanexo As String
    Public Property idanexo As String
    Public Property fechacrea As System.DateTime
    Public Property usuariocrea As String
    Public Property fechamod As System.DateTime
    Public Property usuariomod As String
    Public Property idcliente As String
    Public Property nombrecliente As String
    Public Property ruc As String
    Public Property idgrupocliente As String
    Public Property idclienteinterno As String
    Public Property idtransportista As String
    Public Property nombretransportista As String
    Public Property placavehiculo As String
    Public Property tipoguiaremision As String
    Public Property situacionguia As String
    Public Property guiafacturada As String
    Public Property iddireccionentrega As String
    Public Property direccionentrega As String
    Public Property numeroordencompra As String
    Public Property tipoorden As String
    Public Property guiadevuelta As String
    Public Property idproveedor As String
    Public Property nombreproveedor As String
    Public Property idcompania As String
    Public Property formaventa As String
    Public Property moneda As String
    Public Property idvendedor As String
    Public Property tipocambio As Decimal
    Public Property numeropedido As String
    Public Property direccionfiscal As String
    Public Property importetotalventa As Decimal
    Public Property tipocambiomvc As String
    Public Property subdiario As String
    Public Property comprobante As String
    Public Property descuento1 As Decimal
    Public Property descuento2 As Decimal
    Public Property tipoguia As String
    Public Property flete As Decimal
    Public Property idautorizacion As String
    Public Property tipodocumento3 As String
    Public Property numerodocumento3 As String
    Public Property serie As String
    Public Property serie1 As String
    Public Property finalizado As Boolean
    Public Property c5_dfecanu As System.DateTime
    Public Property idchofer As String
    Public Property marcatr As String
    Public Property licenciatr As String
    Public Property permisomtc As String
    Public Property motivoguia As String
    Public Property nombrechofer As String
    Public Property Ubigeo As String
    Public Property EstadoSunat As String


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub agregar(m As NMovimiento)
        Dim params() As Object = {
            "@IdAlmacen", "@TipoDocumento", "@NumeroDocumento", "@IdAgencia", "@FechaDocumento",
            "@FechaVencimiento", "@TipoMovimiento", "@IdMovimiento", "@Situacion", "@TipoDocumento2",
            "@NumeroDocumento2", "@Solicitante", "@CentroCosto", "@IdAlmacen2", "@Glosa1",
            "@Glosa2", "@Glosa3", "@idTipoAnexo", "@IdAnexo", "@FechaCrea",
            "@UsuarioCrea", "@FechaMod", "@UsuarioMod", "@IdCliente", "@NombreCliente",
            "@RUC", "@IdGrupoCliente", "@IdClienteInterno", "@IdTransportista", "@NombreTransportista",
            "@PlacaVehiculo", "@TipoGuiaRemision", "@SituacionGuia", "@GuiaFacturada", "@IdDireccionEntrega",
            "@DireccionEntrega", "@NumeroOrdenCompra", "@TipoOrden", "@GuiaDevuelta", "@IdProveedor",
            "@NombreProveedor", "@IdCompania", "@FormaVenta", "@Moneda", "@IdVendedor",
            "@TipoCambio", "@NumeroPedido", "@DireccionFiscal", "@ImporteTotalVenta", "@TipoCambioMVC",
            "@Subdiario", "@Comprobante", "@Descuento1", "@Descuento2", "@TipoGuia",
            "@Flete", "@IdAutorizacion", "@TipoDocumento3", "@NumeroDocumento3", "@serie",
            "@Finalizado", "@IdChofer", "@MarcaTR", "@LicenciaTr",
            "@PermisoMTC", "@MotivoGuia", "@NombreChofer"}

        Dim vals() As Object = {
            m.idalmacen, m.tipodocumento, m.numerodocumento, m.idagencia, m.fechadocumento,
            m.fechavencimiento, m.tipomovimiento, m.idmovimiento, m.situacion, m.tipodocumento2,
            m.numerodocumento2, m.solicitante, m.centrocosto, m.idalmacen2, m.glosa1,
            m.glosa2, m.glosa3, m.idtipoanexo, m.idanexo, m.fechacrea,
            m.usuariocrea, m.fechamod, m.usuariomod, m.idcliente, m.nombrecliente,
            m.ruc, m.idgrupocliente, m.idclienteinterno, m.idtransportista, m.nombretransportista,
            m.placavehiculo, m.tipoguiaremision, m.situacionguia, m.guiafacturada, m.iddireccionentrega,
            m.direccionentrega, m.numeroordencompra, m.tipoorden, m.guiadevuelta, m.idproveedor,
            m.nombreproveedor, m.idcompania, m.formaventa, m.moneda, m.idvendedor,
            m.tipocambio, m.numeropedido, m.direccionfiscal, m.importetotalventa, m.tipocambiomvc,
            m.subdiario, m.comprobante, m.descuento1, m.descuento2, m.tipoguia,
            m.flete, m.idautorizacion, m.tipodocumento3, m.numerodocumento3, m.serie,
            m.finalizado, m.idchofer, m.marcatr, m.licenciatr,
            m.permisomtc, m.motivoguia, m.nombrechofer}

        Dim tipoParametro() As Object = {
           SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime,
       SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar,
       SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar,
       SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime,
       SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar,
       SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar,
       SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar,
       SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar,
       SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar,
       SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal,
       SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar,
       SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar,
       SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar,
       SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar
       }

        sql.EjecutarProcedure("Str_AddMovimiento", params, vals, tipoParametro, 67)
        'Return sql.EjecutarProcedure("Str_AddMovimiento", params, vals, tipoParametro, 67)
    End Sub

    Public Sub Agregar1(d As NMovimiento)
        Dim parametros() As Object = {"@idalmacen", "@tipodocumento", "@numerodocumento", "@idagencia", "@fechadocumento", "@fechavencimiento", "@tipomovimiento", "@idmovimiento", "@situacion", "@tipodocumento2", "@numerodocumento2", "@solicitante", "@centrocosto", "@idalmacen2", "@glosa1", "@glosa2", "@glosa3", "@idtipoanexo", "@idanexo", "@fechacrea", "@usuariocrea", "@fechamod", "@usuariomod", "@idcliente", "@nombrecliente", "@ruc", "@idgrupocliente", "@idclienteinterno", "@idtransportista", "@nombretransportista", "@placavehiculo", "@tipoguiaremision", "@situacionguia", "@guiafacturada", "@iddireccionentrega", "@direccionentrega", "@numeroordencompra", "@tipoorden", "@guiadevuelta", "@idproveedor", "@nombreproveedor", "@idcompania", "@formaventa", "@moneda", "@idvendedor", "@tipocambio", "@numeropedido", "@direccionfiscal", "@importetotalventa", "@tipocambiomvc", "@subdiario", "@comprobante", "@descuento1", "@descuento2", "@tipoguia", "@flete", "@idautorizacion", "@tipodocumento3", "@numerodocumento3", "@serie", "@finalizado", "@c5_dfecanu", "@idchofer", "@marcatr", "@licenciatr", "@permisomtc", "@motivoguia", "@nombrechofer", "@Ubigeo", "@EstadoSunat"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char}
        'Dim valores() As Object = {d.idalmacen, d.tipodocumento, d.numerodocumento, d.idagencia, d.fechadocumento, d.fechavencimiento, d.tipomovimiento, d.idmovimiento, d.situacion, d.tipodocumento2, d.numerodocumento2, d.solicitante, d.centrocosto, d.idalmacen2, d.glosa1, d.glosa2, d.glosa3, d.idtipoanexo, d.idanexo, d.fechacrea, d.usuariocrea, d.fechamod, d.usuariomod, d.idcliente, d.nombrecliente, d.ruc, d.idgrupocliente, d.idclienteinterno, d.idtransportista, d.nombretransportista, d.placavehiculo, d.tipoguiaremision, d.situacionguia, d.guiafacturada, d.iddireccionentrega, d.direccionentrega, d.numeroordencompra, d.tipoorden, d.guiadevuelta, d.idproveedor, d.nombreproveedor, d.idcompania, d.formaventa, d.moneda, d.idvendedor, d.tipocambio, d.numeropedido, d.direccionfiscal, d.importetotalventa, d.tipocambiomvc, d.subdiario, d.comprobante, d.descuento1, d.descuento2, d.tipoguia, d.flete, d.idautorizacion, d.tipodocumento3, d.numerodocumento3, d.serie, d.finalizado, d.c5_dfecanu, d.idchofer, d.marcatr, d.licenciatr, d.permisomtc, d.motivoguia, d.nombrechofer, d.Ubigeo, d.EstadoSunat}
        Dim valores() As Object = {d.idalmacen, d.tipodocumento, d.numerodocumento, d.idagencia, DateTime.Now, DateTime.Now, d.tipomovimiento, d.idmovimiento, d.situacion, d.tipodocumento2, d.numerodocumento2, d.solicitante, d.centrocosto, d.idalmacen2, d.glosa1, d.glosa2, d.glosa3, d.idtipoanexo, d.idanexo, DateTime.Now, d.usuariocrea, d.fechamod, d.usuariomod, d.idcliente, d.nombrecliente, d.ruc, d.idgrupocliente, d.idclienteinterno, d.idtransportista, d.nombretransportista, d.placavehiculo, d.tipoguiaremision, d.situacionguia, d.guiafacturada, d.iddireccionentrega, d.direccionentrega, d.numeroordencompra, d.tipoorden, d.guiadevuelta, d.idproveedor, d.nombreproveedor, d.idcompania, d.formaventa, d.moneda, d.idvendedor, d.tipocambio, d.numeropedido, d.direccionfiscal, d.importetotalventa, d.tipocambiomvc, d.subdiario, d.comprobante, d.descuento1, d.descuento2, d.tipoguia, d.flete, d.idautorizacion, d.tipodocumento3, d.numerodocumento3, d.serie, d.finalizado, d.c5_dfecanu, d.idchofer, d.marcatr, d.licenciatr, d.permisomtc, d.motivoguia, d.nombrechofer, d.Ubigeo, d.EstadoSunat}
        sql.EjecutarProcedure("Str_Movimiento_I", parametros, valores, tipoParametro, 70)
    End Sub

    Public Function Agregar2(d As NMovimiento) As String
        Dim parametros() As Object = {"@idalmacen", "@tipodocumento", "@numerodocumento", "@idagencia", "@fechadocumento", "@fechavencimiento", "@tipomovimiento", "@idmovimiento", "@situacion", "@tipodocumento2", "@numerodocumento2", "@solicitante", "@centrocosto", "@idalmacen2", "@glosa1", "@glosa2", "@glosa3", "@idtipoanexo", "@idanexo", "@fechacrea", "@usuariocrea", "@fechamod", "@usuariomod", "@idcliente", "@nombrecliente", "@ruc", "@idgrupocliente", "@idclienteinterno", "@idtransportista", "@nombretransportista", "@placavehiculo", "@tipoguiaremision", "@situacionguia", "@guiafacturada", "@iddireccionentrega", "@direccionentrega", "@numeroordencompra", "@tipoorden", "@guiadevuelta", "@idproveedor", "@nombreproveedor", "@idcompania", "@formaventa", "@moneda", "@idvendedor", "@tipocambio", "@numeropedido", "@direccionfiscal", "@importetotalventa", "@tipocambiomvc", "@subdiario", "@comprobante", "@descuento1", "@descuento2", "@tipoguia", "@flete", "@idautorizacion", "@tipodocumento3", "@numerodocumento3", "@serie", "@finalizado", "@c5_dfecanu", "@idchofer", "@marcatr", "@licenciatr", "@permisomtc", "@motivoguia", "@nombrechofer", "@Ubigeo", "@EstadoSunat"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char}
        'Dim valores() As Object = {d.idalmacen, d.tipodocumento, d.numerodocumento, d.idagencia, d.fechadocumento, d.fechavencimiento, d.tipomovimiento, d.idmovimiento, d.situacion, d.tipodocumento2, d.numerodocumento2, d.solicitante, d.centrocosto, d.idalmacen2, d.glosa1, d.glosa2, d.glosa3, d.idtipoanexo, d.idanexo, d.fechacrea, d.usuariocrea, d.fechamod, d.usuariomod, d.idcliente, d.nombrecliente, d.ruc, d.idgrupocliente, d.idclienteinterno, d.idtransportista, d.nombretransportista, d.placavehiculo, d.tipoguiaremision, d.situacionguia, d.guiafacturada, d.iddireccionentrega, d.direccionentrega, d.numeroordencompra, d.tipoorden, d.guiadevuelta, d.idproveedor, d.nombreproveedor, d.idcompania, d.formaventa, d.moneda, d.idvendedor, d.tipocambio, d.numeropedido, d.direccionfiscal, d.importetotalventa, d.tipocambiomvc, d.subdiario, d.comprobante, d.descuento1, d.descuento2, d.tipoguia, d.flete, d.idautorizacion, d.tipodocumento3, d.numerodocumento3, d.serie, d.finalizado, d.c5_dfecanu, d.idchofer, d.marcatr, d.licenciatr, d.permisomtc, d.motivoguia, d.nombrechofer, d.Ubigeo, d.EstadoSunat}
        Dim valores() As Object = {d.idalmacen, d.tipodocumento, d.numerodocumento, d.idagencia, DateTime.Now, DateTime.Now, d.tipomovimiento, d.idmovimiento, d.situacion, d.tipodocumento2, d.numerodocumento2, d.solicitante, d.centrocosto, d.idalmacen2, d.glosa1, d.glosa2, d.glosa3, d.idtipoanexo, d.idanexo, DateTime.Now, d.usuariocrea, d.fechamod, d.usuariomod, d.idcliente, d.nombrecliente, d.ruc, d.idgrupocliente, d.idclienteinterno, d.idtransportista, d.nombretransportista, d.placavehiculo, d.tipoguiaremision, d.situacionguia, d.guiafacturada, d.iddireccionentrega, d.direccionentrega, d.numeroordencompra, d.tipoorden, d.guiadevuelta, d.idproveedor, d.nombreproveedor, d.idcompania, d.formaventa, d.moneda, d.idvendedor, d.tipocambio, d.numeropedido, d.direccionfiscal, d.importetotalventa, d.tipocambiomvc, d.subdiario, d.comprobante, d.descuento1, d.descuento2, d.tipoguia, d.flete, d.idautorizacion, d.tipodocumento3, d.numerodocumento3, d.serie, d.finalizado, d.c5_dfecanu, d.idchofer, d.marcatr, d.licenciatr, d.permisomtc, d.motivoguia, d.nombrechofer, d.Ubigeo, d.EstadoSunat}
        '   sql.EjecutarProcedure("Str_Movimiento_I", parametros, valores, tipoParametro, 70)

        Try
            sql.EjecutarProcedure("Str_Movimiento_I", parametros, valores, tipoParametro, 70)
            Return "OK"
        Catch
            Return "Error"
        End Try

    End Function

    Public Function Existe(s As NMovimiento) As Boolean
        Dim existeC As String
        Dim bandera As Boolean = False
        Dim valoresC() As Object = {"'" & s.IdAlmacen & "'", "'" & s.TipoDocumento & "'", "'" & s.NumeroDocumento & "'"}
        existeC = sql.ValorEscalar("dbo.Movimiento_Existe", valoresC, 3)
        If existeC = "1" Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    ''' <summary>
    ''' Obtiene verdadero o falso si el documento a eliminar existe como documento de referencia
    ''' </summary>
    ''' <param name="s"></param>d
    ''' <returns></returns>
    Public Function ExisteEnRefrencia(s As NMovimiento) As Boolean
        Dim bandera As Boolean = False
        Dim cadena As String = " select TipoDocumento from movimiento where IdAlmacen2='" & s.idalmacen & "' and TipoDocumento2='" & s.tipodocumento & "' and NumeroDocumento2='" & s.numerodocumento & "' and situacion<>'A'"
        Dim dt As New DataTable
        dt = sql.EjecutarConsulta("d", cadena).Tables(0)
        If dt.Rows.Count > 0 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function

    Public Function Lista(p As NMovimiento) As DataTable
        Dim cadena As String = " select distinct m.IdAlmacen,m.TipoDocumento,m.NumeroDocumento,m.FechaDocumento,m.IdCliente,m.NombreCliente,m.DireccionEntrega,m.IdAgencia from movimiento "
        cadena += " m inner join detallemovimiento dm on m.tipodocumento=dm.tipodocumento and m.idalmacen=dm.idalmacen and m.numerodocumento=dm.numerodocumento "
        cadena += " inner join vcondicion_entrega vt on rtrim(m.tipoguia)=rtrim(vt.idcodigo) "
        cadena += " where m.tipodocumento='" & p.tipodocumento & "' and m.idalmacen='" & p.idalmacen & "' and m.Situacion='V'  and flag='01' and saldo<>0 "
        Dim dt As New DataTable
        dt = sql.EjecutarConsulta("mov", cadena).Tables(0)
        Return dt
    End Function
    Public Function Lista(TipoDocumento As String, Numero As String, idalmacen As String) As DataTable
        Dim cadena As String = " select distinct m.IdAlmacen,m.TipoDocumento,m.NumeroDocumento,m.FechaDocumento,m.IdCliente,m.NombreCliente,m.DireccionEntrega,m.IdAgencia from movimiento "
        cadena += " m inner join detallemovimiento dm on m.tipodocumento=dm.tipodocumento and m.idalmacen=dm.idalmacen and m.numerodocumento=dm.numerodocumento "
        cadena += " inner join vcondicion_entrega vt on rtrim(m.tipoguia)=rtrim(vt.idcodigo) "
        cadena += " where m.tipodocumento='" & TipoDocumento & "' and m.idalmacen='" & idalmacen & "' and m.Situacion='V'  and flag='02'"
        Dim dt As New DataTable
        dt = sql.EjecutarConsulta("mov", cadena).Tables(0)
        Return dt
    End Function
    ''' <summary>
    ''' Lista los partes de ingreso o salida según tipo de documento especificado
    ''' </summary>
    ''' <param name="p"></param>
    ''' <returns></returns>
    Public Function Partes(p As NMovimiento) As DataTable
        Dim cadena As String = "select distinct m.IdAlmacen,m.TipoDocumento,m.NumeroDocumento,m.FechaDocumento,m.IdCliente,m.NombreCliente,DireccionEntrega,IdAgencia from movimiento "
        cadena += " m left join vcondicion_entrega vt on rtrim(m.tipoguia)=rtrim(vt.idcodigo) "
        cadena += " inner join detallemovimiento dm on m.TipoDocumento=dm.TipoDocumento and m.Numerodocumento=dm.NumeroDocumento and m.idalmacen=dm.idalmacen "
        If IsNothing(p.idalmacen) = True Then
            cadena += " where m.tipodocumento='" & p.tipodocumento & "' and m.Situacion='V' and isnull(Saldo,0.0)<>0 "
        Else
            cadena += " where m.tipodocumento='" & p.tipodocumento & "' and m.idalmacen='" & p.idalmacen & "' and  m.Situacion='V' and isnull(Saldo,0.0)<>0 "
        End If

        Dim dt As New DataTable
        dt = sql.EjecutarConsulta("mov", cadena).Tables(0)
        Return dt
    End Function
    Public Function ExisteMovimiento(p As NMovimiento) As DataTable
        Dim cadena As String = "select TipoDocumento,NumeroDocumento,IdAlmacen from Movimiento"
        cadena += " where idmovimiento='RP' AND IDPROVEEDOR='" & p.idproveedor & "'"
        cadena += --" and TipoDocumento2='" & p.tipodocumento2 & "' and NumeroDocumento2='" & p.numerodocumento2 & "'"
        cadena += " AND TipoMovimiento='" & p.tipomovimiento & "' and isnull(Situacion,'')='V' "
        Return sql.EjecutarConsulta("re", cadena).Tables(0)
    End Function

    Public Function ObtenerCabecera(M As NMovimiento) As NMovimiento
        Dim params() As Object = {"@IdAlmacen", "@TipoDocumento", "@NumeroDocumento", "@IdAgencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim vals() As Object = {M.IdAlmacen, M.TipoDocumento, M.NumeroDocumento, M.IdAgencia}
        Dim dt_mon As New DataTable
        dt_mon = sql.ProcedureSQL("Str_FndMov", params, vals, tipoParametro, 4).Tables(0)
        With dt_mon
            M.IdAlmacen = .Rows(0).Item("IdAlmacen").ToString : M.TipoDocumento = .Rows(0).Item("TipoDocumento").ToString
            M.NumeroDocumento = .Rows(0).Item("NumeroDocumento").ToString : M.IdAgencia = .Rows(0).Item("IdAgencia").ToString
            M.FechaDocumento = CDate(.Rows(0).Item("FechaDocumento"))
            M.TipoMovimiento = .Rows(0).Item("TipoMovimiento").ToString : M.IdMovimiento = .Rows(0).Item("IdMovimiento").ToString
            M.Situacion = .Rows(0).Item("Situacion").ToString : M.TipoDocumento2 = .Rows(0).Item("TipoDocumento2").ToString
            M.NumeroDocumento2 = .Rows(0).Item("NumeroDocumento2").ToString : M.Solicitante = .Rows(0).Item("Solicitante").ToString
            M.CentroCosto = .Rows(0).Item("CentroCosto").ToString : M.IdAlmacen2 = .Rows(0).Item("IdAlmacen2").ToString
            M.Glosa1 = .Rows(0).Item("Glosa1").ToString : M.Glosa2 = .Rows(0).Item("Glosa2").ToString : M.Glosa3 = .Rows(0).Item("Glosa3").ToString
            M.idTipoAnexo = .Rows(0).Item("idTipoAnexo").ToString : M.IdAnexo = .Rows(0).Item("IdAnexo").ToString : M.FechaCrea = CDate(.Rows(0).Item("FechaCrea"))
            M.UsuarioCrea = .Rows(0).Item("UsuarioCrea").ToString : M.UsuarioMod = .Rows(0).Item("UsuarioMod").ToString
            M.IdCliente = .Rows(0).Item("IdCliente").ToString : M.RUC = .Rows(0).Item("Ruc").ToString : M.NombreCliente = .Rows(0).Item("NombreCliente").ToString
            M.IdTransportista = .Rows(0).Item("IdTransportista").ToString : M.NombreTransportista = .Rows(0).Item("NombreTransportista").ToString
            M.PlacaVehiculo = .Rows(0).Item("PlacaVehiculo").ToString : M.TipoGuiaRemision = .Rows(0).Item("TipoGuiaRemision").ToString
            M.SituacionGuia = .Rows(0).Item("SituacionGuia").ToString : M.GuiaFacturada = .Rows(0).Item("GuiaFacturada").ToString
            M.IdDireccionEntrega = .Rows(0).Item("IdDireccionEntrega").ToString : M.DireccionEntrega = .Rows(0).Item("DireccionEntrega").ToString
            M.NumeroOrdenCompra = .Rows(0).Item("NumeroOrdenCompra").ToString : M.TipoOrden = .Rows(0).Item("TipoOrden").ToString : M.GuiaDevuelta = .Rows(0).Item("GuiaDevuelta").ToString
            M.IdProveedor = .Rows(0).Item("IdProveedor").ToString : M.NombreProveedor = .Rows(0).Item("NombreProveedor").ToString
            M.Moneda = .Rows(0).Item("Moneda").ToString : M.IdVendedor = .Rows(0).Item("IdVendedor").ToString : M.TipoCambio = CDec(.Rows(0).Item("TipoCambio"))
            M.NumeroPedido = .Rows(0).Item("NumeroPedido").ToString : M.DireccionFiscal = .Rows(0).Item("DireccionFiscal").ToString
            M.ImporteTotalVenta = CDec(.Rows(0).Item("ImporteTotalVenta")) : M.TipoGuia = .Rows(0).Item("TipoGuia").ToString : M.Flete = CDec(.Rows(0).Item("Flete"))
            M.TipoDocumento3 = .Rows(0).Item("TipoDocumento3").ToString : M.NumeroDocumento3 = .Rows(0).Item("NumeroDocumento3").ToString
            M.serie = .Rows(0).Item("Serie").ToString : M.IdChofer = .Rows(0).Item("IdChofer").ToString
            M.LicenciaTr = .Rows(0).Item("LicenciaTr").ToString : M.MarcaTR = .Rows(0).Item("MarcaTr").ToString : M.PermisoMTC = .Rows(0).Item("PermisoMTC").ToString
            M.MotivoGuia = .Rows(0).Item("MotivoGuia").ToString
        End With
        Return M
    End Function
    Public Function listarConsulta() As DataTable
        Dim cadena As String = " SELECT IdAlmacen, TipoDocumento, NumeroDocumento, FechaDocumento, TipoMovimiento, IdMovimiento, isnull(Situacion,'V') AS Situacion, "
        cadena += " case when tipomovimiento='S' Then idCliente else idproveedor end as Ruc,  "
        cadena += " case when tipomovimiento='S' Then NombreCliente else nombreProveedor end as  RazonSocial,  "
        cadena += " TipoDocumento2, NumeroDocumento2,isnull(TipoDocumento3,'') as TipoDocumento3,isnull(NumeroDocumento3,'') as NumeroDocumento3, "
        cadena += " IdAgencia,serie,Moneda,TipoCambio,IdAlmacen2,isnull(TipoGuia,'') as TipoGuia,u.alias,m.fechacrea FROM Movimiento m "
        cadena += " left join PTUsuario u on m.UsuarioCrea=u.IdUsuario "
        Dim dt As DataTable = sql.EjecutarConsulta("ca", cadena).Tables(0)
        Return dt
    End Function
    Public Function listarConsulta(i As String, f As String) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime}
        If Date.TryParse(i, Now.Date) = False Then
            i = Nothing
        End If
        If Date.TryParse(f, Now.Date) = False Then
            f = Nothing
        End If
        Dim valores() As Object = {i, f}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_ListaMovimiento", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function

    Public Function listarConsultaB(i As String, f As String) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime}
        If Date.TryParse(i, Now.Date) = False Then
            i = Nothing
        End If
        If Date.TryParse(f, Now.Date) = False Then
            f = Nothing
        End If
        Dim valores() As Object = {i, f}
        Return sql.Proc_DataReader("Str_ListaMovimiento", parametros, valores, tipoParametro, 2)
    End Function
    Public Function ListaGuias(i As String, f As String, mostrar As String) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF", "@soloconsaldo"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar}
        If Date.TryParse(i, Now.Date) = False Then
            i = Nothing
        End If
        If Date.TryParse(f, Now.Date) = False Then
            f = Nothing
        End If
        Dim valores() As Object = {i, f, mostrar}
        Return sql.Proc_DataReader("Str_ListaMovimientoGuias", parametros, valores, tipoParametro, 3)
    End Function
    Public Function ListaDevoluciones(d As NMovimiento) As DataTable
        Dim parametros() As Object = {"@IdMovimiento"}
        Dim tipoParametros() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idmovimiento}
        Return sql.Proc_DataReader("Str_ListaMovimientoDevoluciones", parametros, valores, tipoParametros, 1)
    End Function
    Public Function BuscarGuias(i As String, f As String) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime}
        If Date.TryParse(i, Now.Date) = False Then
            i = Nothing
        End If
        If Date.TryParse(f, Now.Date) = False Then
            f = Nothing
        End If
        Dim valores() As Object = {i, f}
        Return sql.Proc_DataReader("Str_ListaMovimiento", parametros, valores, tipoParametro, 2)
    End Function


    Public Sub Eliminar(d As NMovimiento)
        Dim parametros() As Object = {"@idAlmacen", "@tipoDocumento", "@numeroDocumento", "@idAgencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.IdAlmacen, d.TipoDocumento, d.NumeroDocumento, d.IdAgencia}
        sql.EjecutarProcedure("Str_Movimiento_D", parametros, valores, tipoParametro, 4)
    End Sub

    Public Sub Actualizar(d As NMovimiento)
        Dim parametros() As Object = {"@idalmacen", "@tipodocumento", "@numerodocumento", "@idagencia", "@fechadocumento", "@fechavencimiento", "@tipomovimiento", "@idmovimiento", "@situacion", "@tipodocumento2", "@numerodocumento2", "@solicitante", "@centrocosto", "@idalmacen2", "@glosa1", "@glosa2", "@glosa3", "@idtipoanexo", "@idanexo", "@fechacrea", "@usuariocrea", "@fechamod", "@usuariomod", "@idcliente", "@nombrecliente", "@ruc", "@idgrupocliente", "@idclienteinterno", "@idtransportista", "@nombretransportista", "@placavehiculo", "@tipoguiaremision", "@situacionguia", "@guiafacturada", "@iddireccionentrega", "@direccionentrega", "@numeroordencompra", "@tipoorden", "@guiadevuelta", "@idproveedor", "@nombreproveedor", "@idcompania", "@formaventa", "@moneda", "@idvendedor", "@tipocambio", "@numeropedido", "@direccionfiscal", "@importetotalventa", "@tipocambiomvc", "@subdiario", "@comprobante", "@descuento1", "@descuento2", "@tipoguia", "@flete", "@idautorizacion", "@tipodocumento3", "@numerodocumento3", "@serie", "@finalizado", "@c5_dfecanu", "@idchofer", "@marcatr", "@licenciatr", "@permisomtc", "@motivoguia", "@nombrechofer", "@Ubigeo", "@EstadoSunat"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char}
        Dim valores() As Object = {d.idalmacen, d.tipodocumento, d.numerodocumento, d.idagencia, d.fechadocumento, d.fechavencimiento, d.tipomovimiento, d.idmovimiento, d.situacion, d.tipodocumento2, d.numerodocumento2, d.solicitante, d.centrocosto, d.idalmacen2, d.glosa1, d.glosa2, d.glosa3, d.idtipoanexo, d.idanexo, d.fechacrea, d.usuariocrea, d.fechamod, d.usuariomod, d.idcliente, d.nombrecliente, d.ruc, d.idgrupocliente, d.idclienteinterno, d.idtransportista, d.nombretransportista, d.placavehiculo, d.tipoguiaremision, d.situacionguia, d.guiafacturada, d.iddireccionentrega, d.direccionentrega, d.numeroordencompra, d.tipoorden, d.guiadevuelta, d.idproveedor, d.nombreproveedor, d.idcompania, d.formaventa, d.moneda, d.idvendedor, d.tipocambio, d.numeropedido, d.direccionfiscal, d.importetotalventa, d.tipocambiomvc, d.subdiario, d.comprobante, d.descuento1, d.descuento2, d.tipoguia, d.flete, d.idautorizacion, d.tipodocumento3, d.numerodocumento3, d.serie, d.finalizado, d.c5_dfecanu, d.idchofer, d.marcatr, d.licenciatr, d.permisomtc, d.motivoguia, d.nombrechofer, d.Ubigeo, d.EstadoSunat}
        sql.EjecutarProcedure("Str_Movimiento_U", parametros, valores, tipoParametro, 70)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idAlmacen", "@tipoDocumento", "@numeroDocumento", "@idAgencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Movimiento_S", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NMovimiento) As NMovimiento
        Dim parametros() As Object = {"@idAlmacen", "@tipoDocumento", "@numeroDocumento", "@idAgencia"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idalmacen, d.tipodocumento, d.numerodocumento, d.idagencia}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Movimiento_S", parametros, valores, tipoParametro, 4).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idalmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.tipodocumento = IIf(dt.Rows(0).Item("tipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoDocumento"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numeroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento"))
            d.idagencia = IIf(dt.Rows(0).Item("idAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAgencia"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechaDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaDocumento"))
            d.fechavencimiento = IIf(dt.Rows(0).Item("fechaVencimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaVencimiento"))
            d.tipomovimiento = IIf(dt.Rows(0).Item("tipoMovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoMovimiento"))
            d.idmovimiento = IIf(dt.Rows(0).Item("idMovimiento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idMovimiento"))
            d.situacion = IIf(dt.Rows(0).Item("situacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("situacion"))
            d.tipodocumento2 = IIf(dt.Rows(0).Item("tipoDocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoDocumento2"))
            d.numerodocumento2 = IIf(dt.Rows(0).Item("numeroDocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento2"))
            d.solicitante = IIf(dt.Rows(0).Item("solicitante") Is DBNull.Value, Nothing, dt.Rows(0).Item("solicitante"))
            d.centrocosto = IIf(dt.Rows(0).Item("centroCosto") Is DBNull.Value, Nothing, dt.Rows(0).Item("centroCosto"))
            d.idalmacen2 = IIf(dt.Rows(0).Item("idAlmacen2") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen2"))
            d.glosa1 = IIf(dt.Rows(0).Item("glosa1") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa1"))
            d.glosa2 = IIf(dt.Rows(0).Item("glosa2") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa2"))
            d.glosa3 = IIf(dt.Rows(0).Item("glosa3") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa3"))
            d.idtipoanexo = IIf(dt.Rows(0).Item("idTipoAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoAnexo"))
            d.idanexo = IIf(dt.Rows(0).Item("idAnexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAnexo"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.fechamod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuarioMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioMod"))
            d.idcliente = IIf(dt.Rows(0).Item("idCliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCliente"))
            d.nombrecliente = IIf(dt.Rows(0).Item("nombreCliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreCliente"))
            d.ruc = IIf(dt.Rows(0).Item("rUC") Is DBNull.Value, Nothing, dt.Rows(0).Item("rUC"))
            d.idgrupocliente = IIf(dt.Rows(0).Item("idGrupoCliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idGrupoCliente"))
            d.idclienteinterno = IIf(dt.Rows(0).Item("idClienteInterno") Is DBNull.Value, Nothing, dt.Rows(0).Item("idClienteInterno"))
            d.idtransportista = IIf(dt.Rows(0).Item("idTransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTransportista"))
            d.nombretransportista = IIf(dt.Rows(0).Item("nombreTransportista") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreTransportista"))
            d.placavehiculo = IIf(dt.Rows(0).Item("placaVehiculo") Is DBNull.Value, Nothing, dt.Rows(0).Item("placaVehiculo"))
            d.tipoguiaremision = IIf(dt.Rows(0).Item("tipoGuiaRemision") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoGuiaRemision"))
            d.situacionguia = IIf(dt.Rows(0).Item("situacionGuia") Is DBNull.Value, Nothing, dt.Rows(0).Item("situacionGuia"))
            d.guiafacturada = IIf(dt.Rows(0).Item("guiaFacturada") Is DBNull.Value, Nothing, dt.Rows(0).Item("guiaFacturada"))
            d.iddireccionentrega = IIf(dt.Rows(0).Item("idDireccionEntrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("idDireccionEntrega"))
            d.direccionentrega = IIf(dt.Rows(0).Item("direccionEntrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccionEntrega"))
            d.numeroordencompra = IIf(dt.Rows(0).Item("numeroOrdenCompra") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroOrdenCompra"))
            d.tipoorden = IIf(dt.Rows(0).Item("tipoOrden") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoOrden"))
            d.guiadevuelta = IIf(dt.Rows(0).Item("guiaDevuelta") Is DBNull.Value, Nothing, dt.Rows(0).Item("guiaDevuelta"))
            d.idproveedor = IIf(dt.Rows(0).Item("idProveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idProveedor"))
            d.nombreproveedor = IIf(dt.Rows(0).Item("nombreProveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreProveedor"))
            d.idcompania = IIf(dt.Rows(0).Item("idCompania") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCompania"))
            d.formaventa = IIf(dt.Rows(0).Item("formaVenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("formaVenta"))
            d.moneda = IIf(dt.Rows(0).Item("moneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("moneda"))
            d.idvendedor = IIf(dt.Rows(0).Item("idVendedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idVendedor"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipoCambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoCambio"))
            d.numeropedido = IIf(dt.Rows(0).Item("numeroPedido") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroPedido"))
            d.direccionfiscal = IIf(dt.Rows(0).Item("direccionFiscal") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccionFiscal"))
            d.importetotalventa = IIf(dt.Rows(0).Item("importeTotalVenta") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeTotalVenta"))
            d.tipocambiomvc = IIf(dt.Rows(0).Item("tipoCambioMVC") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoCambioMVC"))
            d.subdiario = IIf(dt.Rows(0).Item("subdiario") Is DBNull.Value, Nothing, dt.Rows(0).Item("subdiario"))
            d.comprobante = IIf(dt.Rows(0).Item("comprobante") Is DBNull.Value, Nothing, dt.Rows(0).Item("comprobante"))
            d.descuento1 = IIf(dt.Rows(0).Item("descuento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento1"))
            d.descuento2 = IIf(dt.Rows(0).Item("descuento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("descuento2"))
            d.tipoguia = IIf(dt.Rows(0).Item("tipoGuia") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoGuia"))
            d.flete = IIf(dt.Rows(0).Item("flete") Is DBNull.Value, Nothing, dt.Rows(0).Item("flete"))
            d.idautorizacion = IIf(dt.Rows(0).Item("idAutorizacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAutorizacion"))
            d.tipodocumento3 = IIf(dt.Rows(0).Item("tipoDocumento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoDocumento3"))
            d.numerodocumento3 = IIf(dt.Rows(0).Item("numeroDocumento3") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento3"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.finalizado = IIf(dt.Rows(0).Item("finalizado") Is DBNull.Value, Nothing, dt.Rows(0).Item("finalizado"))
            d.c5_dfecanu = IIf(dt.Rows(0).Item("c5_DFECANU") Is DBNull.Value, Nothing, dt.Rows(0).Item("c5_DFECANU"))
            d.idchofer = IIf(dt.Rows(0).Item("idChofer") Is DBNull.Value, Nothing, dt.Rows(0).Item("idChofer"))
            d.marcatr = IIf(dt.Rows(0).Item("marcaTR") Is DBNull.Value, Nothing, dt.Rows(0).Item("marcaTR"))
            d.licenciatr = IIf(dt.Rows(0).Item("licenciaTr") Is DBNull.Value, Nothing, dt.Rows(0).Item("licenciaTr"))
            d.permisomtc = IIf(dt.Rows(0).Item("permisoMTC") Is DBNull.Value, Nothing, dt.Rows(0).Item("permisoMTC"))
            d.motivoguia = IIf(dt.Rows(0).Item("motivoGuia") Is DBNull.Value, Nothing, dt.Rows(0).Item("motivoGuia"))
            d.nombrechofer = IIf(dt.Rows(0).Item("nombreChofer") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreChofer"))
            d.Ubigeo = IIf(dt.Rows(0).Item("Ubigeo") Is DBNull.Value, Nothing, dt.Rows(0).Item("Ubigeo"))
            d.EstadoSunat = IIf(dt.Rows(0).Item("EstadoSunat") Is DBNull.Value, Nothing, dt.Rows(0).Item("EstadoSunat"))
        Else
            d.idalmacen = Nothing
            d.tipodocumento = Nothing
            d.numerodocumento = Nothing
            d.idagencia = Nothing
            d.fechadocumento = Nothing
            d.fechavencimiento = Nothing
            d.tipomovimiento = Nothing
            d.idmovimiento = Nothing
            d.situacion = Nothing
            d.tipodocumento2 = Nothing
            d.numerodocumento2 = Nothing
            d.solicitante = Nothing
            d.centrocosto = Nothing
            d.idalmacen2 = Nothing
            d.glosa1 = Nothing
            d.glosa2 = Nothing
            d.glosa3 = Nothing
            d.idtipoanexo = Nothing
            d.idanexo = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
            d.fechamod = Nothing
            d.usuariomod = Nothing
            d.idcliente = Nothing
            d.nombrecliente = Nothing
            d.ruc = Nothing
            d.idgrupocliente = Nothing
            d.idclienteinterno = Nothing
            d.idtransportista = Nothing
            d.nombretransportista = Nothing
            d.placavehiculo = Nothing
            d.tipoguiaremision = Nothing
            d.situacionguia = Nothing
            d.guiafacturada = Nothing
            d.iddireccionentrega = Nothing
            d.direccionentrega = Nothing
            d.numeroordencompra = Nothing
            d.tipoorden = Nothing
            d.guiadevuelta = Nothing
            d.idproveedor = Nothing
            d.nombreproveedor = Nothing
            d.idcompania = Nothing
            d.formaventa = Nothing
            d.moneda = Nothing
            d.idvendedor = Nothing
            d.tipocambio = Nothing
            d.numeropedido = Nothing
            d.direccionfiscal = Nothing
            d.importetotalventa = Nothing
            d.tipocambiomvc = Nothing
            d.subdiario = Nothing
            d.comprobante = Nothing
            d.descuento1 = Nothing
            d.descuento2 = Nothing
            d.tipoguia = Nothing
            d.flete = Nothing
            d.idautorizacion = Nothing
            d.tipodocumento3 = Nothing
            d.numerodocumento3 = Nothing
            d.serie = Nothing
            d.finalizado = Nothing
            d.c5_dfecanu = Nothing
            d.idchofer = Nothing
            d.marcatr = Nothing
            d.licenciatr = Nothing
            d.permisomtc = Nothing
            d.motivoguia = Nothing
            d.nombrechofer = Nothing
            d.Ubigeo = Nothing
            d.EstadoSunat = Nothing
        End If
        Return d
    End Function


    Public Function Movimientos(campos As String) As DataTable
        Dim cadena As String = " select IdAlmacen,TipoDocumento,NumeroDocumento,IdAgencia from Movimiento "
        cadena += "  where isnull(Situacion,'')='V' and IdAlmacen+TipoDocumento+NumeroDocumento in(" & campos & ") "
        Return sql.EjecutarConsulta("d", cadena).Tables(0)
    End Function
    Public Function MovimientosPorTipo(TipoDocumento As String) As DataTable
        Dim cadena As String = " select  
           [IdAlmacen] ,[TipoDocumento]  ,[NumeroDocumento] ,[IdCliente] ,[NombreCliente]
           ,[RUC], [FechaDocumento], [TipoMovimiento], [IdMovimiento], [TipoDocumento2],[NumeroDocumento2]
           ,[Glosa2] ,[FormaVenta]  ,[Moneda]  ,[TipoCambio]  ,[NumeroPedido]  ,[DireccionFiscal]
           ,[ImporteTotalVenta]  ,[serie]
        from movimiento where TipoDocumento='" + TipoDocumento + "'  
        order by FechaDocumento desc, numerodocumento desc"
        Return sql.EjecutarConsulta("d", cadena).Tables(0)
    End Function
    Public Function ListaCompras() As DataTable
        Dim cadena As String = " select TipoDocumento2,NumeroDocumento2,FechaDocumento,IdProveedor,NombreProveedor,IdAlmacen,TipoDocumento,NumeroDocumento from Movimiento "
        cadena += "  where Tipodocumento='PE' and isnull(Situacion,'')='V' and month(FechaDocumento)=" & Now.Date.Month & " and year(FechaDocumento)=" & Now.Date.Year & ""
        Return sql.EjecutarConsulta("d", cadena).Tables(0)
    End Function
    Public Function ListaCompras(d As NMovimiento) As DataTable
        Dim cadena As String = " select TipoDocumento2,NumeroDocumento2,FechaDocumento,IdProveedor,NombreProveedor,IdAlmacen,TipoDocumento,NumeroDocumento from Movimiento "
        cadena += "  where  Tipodocumento='PE' and isnull(Situacion,'')='V' and IdProveedor='" & d.IdProveedor & "'"
        Return sql.EjecutarConsulta("d", cadena).Tables(0)
    End Function
    Public Function ListaCompras(texto As String) As DataTable
        Dim cadena As String = " select TipoDocumento2,NumeroDocumento2,FechaDocumento,IdProveedor,NombreProveedor,IdAlmacen,TipoDocumento,NumeroDocumento from Movimiento "
        cadena += "  where  Tipodocumento='PE' and isnull(Situacion,'')='V' and (IdProveedor like '%" & texto & "%' or NombreProveedor like '%" & texto & "%' or NumeroDocumento2 like '%" & texto & "%' )"
        Return sql.EjecutarConsulta("d", cadena).Tables(0)
    End Function
    Public Function ListaCompras(TipoDoc As String, texto As String) As DataTable
        Dim cadena As String = " select TipoDocumento2,NumeroDocumento2,FechaDocumento,IdProveedor,NombreProveedor,IdAlmacen,TipoDocumento,NumeroDocumento from Movimiento "
        cadena += "  where  Tipodocumento='PE' and isnull(Situacion,'')='V' and Tipodocumento2='" & TipoDoc & "' and NumeroDocumento2 like '%" & texto & "%'"
        Return sql.EjecutarConsulta("d", cadena).Tables(0)
    End Function

    Public Function RepMovimientoIngreso(almacen As String, FechaI As DateTime, FechaF As DateTime, Mon As String) As DataTable
        Dim params() As Object = {"@IdALmacen", "@fechaI", "@fechaF", "@Moneda"}
        Dim vals() As Object = {almacen, FechaI, FechaF, Mon}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar}
        Return sql.ProcedureSQL("Str_ComprasMovimiento_R", params, vals, tipoParametro, 4).Tables(0)
    End Function
    Public Function RepMovimientoIngresoDetalle(almacen As String, FechaI As DateTime, FechaF As DateTime, Mon As String) As DataTable
        Dim params() As Object = {"@IdALmacen", "@fechaI", "@fechaF", "@Moneda"}
        Dim vals() As Object = {almacen, FechaI, FechaF, Mon}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar}
        Return sql.ProcedureSQL("Str_ComprasMovimientoDetalle_R", params, vals, tipoParametro, 4).Tables(0)
    End Function

    Public Function Movimiento_Almacen(i As DateTime, f As DateTime, ai As String, af As String, d As DataTable, arti As DataTable) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF", "@IdAlmacenI", "@IdAlmacenF", "@IdMov", "@arti"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Structured, SqlDbType.Structured}
        Dim valores() As Object = {i, f, ai, af, d, arti}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_MovimientoAlmacen", parametros, valores, tipoParametro, 6).Tables(0)
        Return dt
    End Function


    Public Function Existe_Movimiento(d As NMovimiento) As Boolean
        Dim parametros() As Object = {"@idalmacen", "@tipodocumento", "@numerodocumento", "@idagencia"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char}
        Dim valores() As Object = {d.idalmacen, d.tipodocumento, d.numerodocumento, d.idagencia}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_Movimiento", parametros, valores, tipoParametro, 4)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function


    Public Function GuiaremisionCPE(TipoDocumento As String, numero As String) As DataSet
        Dim params() As Object = {"@Tipodocumento", "@Numerodocumento"}
        Dim vals() As Object = {TipoDocumento, numero}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Return sql.ProcedureSQL("Str_GuiaRemision", params, vals, tipoParametro, 2)
    End Function
    Public Function GuiaremisionPDF(TipoDocumento As String, numero As String) As DataTable
        Dim params() As Object = {"@Tipodocumento", "@Numerodocumento"}
        Dim vals() As Object = {TipoDocumento, numero}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar}
        Return sql.ProcedureSQL("Str_Guia_PDF", params, vals, tipoParametro, 2).Tables(0)
    End Function

    Public Function KardexValorado(fi As DateTime, ff As DateTime, docmovalmacen As DataTable, articulo As DataTable, mon As String, Optional idalmacen As String = Nothing) As DataTable
        Dim parametros() As Object = {"@FechaI", "@FechaF", "@DocMovAlmacen", "@articulo", "@Idmoneda", "@idalmacen"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Structured, SqlDbType.Structured, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {fi, ff, docmovalmacen, articulo, mon, idalmacen}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_KardexValorado", parametros, valores, tipoParametro, 6).Tables(0)
        Return dt
    End Function

    Public Function Lista_ingreso_confirmar(idalmacen As String, fi As DateTime, ff As DateTime, tipo As String) As DataTable
        Dim parametros() As Object = {"@idalmacen", "@FechaI", "@FechaF", "@Tipo"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char}
        Dim valores() As Object = {idalmacen, fi, ff, tipo}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Lista_ingreso_confirmar", parametros, valores, tipoParametro, 4).Tables(0)
        Return dt
    End Function

    Public Function Existe_mov_Referencia(idalmacen As String, tipodocumento As String, numerodocumento As String) As DataTable
        Dim parametros() As Object = {"@idalmacen", "@Tipodocumento", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {idalmacen, tipodocumento, numerodocumento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Existe_mov_Referencia", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function

#End Region


End Class
