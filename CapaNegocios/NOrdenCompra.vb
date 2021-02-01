Imports CapaDatos
Public Class NOrdenCompra
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idoc As String
    Public Property idagencia As String
    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numerodocumento As String
    Public Property fechadocumento As System.DateTime
    Public Property consinigv As String
    Public Property igv As Decimal
    Public Property idproveedor As String
    Public Property nombreproveedor As String
    Public Property direccion As String
    Public Property ruc As String
    Public Property idalmacen As String
    Public Property idformaventa As String
    Public Property idmoneda As String
    Public Property tipocambio As Decimal
    Public Property importeigv As Decimal
    Public Property importetotal As Decimal
    Public Property idtipodocumento1 As String
    Public Property serie1 As String
    Public Property numerodocumento2 As String
    Public Property observacion As String
    Public Property estado As String
    Public Property usuariocrea As String
    Public Property fechacrea As System.DateTime
    Public Property usuariomod As String
    Public Property fechamod As System.DateTime
    Public Property importetotalmn As Decimal
    Public Property importetotalus As Decimal
    Public Property importeigvmn As Decimal
    Public Property importeigvus As Decimal
    Public Property contacto As String
    Public Property email As String
    Public Property fechaentrega As System.DateTime
    Public Property lugarentrega As String
    Public Property enviodocumento As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NOrdenCompra)

        Dim parametros() As Object = {"@idoc", "@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@fechadocumento", "@consinigv", "@igv", "@idproveedor", "@nombreproveedor", "@direccion", "@ruc", "@idalmacen", "@idformaventa", "@idmoneda", "@tipocambio", "@importeigv", "@importetotal", "@idtipodocumento1", "@serie1", "@numerodocumento2", "@observacion", "@estado", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@contacto", "@email", "@fechaentrega", "@lugarentrega", "@enviodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idoc, d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.fechadocumento, d.consinigv, d.igv, d.idproveedor, d.nombreproveedor, d.direccion, d.ruc, d.idalmacen, d.idformaventa, d.idmoneda, d.tipocambio, d.importeigv, d.importetotal, d.idtipodocumento1, d.serie1, d.numerodocumento2, d.observacion, d.estado, d.usuariocrea, d.fechacrea, d.usuariomod, d.fechamod, d.contacto, d.email, d.fechaentrega, d.lugarentrega, d.enviodocumento}
        sql.EjecutarProcedure("Str_OrdenCompra_I", parametros, valores, tipoParametro, 32)
    End Sub
    Public Sub Actualizar(d As NOrdenCompra)
        Dim parametros() As Object = {"@idoc", "@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@fechadocumento", "@consinigv", "@igv", "@idproveedor", "@nombreproveedor", "@direccion", "@ruc", "@idalmacen", "@idformaventa", "@idmoneda", "@tipocambio", "@importeigv", "@importetotal", "@idtipodocumento1", "@serie1", "@numerodocumento2", "@observacion", "@estado", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@contacto", "@email", "@fechaentrega", "@lugarentrega", "@enviodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idoc, d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.fechadocumento, d.consinigv, d.igv, d.idproveedor, d.nombreproveedor, d.direccion, d.ruc, d.idalmacen, d.idformaventa, d.idmoneda, d.tipocambio, d.importeigv, d.importetotal, d.idtipodocumento1, d.serie1, d.numerodocumento2, d.observacion, d.estado, d.usuariocrea, d.fechacrea, d.usuariomod, d.fechamod, d.contacto, d.email, d.fechaentrega, d.lugarentrega, d.enviodocumento}
        sql.EjecutarProcedure("Str_OrdenCompra_U", parametros, valores, tipoParametro, 32)
    End Sub
    Public Function Agregar(d As NOrdenCompra, Retornatable As Boolean) As NOrdenCompra

        Dim parametros() As Object = {"@idoc", "@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@fechadocumento", "@consinigv", "@igv", "@idproveedor", "@nombreproveedor", "@direccion", "@ruc", "@idalmacen", "@idformaventa", "@idmoneda", "@tipocambio", "@importeigv", "@importetotal", "@idtipodocumento1", "@serie1", "@numerodocumento2", "@observacion", "@estado", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@contacto", "@email", "@fechaentrega", "@lugarentrega", "@enviodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idoc, d.idagencia, d.idtipodocumento, d.serie, d.numerodocumento, d.fechadocumento, d.consinigv, d.igv, d.idproveedor, d.nombreproveedor, d.direccion, d.ruc, d.idalmacen, d.idformaventa, d.idmoneda, d.tipocambio, d.importeigv, d.importetotal, d.idtipodocumento1, d.serie1, d.numerodocumento2, d.observacion, d.estado, d.usuariocrea, d.fechacrea, d.usuariomod, d.fechamod, d.contacto, d.email, d.fechaentrega, d.lugarentrega, d.enviodocumento}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_OrdenCompra_I_S", parametros, valores, tipoParametro, 32).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idoc = IIf(dt.Rows(0).Item("idoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("idoc"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.consinigv = IIf(dt.Rows(0).Item("consinigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("consinigv"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.nombreproveedor = IIf(dt.Rows(0).Item("nombreproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreproveedor"))
            d.direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.ruc = IIf(dt.Rows(0).Item("ruc") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruc"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idformaventa = IIf(dt.Rows(0).Item("idformaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idformaventa"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.importetotal = IIf(dt.Rows(0).Item("importetotal") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotal"))
            d.idtipodocumento1 = IIf(dt.Rows(0).Item("idtipodocumento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento1"))
            d.serie1 = IIf(dt.Rows(0).Item("serie1") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie1"))
            d.numerodocumento2 = IIf(dt.Rows(0).Item("numerodocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento2"))
            d.observacion = IIf(dt.Rows(0).Item("observacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("observacion"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.importetotalmn = IIf(dt.Rows(0).Item("importetotalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalmn"))
            d.importetotalus = IIf(dt.Rows(0).Item("importetotalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalus"))
            d.importeigvmn = IIf(dt.Rows(0).Item("importeigvmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvmn"))
            d.importeigvus = IIf(dt.Rows(0).Item("importeigvus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvus"))
            d.contacto = IIf(dt.Rows(0).Item("contacto") Is DBNull.Value, Nothing, dt.Rows(0).Item("contacto"))
            d.email = IIf(dt.Rows(0).Item("email") Is DBNull.Value, Nothing, dt.Rows(0).Item("email"))
            d.fechaentrega = IIf(dt.Rows(0).Item("fechaentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaentrega"))
            d.lugarentrega = IIf(dt.Rows(0).Item("lugarentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarentrega"))
            d.enviodocumento = IIf(dt.Rows(0).Item("enviodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("enviodocumento"))
        Else
            d.idoc = Nothing
            d.idagencia = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.fechadocumento = Nothing
            d.consinigv = Nothing
            d.igv = Nothing
            d.idproveedor = Nothing
            d.nombreproveedor = Nothing
            d.direccion = Nothing
            d.ruc = Nothing
            d.idalmacen = Nothing
            d.idformaventa = Nothing
            d.idmoneda = Nothing
            d.tipocambio = Nothing
            d.importeigv = Nothing
            d.importetotal = Nothing
            d.idtipodocumento1 = Nothing
            d.serie1 = Nothing
            d.numerodocumento2 = Nothing
            d.observacion = Nothing
            d.estado = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.usuariomod = Nothing
            d.fechamod = Nothing
            d.importetotalmn = Nothing
            d.importetotalus = Nothing
            d.importeigvmn = Nothing
            d.importeigvus = Nothing
            d.contacto = Nothing
            d.email = Nothing
            d.fechaentrega = Nothing
            d.lugarentrega = Nothing
            d.enviodocumento = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NOrdenCompra, Retornatable As Boolean) As NOrdenCompra

        Dim parametros() As Object = {"@idoc", "@idagencia", "@idtipodocumento", "@serie", "@numerodocumento", "@fechadocumento", "@consinigv", "@igv", "@idproveedor", "@nombreproveedor", "@direccion", "@ruc", "@idalmacen", "@idformaventa", "@idmoneda", "@tipocambio", "@importeigv", "@importetotal", "@idtipodocumento1", "@serie1", "@numerodocumento2", "@observacion", "@estado", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@contacto", "@email", "@fechaentrega", "@lugarentrega", "@enviodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.enviodocumento = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_OrdenCompra_U_S", parametros, valores, tipoParametro, 104).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idoc = IIf(dt.Rows(0).Item("idoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("idoc"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.consinigv = IIf(dt.Rows(0).Item("consinigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("consinigv"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.nombreproveedor = IIf(dt.Rows(0).Item("nombreproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreproveedor"))
            d.direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.ruc = IIf(dt.Rows(0).Item("ruc") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruc"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idformaventa = IIf(dt.Rows(0).Item("idformaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idformaventa"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.importetotal = IIf(dt.Rows(0).Item("importetotal") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotal"))
            d.idtipodocumento1 = IIf(dt.Rows(0).Item("idtipodocumento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento1"))
            d.serie1 = IIf(dt.Rows(0).Item("serie1") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie1"))
            d.numerodocumento2 = IIf(dt.Rows(0).Item("numerodocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento2"))
            d.observacion = IIf(dt.Rows(0).Item("observacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("observacion"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.importetotalmn = IIf(dt.Rows(0).Item("importetotalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalmn"))
            d.importetotalus = IIf(dt.Rows(0).Item("importetotalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalus"))
            d.importeigvmn = IIf(dt.Rows(0).Item("importeigvmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvmn"))
            d.importeigvus = IIf(dt.Rows(0).Item("importeigvus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvus"))
            d.contacto = IIf(dt.Rows(0).Item("contacto") Is DBNull.Value, Nothing, dt.Rows(0).Item("contacto"))
            d.email = IIf(dt.Rows(0).Item("email") Is DBNull.Value, Nothing, dt.Rows(0).Item("email"))
            d.fechaentrega = IIf(dt.Rows(0).Item("fechaentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaentrega"))
            d.lugarentrega = IIf(dt.Rows(0).Item("lugarentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarentrega"))
            d.enviodocumento = IIf(dt.Rows(0).Item("enviodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("enviodocumento"))
        Else
            d.idoc = Nothing
            d.idagencia = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.fechadocumento = Nothing
            d.consinigv = Nothing
            d.igv = Nothing
            d.idproveedor = Nothing
            d.nombreproveedor = Nothing
            d.direccion = Nothing
            d.ruc = Nothing
            d.idalmacen = Nothing
            d.idformaventa = Nothing
            d.idmoneda = Nothing
            d.tipocambio = Nothing
            d.importeigv = Nothing
            d.importetotal = Nothing
            d.idtipodocumento1 = Nothing
            d.serie1 = Nothing
            d.numerodocumento2 = Nothing
            d.observacion = Nothing
            d.estado = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.usuariomod = Nothing
            d.fechamod = Nothing
            d.importetotalmn = Nothing
            d.importetotalus = Nothing
            d.importeigvmn = Nothing
            d.importeigvus = Nothing
            d.contacto = Nothing
            d.email = Nothing
            d.fechaentrega = Nothing
            d.lugarentrega = Nothing
            d.enviodocumento = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NOrdenCompra)


        Dim parametros() As Object = {"@idoc"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idoc}
        sql.EjecutarProcedure("Str_OrdenCompra_D", parametros, valores, tipoParametro, 1)
    End Sub
    'Public Function Existe_OrdenCompra(d As NOrdenCompra) As Boolean
    '    Dim parametros() As Object = {"@idoc"}
    '    Dim tipoParametro() As Object = {SqlDbType.VarChar}
    '    Dim valores() As Object = {d.idoc}
    '    Dim resultado As Integer
    '    Dim bandera As Boolean = False
    '    resultado = sql.procedimiento_escalar("Existe_OrdenCompra", parametros, valores, tipoParametro, 1)
    '    If resultado = 1 Then
    '        bandera = True
    '    Else
    '        bandera = False
    '    End If
    '    Return bandera
    'End Function
    Public Function Existe_OrdenCompra(d As NOrdenCompra) As Boolean
        Dim parametros() As Object = {"@idTipoDocumento", "@Serie", "@numerodocumento"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_OrdenCompra", parametros, valores, tipoParametro, 3)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idoc"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_OrdenCompra_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NOrdenCompra) As DataTable
        Dim parametros() As Object = {"@idoc"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idoc}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_OrdenCompra_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NOrdenCompra) As NOrdenCompra
        Dim parametros() As Object = {"@idoc"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.idoc}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_OrdenCompra_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idoc = IIf(dt.Rows(0).Item("idoc") Is DBNull.Value, Nothing, dt.Rows(0).Item("idoc"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.fechadocumento = IIf(dt.Rows(0).Item("fechadocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechadocumento"))
            d.consinigv = IIf(dt.Rows(0).Item("consinigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("consinigv"))
            d.igv = IIf(dt.Rows(0).Item("igv") Is DBNull.Value, Nothing, dt.Rows(0).Item("igv"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.nombreproveedor = IIf(dt.Rows(0).Item("nombreproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreproveedor"))
            d.direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.ruc = IIf(dt.Rows(0).Item("ruc") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruc"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idformaventa = IIf(dt.Rows(0).Item("idformaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idformaventa"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.importeigv = IIf(dt.Rows(0).Item("importeigv") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigv"))
            d.importetotal = IIf(dt.Rows(0).Item("importetotal") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotal"))
            d.idtipodocumento1 = IIf(dt.Rows(0).Item("idtipodocumento1") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento1"))
            d.serie1 = IIf(dt.Rows(0).Item("serie1") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie1"))
            d.numerodocumento2 = IIf(dt.Rows(0).Item("numerodocumento2") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento2"))
            d.observacion = IIf(dt.Rows(0).Item("observacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("observacion"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.importetotalmn = IIf(dt.Rows(0).Item("importetotalmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalmn"))
            d.importetotalus = IIf(dt.Rows(0).Item("importetotalus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importetotalus"))
            d.importeigvmn = IIf(dt.Rows(0).Item("importeigvmn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvmn"))
            d.importeigvus = IIf(dt.Rows(0).Item("importeigvus") Is DBNull.Value, Nothing, dt.Rows(0).Item("importeigvus"))
            d.contacto = IIf(dt.Rows(0).Item("contacto") Is DBNull.Value, Nothing, dt.Rows(0).Item("contacto"))
            d.email = IIf(dt.Rows(0).Item("email") Is DBNull.Value, Nothing, dt.Rows(0).Item("email"))
            d.fechaentrega = IIf(dt.Rows(0).Item("fechaentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaentrega"))
            d.lugarentrega = IIf(dt.Rows(0).Item("lugarentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarentrega"))
            d.enviodocumento = IIf(dt.Rows(0).Item("enviodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("enviodocumento"))
        Else
            d.idoc = Nothing
            d.idagencia = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.fechadocumento = Nothing
            d.consinigv = Nothing
            d.igv = Nothing
            d.idproveedor = Nothing
            d.nombreproveedor = Nothing
            d.direccion = Nothing
            d.ruc = Nothing
            d.idalmacen = Nothing
            d.idformaventa = Nothing
            d.idmoneda = Nothing
            d.tipocambio = Nothing
            d.importeigv = Nothing
            d.importetotal = Nothing
            d.idtipodocumento1 = Nothing
            d.serie1 = Nothing
            d.numerodocumento2 = Nothing
            d.observacion = Nothing
            d.estado = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.usuariomod = Nothing
            d.fechamod = Nothing
            d.importetotalmn = Nothing
            d.importetotalus = Nothing
            d.importeigvmn = Nothing
            d.importeigvus = Nothing
            d.contacto = Nothing
            d.email = Nothing
            d.fechaentrega = Nothing
            d.lugarentrega = Nothing
            d.enviodocumento = Nothing
        End If
        Return d
    End Function
#End Region

End Class
