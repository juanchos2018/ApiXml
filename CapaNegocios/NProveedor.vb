Imports CapaDatos
Public Class NProveedor
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property tipoanexo As String
    Public Property idproveedor As String
    Public Property nombre As String
    Public Property direccion As String
    Public Property localidad As String
    Public Property pais As String
    Public Property telefono1 As String
    Public Property telefono2 As String
    Public Property telefono3 As String
    Public Property fax As String
    Public Property tipoacreedor As String
    Public Property giro As String
    Public Property representante As String
    Public Property cargorepresentante As String
    Public Property telefonorep As String
    Public Property fechacrea As System.DateTime
    Public Property fechamod As System.DateTime
    Public Property usuariocrea As String
    Public Property estado As String
    Public Property abreviado As String
    Public Property ruc As String
    Public Property email As String
    Public Property idformapago As String
    Public Property idtipodocumentoiden As String
    Public Property numerodocumentoiden As String
    Public Property tipopersona As String
    Public Property tipoprocedencia As String
    Public Property incrementopublico As Decimal
    Public Property dsctoproveedor As Decimal
    Public Property tipodocsunat As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NProveedor)

        Dim parametros() As Object = {"@tipoanexo", "@idproveedor", "@nombre", "@direccion", "@localidad", "@pais", "@telefono1", "@telefono2", "@telefono3", "@fax", "@tipoacreedor", "@giro", "@representante", "@cargorepresentante", "@telefonorep", "@fechacrea", "@fechamod", "@usuariocrea", "@estado", "@abreviado", "@ruc", "@email", "@idformapago", "@idtipodocumentoiden", "@numerodocumentoiden", "@tipopersona", "@tipoprocedencia", "@incrementopublico", "@dsctoproveedor", "@tipodocsunat"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipoanexo, d.idproveedor, d.nombre, d.direccion, d.localidad, d.pais, d.telefono1, d.telefono2, d.telefono3, d.fax, d.tipoacreedor, d.giro, d.representante, d.cargorepresentante, d.telefonorep, d.fechacrea, d.fechamod, d.usuariocrea, d.estado, d.abreviado, d.ruc, d.email, d.idformapago, d.idtipodocumentoiden, d.numerodocumentoiden, d.tipopersona, d.tipoprocedencia, d.incrementopublico, d.dsctoproveedor, d.tipodocsunat}
        sql.EjecutarProcedure("Str_Proveedor_I", parametros, valores, tipoParametro, 30)
    End Sub
    Public Sub Actualizar(d As NProveedor)
        Dim parametros() As Object = {"@tipoanexo", "@idproveedor", "@nombre", "@direccion", "@localidad", "@pais", "@telefono1", "@telefono2", "@telefono3", "@fax", "@tipoacreedor", "@giro", "@representante", "@cargorepresentante", "@telefonorep", "@fechacrea", "@fechamod", "@usuariocrea", "@estado", "@abreviado", "@ruc", "@email", "@idformapago", "@idtipodocumentoiden", "@numerodocumentoiden", "@tipopersona", "@tipoprocedencia", "@incrementopublico", "@dsctoproveedor", "@tipodocsunat"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipoanexo, d.idproveedor, d.nombre, d.direccion, d.localidad, d.pais, d.telefono1, d.telefono2, d.telefono3, d.fax, d.tipoacreedor, d.giro, d.representante, d.cargorepresentante, d.telefonorep, d.fechacrea, d.fechamod, d.usuariocrea, d.estado, d.abreviado, d.ruc, d.email, d.idformapago, d.idtipodocumentoiden, d.numerodocumentoiden, d.tipopersona, d.tipoprocedencia, d.incrementopublico, d.dsctoproveedor, d.tipodocsunat}
        sql.EjecutarProcedure("Str_Proveedor_U", parametros, valores, tipoParametro, 30)
    End Sub
    Public Function Agregar(d As NProveedor, Retornatable As Boolean) As NProveedor

        Dim parametros() As Object = {"@tipoanexo", "@idproveedor", "@nombre", "@direccion", "@localidad", "@pais", "@telefono1", "@telefono2", "@telefono3", "@fax", "@tipoacreedor", "@giro", "@representante", "@cargorepresentante", "@telefonorep", "@fechacrea", "@fechamod", "@usuariocrea", "@estado", "@abreviado", "@ruc", "@email", "@idformapago", "@idtipodocumentoiden", "@numerodocumentoiden", "@tipopersona", "@tipoprocedencia", "@incrementopublico", "@dsctoproveedor", "@tipodocsunat"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipoanexo, d.idproveedor, d.nombre, d.direccion, d.localidad, d.pais, d.telefono1, d.telefono2, d.telefono3, d.fax, d.tipoacreedor, d.giro, d.representante, d.cargorepresentante, d.telefonorep, d.fechacrea, d.fechamod, d.usuariocrea, d.estado, d.abreviado, d.ruc, d.email, d.idformapago, d.idtipodocumentoiden, d.numerodocumentoiden, d.tipopersona, d.tipoprocedencia, d.incrementopublico, d.dsctoproveedor, d.tipodocsunat}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Proveedor_I_S", parametros, valores, tipoParametro, 30).Tables(0)
        If dt.Rows.Count > 0 Then
            d.tipoanexo = IIf(dt.Rows(0).Item("tipoanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoanexo"))

            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))

            d.nombre = IIf(dt.Rows(0).Item("nombre") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombre"))

            d.direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))

            d.localidad = IIf(dt.Rows(0).Item("localidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("localidad"))

            d.pais = IIf(dt.Rows(0).Item("pais") Is DBNull.Value, Nothing, dt.Rows(0).Item("pais"))

            d.telefono1 = IIf(dt.Rows(0).Item("telefono1") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono1"))

            d.telefono2 = IIf(dt.Rows(0).Item("telefono2") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono2"))

            d.telefono3 = IIf(dt.Rows(0).Item("telefono3") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono3"))

            d.fax = IIf(dt.Rows(0).Item("fax") Is DBNull.Value, Nothing, dt.Rows(0).Item("fax"))

            d.tipoacreedor = IIf(dt.Rows(0).Item("tipoacreedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoacreedor"))

            d.giro = IIf(dt.Rows(0).Item("giro") Is DBNull.Value, Nothing, dt.Rows(0).Item("giro"))

            d.representante = IIf(dt.Rows(0).Item("representante") Is DBNull.Value, Nothing, dt.Rows(0).Item("representante"))

            d.cargorepresentante = IIf(dt.Rows(0).Item("cargorepresentante") Is DBNull.Value, Nothing, dt.Rows(0).Item("cargorepresentante"))

            d.telefonorep = IIf(dt.Rows(0).Item("telefonorep") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefonorep"))

            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))

            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))

            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))

            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))

            d.abreviado = IIf(dt.Rows(0).Item("abreviado") Is DBNull.Value, Nothing, dt.Rows(0).Item("abreviado"))

            d.ruc = IIf(dt.Rows(0).Item("ruc") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruc"))

            d.email = IIf(dt.Rows(0).Item("email") Is DBNull.Value, Nothing, dt.Rows(0).Item("email"))

            d.idformapago = IIf(dt.Rows(0).Item("idformapago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idformapago"))

            d.idtipodocumentoiden = IIf(dt.Rows(0).Item("idtipodocumentoiden") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumentoiden"))

            d.numerodocumentoiden = IIf(dt.Rows(0).Item("numerodocumentoiden") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumentoiden"))

            d.tipopersona = IIf(dt.Rows(0).Item("tipopersona") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipopersona"))

            d.tipoprocedencia = IIf(dt.Rows(0).Item("tipoprocedencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoprocedencia"))

            d.incrementopublico = IIf(dt.Rows(0).Item("incrementopublico") Is DBNull.Value, Nothing, dt.Rows(0).Item("incrementopublico"))

            d.dsctoproveedor = IIf(dt.Rows(0).Item("dsctoproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("dsctoproveedor"))

            d.tipodocsunat = IIf(dt.Rows(0).Item("tipodocsunat") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocsunat"))

        Else
            d.tipoanexo = Nothing
            d.idproveedor = Nothing
            d.nombre = Nothing
            d.direccion = Nothing
            d.localidad = Nothing
            d.pais = Nothing
            d.telefono1 = Nothing
            d.telefono2 = Nothing
            d.telefono3 = Nothing
            d.fax = Nothing
            d.tipoacreedor = Nothing
            d.giro = Nothing
            d.representante = Nothing
            d.cargorepresentante = Nothing
            d.telefonorep = Nothing
            d.fechacrea = Nothing
            d.fechamod = Nothing
            d.usuariocrea = Nothing
            d.estado = Nothing
            d.abreviado = Nothing
            d.ruc = Nothing
            d.email = Nothing
            d.idformapago = Nothing
            d.idtipodocumentoiden = Nothing
            d.numerodocumentoiden = Nothing
            d.tipopersona = Nothing
            d.tipoprocedencia = Nothing
            d.incrementopublico = Nothing
            d.dsctoproveedor = Nothing
            d.tipodocsunat = Nothing

        End If
        Return d
    End Function
    Public Function Actualizar(d As NProveedor, Retornatable As Boolean) As NProveedor
        Dim parametros() As Object = {"@tipoanexo", "@idproveedor", "@nombre", "@direccion", "@localidad", "@pais", "@telefono1", "@telefono2", "@telefono3", "@fax", "@tipoacreedor", "@giro", "@representante", "@cargorepresentante", "@telefonorep", "@fechacrea", "@fechamod", "@usuariocrea", "@estado", "@abreviado", "@ruc", "@email", "@idformapago", "@idtipodocumentoiden", "@numerodocumentoiden", "@tipopersona", "@tipoprocedencia", "@incrementopublico", "@dsctoproveedor", "@tipodocsunat"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar}
        Dim valores() As Object = {d.tipoanexo, d.idproveedor, d.nombre, d.direccion, d.localidad, d.pais, d.telefono1, d.telefono2, d.telefono3, d.fax, d.tipoacreedor, d.giro, d.representante, d.cargorepresentante, d.telefonorep, d.fechacrea, d.fechamod, d.usuariocrea, d.estado, d.abreviado, d.ruc, d.email, d.idformapago, d.idtipodocumentoiden, d.numerodocumentoiden, d.tipopersona, d.tipoprocedencia, d.incrementopublico, d.dsctoproveedor, d.tipodocsunat}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Proveedor_U_S", parametros, valores, tipoParametro, 60).Tables(0)
        If dt.Rows.Count > 0 Then
            d.tipoanexo = IIf(dt.Rows(0).Item("tipoanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoanexo"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.nombre = IIf(dt.Rows(0).Item("nombre") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombre"))
            d.direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.localidad = IIf(dt.Rows(0).Item("localidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("localidad"))
            d.pais = IIf(dt.Rows(0).Item("pais") Is DBNull.Value, Nothing, dt.Rows(0).Item("pais"))
            d.telefono1 = IIf(dt.Rows(0).Item("telefono1") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono1"))
            d.telefono2 = IIf(dt.Rows(0).Item("telefono2") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono2"))
            d.telefono3 = IIf(dt.Rows(0).Item("telefono3") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono3"))
            d.fax = IIf(dt.Rows(0).Item("fax") Is DBNull.Value, Nothing, dt.Rows(0).Item("fax"))
            d.tipoacreedor = IIf(dt.Rows(0).Item("tipoacreedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoacreedor"))
            d.giro = IIf(dt.Rows(0).Item("giro") Is DBNull.Value, Nothing, dt.Rows(0).Item("giro"))
            d.representante = IIf(dt.Rows(0).Item("representante") Is DBNull.Value, Nothing, dt.Rows(0).Item("representante"))
            d.cargorepresentante = IIf(dt.Rows(0).Item("cargorepresentante") Is DBNull.Value, Nothing, dt.Rows(0).Item("cargorepresentante"))
            d.telefonorep = IIf(dt.Rows(0).Item("telefonorep") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefonorep"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.abreviado = IIf(dt.Rows(0).Item("abreviado") Is DBNull.Value, Nothing, dt.Rows(0).Item("abreviado"))
            d.ruc = IIf(dt.Rows(0).Item("ruc") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruc"))
            d.email = IIf(dt.Rows(0).Item("email") Is DBNull.Value, Nothing, dt.Rows(0).Item("email"))
            d.idformapago = IIf(dt.Rows(0).Item("idformapago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idformapago"))
            d.idtipodocumentoiden = IIf(dt.Rows(0).Item("idtipodocumentoiden") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumentoiden"))
            d.numerodocumentoiden = IIf(dt.Rows(0).Item("numerodocumentoiden") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumentoiden"))
            d.tipopersona = IIf(dt.Rows(0).Item("tipopersona") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipopersona"))
            d.tipoprocedencia = IIf(dt.Rows(0).Item("tipoprocedencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoprocedencia"))
            d.incrementopublico = IIf(dt.Rows(0).Item("incrementopublico") Is DBNull.Value, Nothing, dt.Rows(0).Item("incrementopublico"))
            d.dsctoproveedor = IIf(dt.Rows(0).Item("dsctoproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("dsctoproveedor"))
            d.tipodocsunat = IIf(dt.Rows(0).Item("tipodocsunat") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocsunat"))
        Else
            d.tipoanexo = Nothing
            d.idproveedor = Nothing
            d.nombre = Nothing
            d.direccion = Nothing
            d.localidad = Nothing
            d.pais = Nothing
            d.telefono1 = Nothing
            d.telefono2 = Nothing
            d.telefono3 = Nothing
            d.fax = Nothing
            d.tipoacreedor = Nothing
            d.giro = Nothing
            d.representante = Nothing
            d.cargorepresentante = Nothing
            d.telefonorep = Nothing
            d.fechacrea = Nothing
            d.fechamod = Nothing
            d.usuariocrea = Nothing
            d.estado = Nothing
            d.abreviado = Nothing
            d.ruc = Nothing
            d.email = Nothing
            d.idformapago = Nothing
            d.idtipodocumentoiden = Nothing
            d.numerodocumentoiden = Nothing
            d.tipopersona = Nothing
            d.tipoprocedencia = Nothing
            d.incrementopublico = Nothing
            d.dsctoproveedor = Nothing
            d.tipodocsunat = Nothing

        End If
        Return d
    End Function
    Public Sub Eliminar(d As NProveedor)
        Dim parametros() As Object = {"@idproveedor"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.idproveedor}
        sql.EjecutarProcedure("Str_Proveedor_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_Proveedor(d As NProveedor) As Boolean
        Dim parametros() As Object = {"@idproveedor"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.idproveedor}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_Proveedor", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idproveedor"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Proveedor_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NProveedor) As DataTable
        Dim parametros() As Object = {"@idproveedor"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.idproveedor}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Proveedor_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NProveedor) As NProveedor
        Dim parametros() As Object = {"@idproveedor"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.idproveedor}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Proveedor_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.tipoanexo = IIf(dt.Rows(0).Item("tipoanexo") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoanexo"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.nombre = IIf(dt.Rows(0).Item("nombre") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombre"))
            d.direccion = IIf(dt.Rows(0).Item("direccion") Is DBNull.Value, Nothing, dt.Rows(0).Item("direccion"))
            d.localidad = IIf(dt.Rows(0).Item("localidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("localidad"))
            d.pais = IIf(dt.Rows(0).Item("pais") Is DBNull.Value, Nothing, dt.Rows(0).Item("pais"))
            d.telefono1 = IIf(dt.Rows(0).Item("telefono1") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono1"))
            d.telefono2 = IIf(dt.Rows(0).Item("telefono2") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono2"))
            d.telefono3 = IIf(dt.Rows(0).Item("telefono3") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefono3"))
            d.fax = IIf(dt.Rows(0).Item("fax") Is DBNull.Value, Nothing, dt.Rows(0).Item("fax"))
            d.tipoacreedor = IIf(dt.Rows(0).Item("tipoacreedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoacreedor"))
            d.giro = IIf(dt.Rows(0).Item("giro") Is DBNull.Value, Nothing, dt.Rows(0).Item("giro"))
            d.representante = IIf(dt.Rows(0).Item("representante") Is DBNull.Value, Nothing, dt.Rows(0).Item("representante"))
            d.cargorepresentante = IIf(dt.Rows(0).Item("cargorepresentante") Is DBNull.Value, Nothing, dt.Rows(0).Item("cargorepresentante"))
            d.telefonorep = IIf(dt.Rows(0).Item("telefonorep") Is DBNull.Value, Nothing, dt.Rows(0).Item("telefonorep"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.abreviado = IIf(dt.Rows(0).Item("abreviado") Is DBNull.Value, Nothing, dt.Rows(0).Item("abreviado"))
            d.ruc = IIf(dt.Rows(0).Item("ruc") Is DBNull.Value, Nothing, dt.Rows(0).Item("ruc"))
            d.email = IIf(dt.Rows(0).Item("email") Is DBNull.Value, Nothing, dt.Rows(0).Item("email"))
            d.idformapago = IIf(dt.Rows(0).Item("idformapago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idformapago"))
            d.idtipodocumentoiden = IIf(dt.Rows(0).Item("idtipodocumentoiden") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumentoiden"))
            d.numerodocumentoiden = IIf(dt.Rows(0).Item("numerodocumentoiden") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumentoiden"))
            d.tipopersona = IIf(dt.Rows(0).Item("tipopersona") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipopersona"))
            d.tipoprocedencia = IIf(dt.Rows(0).Item("tipoprocedencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipoprocedencia"))
            d.incrementopublico = IIf(dt.Rows(0).Item("incrementopublico") Is DBNull.Value, Nothing, dt.Rows(0).Item("incrementopublico"))
            d.dsctoproveedor = IIf(dt.Rows(0).Item("dsctoproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("dsctoproveedor"))
            d.tipodocsunat = IIf(dt.Rows(0).Item("tipodocsunat") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocsunat"))
        Else
            d.tipoanexo = Nothing
            d.idproveedor = Nothing
            d.nombre = Nothing
            d.direccion = Nothing
            d.localidad = Nothing
            d.pais = Nothing
            d.telefono1 = Nothing
            d.telefono2 = Nothing
            d.telefono3 = Nothing
            d.fax = Nothing
            d.tipoacreedor = Nothing
            d.giro = Nothing
            d.representante = Nothing
            d.cargorepresentante = Nothing
            d.telefonorep = Nothing
            d.fechacrea = Nothing
            d.fechamod = Nothing
            d.usuariocrea = Nothing
            d.estado = Nothing
            d.abreviado = Nothing
            d.ruc = Nothing
            d.email = Nothing
            d.idformapago = Nothing
            d.idtipodocumentoiden = Nothing
            d.numerodocumentoiden = Nothing
            d.tipopersona = Nothing
            d.tipoprocedencia = Nothing
            d.incrementopublico = Nothing
            d.dsctoproveedor = Nothing
            d.tipodocsunat = Nothing

        End If
        Return d
    End Function
#End Region



    Public Function Lista_Proveedor() As DataTable
        Return sql.EjecutarConsulta("d", "select cast(0 as bit) as Flg,IdProveedor,Nombre from Proveedor ").Tables(0)
    End Function




End Class
