Imports CapaDatos
Public Class NContado_Compras
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property identidad As Long
    Public Property idagencia As String
    Public Property idalmacen As String
    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numerodocumento As String
    Public Property idproveedor As String
    Public Property tipodocumento_pago As String
    Public Property numerodocumento_pago As String
    Public Property idmoneda As String
    Public Property debehaber As String
    Public Property estado As String
    Public Property glosa As String
    Public Property tipocambio As Decimal
    Public Property pago As Decimal
    Public Property pagous As Decimal
    Public Property pagomn As Decimal
    Public Property usuariocrea As String
    Public Property fechacrea As System.DateTime
    Public Property usuariomod As String
    Public Property fechamod As System.DateTime
    Public Property idmoneda_pago As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NContado_Compras)

        Dim parametros() As Object = {"@idagencia", "@idalmacen", "@idtipodocumento", "@serie", "@numerodocumento", "@idproveedor", "@tipodocumento_pago", "@numerodocumento_pago", "@idmoneda", "@debehaber", "@estado", "@glosa", "@tipocambio", "@pago", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@idmoneda_pago"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idalmacen, d.idtipodocumento, d.serie, d.numerodocumento, d.idproveedor, d.tipodocumento_pago, d.numerodocumento_pago, d.idmoneda, d.debehaber, d.estado, d.glosa, d.tipocambio, d.pago, d.usuariocrea, d.fechacrea, d.usuariomod, d.fechamod, d.idmoneda_pago}
        sql.EjecutarProcedure("Str_Contado_Compras_I", parametros, valores, tipoParametro, 19)
    End Sub
    Public Sub Actualizar(d As NContado_Compras)
        Dim parametros() As Object = {"@idagencia", "@idalmacen", "@idtipodocumento", "@serie", "@numerodocumento", "@idproveedor", "@tipodocumento_pago", "@numerodocumento_pago", "@idmoneda", "@debehaber", "@estado", "@glosa", "@tipocambio", "@pago", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@idmoneda_pago"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idalmacen, d.idtipodocumento, d.serie, d.numerodocumento, d.idproveedor, d.tipodocumento_pago, d.numerodocumento_pago, d.idmoneda, d.debehaber, d.estado, d.glosa, d.tipocambio, d.pago, d.usuariocrea, d.fechacrea, d.usuariomod, d.fechamod, d.idmoneda_pago}
        sql.EjecutarProcedure("Str_Contado_Compras_U", parametros, valores, tipoParametro, 19)
    End Sub
    Public Function Agregar(d As NContado_Compras, Retornatable As Boolean) As NContado_Compras

        Dim parametros() As Object = {"@idagencia", "@idalmacen", "@idtipodocumento", "@serie", "@numerodocumento", "@idproveedor", "@tipodocumento_pago", "@numerodocumento_pago", "@idmoneda", "@debehaber", "@estado", "@glosa", "@tipocambio", "@pago", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@idmoneda_pago"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idagencia, d.idalmacen, d.idtipodocumento, d.serie, d.numerodocumento, d.idproveedor, d.tipodocumento_pago, d.numerodocumento_pago, d.idmoneda, d.debehaber, d.estado, d.glosa, d.tipocambio, d.pago, d.usuariocrea, d.fechacrea, d.usuariomod, d.fechamod, d.idmoneda_pago}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Contado_Compras_I_S", parametros, valores, tipoParametro, 19).Tables(0)
        If dt.Rows.Count > 0 Then
            d.identidad = IIf(dt.Rows(0).Item("identidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("identidad"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.tipodocumento_pago = IIf(dt.Rows(0).Item("tipodocumento_pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocumento_pago"))
            d.numerodocumento_pago = IIf(dt.Rows(0).Item("numerodocumento_pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento_pago"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.debehaber = IIf(dt.Rows(0).Item("debehaber") Is DBNull.Value, Nothing, dt.Rows(0).Item("debehaber"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.pago = IIf(dt.Rows(0).Item("pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("pago"))
            d.pagous = IIf(dt.Rows(0).Item("pagous") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagous"))
            d.pagomn = IIf(dt.Rows(0).Item("pagomn") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagomn"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.idmoneda_pago = IIf(dt.Rows(0).Item("idmoneda_pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda_pago"))
        Else
            d.identidad = Nothing
            d.idagencia = Nothing
            d.idalmacen = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.idproveedor = Nothing
            d.tipodocumento_pago = Nothing
            d.numerodocumento_pago = Nothing
            d.idmoneda = Nothing
            d.debehaber = Nothing
            d.estado = Nothing
            d.glosa = Nothing
            d.tipocambio = Nothing
            d.pago = Nothing
            d.pagous = Nothing
            d.pagomn = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.usuariomod = Nothing
            d.fechamod = Nothing
            d.idmoneda_pago = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NContado_Compras, Retornatable As Boolean) As NContado_Compras

        Dim parametros() As Object = {"@idagencia", "@idalmacen", "@idtipodocumento", "@serie", "@numerodocumento", "@idproveedor", "@tipodocumento_pago", "@numerodocumento_pago", "@idmoneda", "@debehaber", "@estado", "@glosa", "@tipocambio", "@pago", "@usuariocrea", "@fechacrea", "@usuariomod", "@fechamod", "@idmoneda_pago"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.VarChar}
        Dim valores() As Object = {d.idmoneda_pago = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Contado_Compras_U_S", parametros, valores, tipoParametro, 63).Tables(0)
        If dt.Rows.Count > 0 Then
            d.identidad = IIf(dt.Rows(0).Item("identidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("identidad"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.tipodocumento_pago = IIf(dt.Rows(0).Item("tipodocumento_pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocumento_pago"))
            d.numerodocumento_pago = IIf(dt.Rows(0).Item("numerodocumento_pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento_pago"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.debehaber = IIf(dt.Rows(0).Item("debehaber") Is DBNull.Value, Nothing, dt.Rows(0).Item("debehaber"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.pago = IIf(dt.Rows(0).Item("pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("pago"))
            d.pagous = IIf(dt.Rows(0).Item("pagous") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagous"))
            d.pagomn = IIf(dt.Rows(0).Item("pagomn") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagomn"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.idmoneda_pago = IIf(dt.Rows(0).Item("idmoneda_pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda_pago"))
        Else
            d.identidad = Nothing
            d.idagencia = Nothing
            d.idalmacen = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.idproveedor = Nothing
            d.tipodocumento_pago = Nothing
            d.numerodocumento_pago = Nothing
            d.idmoneda = Nothing
            d.debehaber = Nothing
            d.estado = Nothing
            d.glosa = Nothing
            d.tipocambio = Nothing
            d.pago = Nothing
            d.pagous = Nothing
            d.pagomn = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.usuariomod = Nothing
            d.fechamod = Nothing
            d.idmoneda_pago = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NContado_Compras)
        Dim parametros() As Object = {"@identidad"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.identidad}
        sql.EjecutarProcedure("Str_Contado_Compras_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_Contado_Compras(d As NContado_Compras) As Boolean
        Dim parametros() As Object = {"@identidad"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.identidad}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_Contado_Compras", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@identidad"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Contado_Compras_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NContado_Compras) As DataTable
        Dim parametros() As Object = {"@identidad"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.identidad}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Contado_Compras_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NContado_Compras) As NContado_Compras
        Dim parametros() As Object = {"@identidad"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.identidad}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Contado_Compras_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.identidad = IIf(dt.Rows(0).Item("identidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("identidad"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.idproveedor = IIf(dt.Rows(0).Item("idproveedor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproveedor"))
            d.tipodocumento_pago = IIf(dt.Rows(0).Item("tipodocumento_pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocumento_pago"))
            d.numerodocumento_pago = IIf(dt.Rows(0).Item("numerodocumento_pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento_pago"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.debehaber = IIf(dt.Rows(0).Item("debehaber") Is DBNull.Value, Nothing, dt.Rows(0).Item("debehaber"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.pago = IIf(dt.Rows(0).Item("pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("pago"))
            d.pagous = IIf(dt.Rows(0).Item("pagous") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagous"))
            d.pagomn = IIf(dt.Rows(0).Item("pagomn") Is DBNull.Value, Nothing, dt.Rows(0).Item("pagomn"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariomod = IIf(dt.Rows(0).Item("usuariomod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariomod"))
            d.fechamod = IIf(dt.Rows(0).Item("fechamod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechamod"))
            d.idmoneda_pago = IIf(dt.Rows(0).Item("idmoneda_pago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda_pago"))
        Else
            d.identidad = Nothing
            d.idagencia = Nothing
            d.idalmacen = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.idproveedor = Nothing
            d.tipodocumento_pago = Nothing
            d.numerodocumento_pago = Nothing
            d.idmoneda = Nothing
            d.debehaber = Nothing
            d.estado = Nothing
            d.glosa = Nothing
            d.tipocambio = Nothing
            d.pago = Nothing
            d.pagous = Nothing
            d.pagomn = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
            d.usuariomod = Nothing
            d.fechamod = Nothing
            d.idmoneda_pago = Nothing
        End If
        Return d
    End Function
#End Region


End Class
