Imports CapaDatos
Public Class Ntbl_Liquidacion_Caja
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idcajaventa As Long
    Public Property idtipodocumento As String
    Public Property serie As String
    Public Property numerodocumento As String
    Public Property estado As String
    Public Property tipocambio As Decimal
    Public Property idmoneda As String
    Public Property importepago As Decimal
    Public Property importepagomn As Decimal
    Public Property tipooperacion As String
    Public Property idmonedapago As String
    Public Property nrodocref As String
    Public Property tipodocref As String
    Public Property tipo_origen As String
    Public Property fechacobro As System.DateTime

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As Ntbl_Liquidacion_Caja)

        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@estado", "@tipocambio", "@idmoneda", "@importepago", "@importepagomn", "@tipooperacion", "@idmonedapago", "@nrodocref", "@tipodocref", "@tipo_origen", "@fechacobro"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.estado, d.tipocambio, d.idmoneda, d.importepago, d.importepagomn, d.tipooperacion, d.idmonedapago, d.nrodocref, d.tipodocref, d.tipo_origen, d.fechacobro}
        sql.EjecutarProcedure("Str_Tbl_Liquidacion_Caja_I", parametros, valores, tipoParametro, 14)
    End Sub
    Public Sub Actualizar(d As Ntbl_Liquidacion_Caja)
        Dim parametros() As Object = {"@idCajaVenta", "@idtipodocumento", "@serie", "@numerodocumento", "@estado", "@tipocambio", "@idmoneda", "@importepago", "@importepagomn", "@tipooperacion", "@idmonedapago", "@nrodocref", "@tipodocref", "@tipo_origen", "@fechacobro"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime}
        Dim valores() As Object = {d.idcajaventa, d.idtipodocumento, d.serie, d.numerodocumento, d.estado, d.tipocambio, d.idmoneda, d.importepago, d.importepagomn, d.tipooperacion, d.idmonedapago, d.nrodocref, d.tipodocref, d.tipo_origen, d.fechacobro}
        sql.EjecutarProcedure("Str_Tbl_Liquidacion_Caja_U", parametros, valores, tipoParametro, 15)
    End Sub
    Public Function Agregar(d As Ntbl_Liquidacion_Caja, Retornatable As Boolean) As Ntbl_Liquidacion_Caja

        Dim parametros() As Object = {"@idtipodocumento", "@serie", "@numerodocumento", "@estado", "@tipocambio", "@idmoneda", "@importepago", "@importepagomn", "@tipooperacion", "@idmonedapago", "@nrodocref", "@tipodocref", "@tipo_origen", "@fechacobro"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime}
        Dim valores() As Object = {d.idtipodocumento, d.serie, d.numerodocumento, d.estado, d.tipocambio, d.idmoneda, d.importepago, d.importepagomn, d.tipooperacion, d.idmonedapago, d.nrodocref, d.tipodocref, d.tipo_origen, d.fechacobro}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Liquidacion_Caja_I_S", parametros, valores, tipoParametro, 14).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idcajaventa = IIf(dt.Rows(0).Item("idcajaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcajaventa"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.importepago = IIf(dt.Rows(0).Item("importepago") Is DBNull.Value, Nothing, dt.Rows(0).Item("importepago"))
            d.importepagomn = IIf(dt.Rows(0).Item("importepagomn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importepagomn"))
            d.tipooperacion = IIf(dt.Rows(0).Item("tipooperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipooperacion"))
            d.idmonedapago = IIf(dt.Rows(0).Item("idmonedapago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmonedapago"))
            d.nrodocref = IIf(dt.Rows(0).Item("nrodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocref"))
            d.tipodocref = IIf(dt.Rows(0).Item("tipodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocref"))
            d.tipo_origen = IIf(dt.Rows(0).Item("tipo_origen") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipo_origen"))
            d.fechacobro = IIf(dt.Rows(0).Item("fechacobro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacobro"))

        Else
            d.idcajaventa = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.estado = Nothing
            d.tipocambio = Nothing
            d.idmoneda = Nothing
            d.importepago = Nothing
            d.importepagomn = Nothing
            d.tipooperacion = Nothing
            d.idmonedapago = Nothing
            d.nrodocref = Nothing
            d.tipodocref = Nothing
            d.tipo_origen = Nothing
            d.fechacobro = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_Liquidacion_Caja, Retornatable As Boolean) As Ntbl_Liquidacion_Caja
        Dim parametros() As Object = {"@idCajaVenta", "@idtipodocumento", "@serie", "@numerodocumento", "@estado", "@tipocambio", "@idmoneda", "@importepago", "@importepagomn", "@tipooperacion", "@idmonedapago", "@nrodocref", "@tipodocref", "@tipo_origen", "@fechacobro"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.DateTime}
        Dim valores() As Object = {d.idcajaventa, d.idtipodocumento, d.serie, d.numerodocumento, d.estado, d.tipocambio, d.idmoneda, d.importepago, d.importepagomn, d.tipooperacion, d.idmonedapago, d.nrodocref, d.tipodocref, d.tipo_origen, d.fechacobro}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Liquidacion_Caja_U_S", parametros, valores, tipoParametro, 15).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idcajaventa = IIf(dt.Rows(0).Item("idcajaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcajaventa"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.importepago = IIf(dt.Rows(0).Item("importepago") Is DBNull.Value, Nothing, dt.Rows(0).Item("importepago"))
            d.importepagomn = IIf(dt.Rows(0).Item("importepagomn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importepagomn"))
            d.tipooperacion = IIf(dt.Rows(0).Item("tipooperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipooperacion"))
            d.idmonedapago = IIf(dt.Rows(0).Item("idmonedapago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmonedapago"))
            d.nrodocref = IIf(dt.Rows(0).Item("nrodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocref"))
            d.tipodocref = IIf(dt.Rows(0).Item("tipodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocref"))
            d.tipo_origen = IIf(dt.Rows(0).Item("tipo_origen") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipo_origen"))
            d.fechacobro = IIf(dt.Rows(0).Item("fechacobro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacobro"))
        Else
            d.idcajaventa = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.estado = Nothing
            d.tipocambio = Nothing
            d.idmoneda = Nothing
            d.importepago = Nothing
            d.importepagomn = Nothing
            d.tipooperacion = Nothing
            d.idmonedapago = Nothing
            d.nrodocref = Nothing
            d.tipodocref = Nothing
            d.tipo_origen = Nothing
            d.fechacobro = Nothing

        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_Liquidacion_Caja)
        Dim parametros() As Object = {"@idcajaventa"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idcajaventa}
        sql.EjecutarProcedure("Str_Tbl_Liquidacion_Caja_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idcajaventa"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Liquidacion_Caja_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_Liquidacion_Caja) As DataTable
        Dim parametros() As Object = {"@idcajaventa"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idcajaventa}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Liquidacion_Caja_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(idtd As String, serie As String, nro As String, tipocobro As String) As DataTable
        Dim dt As New DataTable
        Dim c As String = " SELECT     IdCajaVenta, IdTipoDocumento, Serie, NumeroDocumento, Estado, TipoCambio, IdMoneda, ImportePago, ImportePagoMN, TipoOperacion, IdMonedaPago, NroDocRef, TipoDocRef,FechaCobro "
        c += " FROM         Tbl_Liquidacion_Caja where IdTipoDocumento='" & idtd & "' and serie='" & serie & "' AND NUMERODOCUMENTO='" & nro & "' and Tipo_Origen='" & tipocobro & "'"
        dt = sql.EjecutarConsulta("d", c).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Ntbl_Liquidacion_Caja) As Ntbl_Liquidacion_Caja
        Dim parametros() As Object = {"@idcajaventa"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idcajaventa}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Liquidacion_Caja_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idcajaventa = IIf(dt.Rows(0).Item("idcajaventa") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcajaventa"))
            d.idtipodocumento = IIf(dt.Rows(0).Item("idtipodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idtipodocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numerodocumento = IIf(dt.Rows(0).Item("numerodocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numerodocumento"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.tipocambio = IIf(dt.Rows(0).Item("tipocambio") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipocambio"))
            d.idmoneda = IIf(dt.Rows(0).Item("idmoneda") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmoneda"))
            d.importepago = IIf(dt.Rows(0).Item("importepago") Is DBNull.Value, Nothing, dt.Rows(0).Item("importepago"))
            d.importepagomn = IIf(dt.Rows(0).Item("importepagomn") Is DBNull.Value, Nothing, dt.Rows(0).Item("importepagomn"))
            d.tipooperacion = IIf(dt.Rows(0).Item("tipooperacion") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipooperacion"))
            d.idmonedapago = IIf(dt.Rows(0).Item("idmonedapago") Is DBNull.Value, Nothing, dt.Rows(0).Item("idmonedapago"))
            d.nrodocref = IIf(dt.Rows(0).Item("nrodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocref"))
            d.tipodocref = IIf(dt.Rows(0).Item("tipodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocref"))
            d.tipo_origen = IIf(dt.Rows(0).Item("tipo_origen") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipo_origen"))
            d.fechacobro = IIf(dt.Rows(0).Item("fechacobro") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacobro"))
        Else
            d.idcajaventa = Nothing
            d.idtipodocumento = Nothing
            d.serie = Nothing
            d.numerodocumento = Nothing
            d.estado = Nothing
            d.tipocambio = Nothing
            d.idmoneda = Nothing
            d.importepago = Nothing
            d.importepagomn = Nothing
            d.tipooperacion = Nothing
            d.idmonedapago = Nothing
            d.nrodocref = Nothing
            d.tipodocref = Nothing
            d.tipo_origen = Nothing
            d.fechacobro = Nothing
        End If
        Return d
    End Function


#End Region



End Class
