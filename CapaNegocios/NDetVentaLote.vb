Imports CapaDatos

Public Class NDetVentaLote
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idLote As Long
    Public Property idAgencia As String
    Public Property idAlmacen As String
    Public Property idTipoDocumento As String
    Public Property serie As String
    Public Property numeroDocumento As String
    Public Property idCliente As String
    Public Property idArticulo As String
    Public Property item As String
    Public Property idInstitucion As String
    Public Property cantidad As Decimal
    Public Property estado As String
    Public Property usuarioCrea As String
    Public Property fechaCrea As System.DateTime
    Public Property usuarioMod As String
    Public Property fechaMod As System.DateTime
    Public Property saldo As Decimal
    Public Property nroLote As String
    Public Property nroProceso As String
    Public Property comentario As String
#End Region
#Region "Constructors"
    Public Sub New()
    End Sub
#End Region


#Region "Metodos"
    Public Sub Agregar(d As NDetVentaLote)

        Dim parametros() As Object = {"@idAgencia", "@idAlmacen", "@idTipoDocumento", "@serie", "@numeroDocumento", "@idCliente", "@idArticulo", "@item", "@idInstitucion", "@cantidad", "@estado", "@usuarioCrea", "@fechaCrea", "@usuarioMod", "@fechaMod", "@saldo", "@nroLote", "@nroProceso", "@comentario"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idAgencia, d.idAlmacen, d.idTipoDocumento, d.serie, d.numeroDocumento, d.idCliente, d.idArticulo, d.item, d.idInstitucion, d.cantidad, d.estado, d.usuarioCrea, d.fechaCrea, d.usuarioMod, d.fechaMod, d.saldo, d.nroLote, d.nroProceso, d.comentario}
        Sql.EjecutarProcedure("Str_DetVenta_Lote_I", parametros, valores, tipoParametro, 19)
    End Sub
    Public Sub Actualizar(d As NDetVentaLote)
        Dim parametros() As Object = {"@idAgencia", "@idAlmacen", "@idTipoDocumento", "@serie", "@numeroDocumento", "@idCliente", "@idArticulo", "@item", "@idInstitucion", "@cantidad", "@estado", "@usuarioCrea", "@fechaCrea", "@usuarioMod", "@fechaMod", "@saldo", "@nroLote", "@nroProceso", "@comentario"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idAgencia, d.idAlmacen, d.idTipoDocumento, d.serie, d.numeroDocumento, d.idCliente, d.idArticulo, d.item, d.idInstitucion, d.cantidad, d.estado, d.usuarioCrea, d.fechaCrea, d.usuarioMod, d.fechaMod, d.saldo, d.nroLote, d.nroProceso, d.comentario}
        Sql.EjecutarProcedure("Str_DetVenta_Lote_U", parametros, valores, tipoParametro, 19)
    End Sub
    Public Sub Eliminar(d As NDetVentaLote)
        Dim parametros() As Object = {"@idLote"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idLote}
        Sql.EjecutarProcedure("Str_DetVenta_Lote_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idLote"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_DetVenta_Lote_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NDetVentaLote) As DataTable
        Dim parametros() As Object = {"@idLote"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idLote}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_DetVenta_Lote_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NDetVentaLote) As NDetVentaLote
        Dim parametros() As Object = {"@idLote"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idLote}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_DetVenta_Lote_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idLote = IIf(dt.Rows(0).Item("idLote") Is DBNull.Value, Nothing, dt.Rows(0).Item("idLote"))
            d.idAgencia = IIf(dt.Rows(0).Item("idAgencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAgencia"))
            d.idAlmacen = IIf(dt.Rows(0).Item("idAlmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idAlmacen"))
            d.idTipoDocumento = IIf(dt.Rows(0).Item("idTipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("idTipoDocumento"))
            d.serie = IIf(dt.Rows(0).Item("serie") Is DBNull.Value, Nothing, dt.Rows(0).Item("serie"))
            d.numeroDocumento = IIf(dt.Rows(0).Item("numeroDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("numeroDocumento"))
            d.idCliente = IIf(dt.Rows(0).Item("idCliente") Is DBNull.Value, Nothing, dt.Rows(0).Item("idCliente"))
            d.idArticulo = IIf(dt.Rows(0).Item("idArticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idArticulo"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idInstitucion = IIf(dt.Rows(0).Item("idInstitucion") Is DBNull.Value, Nothing, dt.Rows(0).Item("idInstitucion"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.usuarioCrea = IIf(dt.Rows(0).Item("usuarioCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioCrea"))
            d.fechaCrea = IIf(dt.Rows(0).Item("fechaCrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaCrea"))
            d.usuarioMod = IIf(dt.Rows(0).Item("usuarioMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuarioMod"))
            d.fechaMod = IIf(dt.Rows(0).Item("fechaMod") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechaMod"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.nroLote = IIf(dt.Rows(0).Item("nroLote") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroLote"))
            d.nroProceso = IIf(dt.Rows(0).Item("nroProceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroProceso"))
            d.comentario = IIf(dt.Rows(0).Item("comentario") Is DBNull.Value, Nothing, dt.Rows(0).Item("comentario"))
        Else
            d.idLote = Nothing
            d.idAgencia = Nothing
            d.idAlmacen = Nothing
            d.idTipoDocumento = Nothing
            d.serie = Nothing
            d.numeroDocumento = Nothing
            d.idCliente = Nothing
            d.idArticulo = Nothing
            d.item = Nothing
            d.idInstitucion = Nothing
            d.cantidad = Nothing
            d.estado = Nothing
            d.usuarioCrea = Nothing
            d.fechaCrea = Nothing
            d.usuarioMod = Nothing
            d.fechaMod = Nothing
            d.saldo = Nothing
            d.nroLote = Nothing
            d.nroProceso = Nothing
            d.comentario = Nothing
        End If
        Return d
    End Function
#End Region


End Class
