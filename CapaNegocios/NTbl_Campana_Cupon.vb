Imports CapaDatos
Public Class NTbl_Campana_Cupon
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property id As Long
    Public Property idcampana As String
    Public Property fechasorteo As System.DateTime
    Public Property fechainiciocampana As System.DateTime
    Public Property fechafinalcampana As System.DateTime
    Public Property lugarsorteo As String
    Public Property premiocampana As String
    Public Property fechacrea As System.DateTime
    Public Property usuariocrea As String
    Public Property estadocampana As String
    Public Property Factor As Decimal
    Public Property IdTipoDocumento As String
    Public Property Hora As String

#End Region
#Region "Constructors"
    Public Sub New()
    End Sub
#End Region
#Region "Metodos"
    Public Sub Agregar(d As NTbl_Campana_Cupon)

        Dim parametros() As Object = {"@idcampana", "@fechasorteo", "@fechainiciocampana", "@fechafinalcampana", "@lugarsorteo", "@premiocampana", "@fechacrea", "@usuariocrea", "@estadocampana", "@Factor", "@IdTipoDocumento", "@Hora"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idcampana, d.fechasorteo, d.fechainiciocampana, d.fechafinalcampana, d.lugarsorteo, d.premiocampana, d.fechacrea, d.usuariocrea, d.estadocampana, d.Factor, d.IdTipoDocumento, d.Hora}
        sql.EjecutarProcedure("Str_Tbl_Campana_Cupon_I", parametros, valores, tipoParametro, 12)
    End Sub
    Public Sub Actualizar(d As NTbl_Campana_Cupon)
        Dim parametros() As Object = {"@id", "@idcampana", "@fechasorteo", "@fechainiciocampana", "@fechafinalcampana", "@lugarsorteo", "@premiocampana", "@fechacrea", "@usuariocrea", "@estadocampana", "Factor", "@IdTipoDocumento", "@Hora"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.id, d.idcampana, d.fechasorteo, d.fechainiciocampana, d.fechafinalcampana, d.lugarsorteo, d.premiocampana, d.fechacrea, d.usuariocrea, d.estadocampana, d.Factor, d.IdTipoDocumento, d.Hora}
        sql.EjecutarProcedure("Str_Tbl_Campana_Cupon_U", parametros, valores, tipoParametro, 13)
    End Sub
    Public Function Agregar(d As NTbl_Campana_Cupon, Retornatable As Boolean) As NTbl_Campana_Cupon
        Dim parametros() As Object = {"@idcampana", "@fechasorteo", "@fechainiciocampana", "@fechafinalcampana", "@lugarsorteo", "@premiocampana", "@fechacrea", "@usuariocrea", "@estadocampana", "Factor", "@IdTipoDocumento", "@Hora"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idcampana, d.fechasorteo, d.fechainiciocampana, d.fechafinalcampana, d.lugarsorteo, d.premiocampana, d.fechacrea, d.usuariocrea, d.estadocampana, d.Factor, d.IdTipoDocumento, d.Hora}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Campana_Cupon_I_S", parametros, valores, tipoParametro, 12).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.idcampana = IIf(dt.Rows(0).Item("idcampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcampana"))
            d.fechasorteo = IIf(dt.Rows(0).Item("fechasorteo") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechasorteo"))
            d.fechainiciocampana = IIf(dt.Rows(0).Item("fechainiciocampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechainiciocampana"))
            d.fechafinalcampana = IIf(dt.Rows(0).Item("fechafinalcampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechafinalcampana"))
            d.lugarsorteo = IIf(dt.Rows(0).Item("lugarsorteo") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarsorteo"))
            d.premiocampana = IIf(dt.Rows(0).Item("premiocampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("premiocampana"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.estadocampana = IIf(dt.Rows(0).Item("estadocampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadocampana"))
            d.Factor = IIf(dt.Rows(0).Item("Factor") Is DBNull.Value, Nothing, dt.Rows(0).Item("Factor"))
            d.IdTipoDocumento = IIf(dt.Rows(0).Item("IdTipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("IdTipoDocumento"))
            d.Hora = IIf(dt.Rows(0).Item("Hora") Is DBNull.Value, Nothing, dt.Rows(0).Item("Hora"))
        Else
            d.id = Nothing
            d.idcampana = Nothing
            d.fechasorteo = Nothing
            d.fechainiciocampana = Nothing
            d.fechafinalcampana = Nothing
            d.lugarsorteo = Nothing
            d.premiocampana = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
            d.estadocampana = Nothing
            d.Factor = Nothing
            d.IdTipoDocumento = Nothing
            d.Hora = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NTbl_Campana_Cupon, Retornatable As Boolean) As NTbl_Campana_Cupon
        Dim parametros() As Object = {"@id", "@idcampana", "@fechasorteo", "@fechainiciocampana", "@fechafinalcampana", "@lugarsorteo", "@premiocampana", "@fechacrea", "@usuariocrea", "@estadocampana", "Factor", "@IdTipoDocumento", "@Hora"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.id, d.idcampana, d.fechasorteo, d.fechainiciocampana, d.fechafinalcampana, d.lugarsorteo, d.premiocampana, d.fechacrea, d.usuariocrea, d.estadocampana, d.Factor, d.IdTipoDocumento, d.Hora}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Campana_Cupon_U_S", parametros, valores, tipoParametro, 13).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.idcampana = IIf(dt.Rows(0).Item("idcampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcampana"))
            d.fechasorteo = IIf(dt.Rows(0).Item("fechasorteo") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechasorteo"))
            d.fechainiciocampana = IIf(dt.Rows(0).Item("fechainiciocampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechainiciocampana"))
            d.fechafinalcampana = IIf(dt.Rows(0).Item("fechafinalcampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechafinalcampana"))
            d.lugarsorteo = IIf(dt.Rows(0).Item("lugarsorteo") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarsorteo"))
            d.premiocampana = IIf(dt.Rows(0).Item("premiocampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("premiocampana"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.estadocampana = IIf(dt.Rows(0).Item("estadocampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadocampana"))
            d.Factor = IIf(dt.Rows(0).Item("Factor") Is DBNull.Value, Nothing, dt.Rows(0).Item("Factor"))
            d.IdTipoDocumento = IIf(dt.Rows(0).Item("IdTipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("IdTipoDocumento"))
            d.Hora = IIf(dt.Rows(0).Item("Hora") Is DBNull.Value, Nothing, dt.Rows(0).Item("Hora"))
        Else
            d.id = Nothing
            d.idcampana = Nothing
            d.fechasorteo = Nothing
            d.fechainiciocampana = Nothing
            d.fechafinalcampana = Nothing
            d.lugarsorteo = Nothing
            d.premiocampana = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
            d.estadocampana = Nothing
            d.Factor = Nothing
            d.IdTipoDocumento = Nothing
            d.Hora = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NTbl_Campana_Cupon)
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        sql.EjecutarProcedure("Str_Tbl_Campana_Cupon_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Campana_Cupon_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTbl_Campana_Cupon) As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Campana_Cupon_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTbl_Campana_Cupon) As NTbl_Campana_Cupon
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Campana_Cupon_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.idcampana = IIf(dt.Rows(0).Item("idcampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("idcampana"))
            d.fechasorteo = IIf(dt.Rows(0).Item("fechasorteo") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechasorteo"))
            d.fechainiciocampana = IIf(dt.Rows(0).Item("fechainiciocampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechainiciocampana"))
            d.fechafinalcampana = IIf(dt.Rows(0).Item("fechafinalcampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechafinalcampana"))
            d.lugarsorteo = IIf(dt.Rows(0).Item("lugarsorteo") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarsorteo"))
            d.premiocampana = IIf(dt.Rows(0).Item("premiocampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("premiocampana"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.estadocampana = IIf(dt.Rows(0).Item("estadocampana") Is DBNull.Value, Nothing, dt.Rows(0).Item("estadocampana"))
            d.Factor = IIf(dt.Rows(0).Item("Factor") Is DBNull.Value, Nothing, dt.Rows(0).Item("Factor"))
            d.IdTipoDocumento = IIf(dt.Rows(0).Item("IdTipoDocumento") Is DBNull.Value, Nothing, dt.Rows(0).Item("IdTipoDocumento"))
            d.Hora = IIf(dt.Rows(0).Item("Hora") Is DBNull.Value, Nothing, dt.Rows(0).Item("Hora"))

        Else
            d.id = Nothing
            d.idcampana = Nothing
            d.fechasorteo = Nothing
            d.fechainiciocampana = Nothing
            d.fechafinalcampana = Nothing
            d.lugarsorteo = Nothing
            d.premiocampana = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
            d.estadocampana = Nothing
            d.Factor = Nothing
            d.IdTipoDocumento = Nothing
            d.Hora = Nothing
        End If
        Return d
    End Function

    ''' <summary>
    ''' Obtiene el ID de cupon
    ''' </summary>
    ''' <returns></returns>
    Public Function Get_IdCupon(idtipodocumento As String) As Integer
        Dim ca As String = " select Id from tbl_campana_cupon where EstadoCampana='V' and IdTipoDocumento='" & idtipodocumento & "'"
        Dim dt As New DataTable
        dt = sql.EjecutarConsulta("d", ca).Tables(0)
        Dim x As Integer = 0
        If dt.Rows.Count > 0 Then
            x = dt.Rows(0).Item(0)
        End If
        Return x
    End Function
#End Region
End Class
