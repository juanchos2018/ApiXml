Imports CapaDatos
Public Class NTbl_Parametros_puntos
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property id As Long
    Public Property monto_puntos As Decimal
    Public Property fechainicio As System.DateTime
    Public Property frase As String
    Public Property fechacrea As System.DateTime
    Public Property usuariocrea As String
    Public Property isboleta As Boolean
    Public Property isfactura As Boolean
    Public Property estado As String
    Public Property RazonSocial As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NTbl_Parametros_puntos)
        Dim parametros() As Object = {"@monto_puntos", "@fechainicio", "@frase", "@fechacrea", "@usuariocrea", "@isboleta", "@isfactura", "@estado", "@RazonSocial"}
        Dim tipoParametro() As Object = {SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.monto_puntos, d.fechainicio, d.frase, d.fechacrea, d.usuariocrea, d.isboleta, d.isfactura, d.estado, d.RazonSocial}
        sql.EjecutarProcedure("Str_Tbl_Parametros_puntos_I", parametros, valores, tipoParametro, 9)
    End Sub
    Public Sub Actualizar(d As NTbl_Parametros_puntos)
        Dim parametros() As Object = {"@id", "@monto_puntos", "@fechainicio", "@frase", "@fechacrea", "@usuariocrea", "@isboleta", "@isfactura", "@estado", "@RazonSocial"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.id, d.monto_puntos, d.fechainicio, d.frase, d.fechacrea, d.usuariocrea, d.isboleta, d.isfactura, d.estado, d.RazonSocial}
        sql.EjecutarProcedure("Str_Tbl_Parametros_puntos_U", parametros, valores, tipoParametro, 10)
    End Sub
    Public Function Agregar(d As NTbl_Parametros_puntos, Retornatable As Boolean) As NTbl_Parametros_puntos
        Dim parametros() As Object = {"@monto_puntos", "@fechainicio", "@frase", "@fechacrea", "@usuariocrea", "@isboleta", "@isfactura", "@estado", "@RazonSocial"}
        Dim tipoParametro() As Object = {SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.monto_puntos, d.fechainicio, d.frase, d.fechacrea, d.usuariocrea, d.isboleta, d.isfactura, d.estado, d.RazonSocial}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Parametros_puntos_I_S", parametros, valores, tipoParametro, 9).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.monto_puntos = IIf(dt.Rows(0).Item("monto_puntos") Is DBNull.Value, Nothing, dt.Rows(0).Item("monto_puntos"))
            d.fechainicio = IIf(dt.Rows(0).Item("fechainicio") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechainicio"))
            d.frase = IIf(dt.Rows(0).Item("frase") Is DBNull.Value, Nothing, dt.Rows(0).Item("frase"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.isboleta = IIf(dt.Rows(0).Item("isboleta") Is DBNull.Value, Nothing, dt.Rows(0).Item("isboleta"))
            d.isfactura = IIf(dt.Rows(0).Item("isfactura") Is DBNull.Value, Nothing, dt.Rows(0).Item("isfactura"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.RazonSocial = IIf(dt.Rows(0).Item("RazonSocial") Is DBNull.Value, Nothing, dt.Rows(0).Item("RazonSocial"))
        Else
            d.id = Nothing
            d.monto_puntos = Nothing
            d.fechainicio = Nothing
            d.frase = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
            d.isboleta = Nothing
            d.isfactura = Nothing
            d.estado = Nothing
            d.RazonSocial = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NTbl_Parametros_puntos, Retornatable As Boolean) As NTbl_Parametros_puntos
        Dim parametros() As Object = {"@id", "@monto_puntos", "@fechainicio", "@frase", "@fechacrea", "@usuariocrea", "@isboleta", "@isfactura", "@estado", "@RazonSocial"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.Decimal, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.DateTime, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.id, d.monto_puntos, d.fechainicio, d.frase, d.fechacrea, d.usuariocrea, d.isboleta, d.isfactura, d.estado, d.RazonSocial}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Parametros_puntos_U_S", parametros, valores, tipoParametro, 10).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.monto_puntos = IIf(dt.Rows(0).Item("monto_puntos") Is DBNull.Value, Nothing, dt.Rows(0).Item("monto_puntos"))
            d.fechainicio = IIf(dt.Rows(0).Item("fechainicio") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechainicio"))
            d.frase = IIf(dt.Rows(0).Item("frase") Is DBNull.Value, Nothing, dt.Rows(0).Item("frase"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.isboleta = IIf(dt.Rows(0).Item("isboleta") Is DBNull.Value, Nothing, dt.Rows(0).Item("isboleta"))
            d.isfactura = IIf(dt.Rows(0).Item("isfactura") Is DBNull.Value, Nothing, dt.Rows(0).Item("isfactura"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.RazonSocial = IIf(dt.Rows(0).Item("RazonSocial") Is DBNull.Value, Nothing, dt.Rows(0).Item("RazonSocial"))
        Else
            d.id = Nothing
            d.monto_puntos = Nothing
            d.fechainicio = Nothing
            d.frase = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
            d.isboleta = Nothing
            d.isfactura = Nothing
            d.estado = Nothing
            d.RazonSocial = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NTbl_Parametros_puntos)
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        sql.EjecutarProcedure("Str_Tbl_Parametros_puntos_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Parametros_puntos_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTbl_Parametros_puntos) As DataTable
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Parametros_puntos_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTbl_Parametros_puntos) As NTbl_Parametros_puntos
        Dim parametros() As Object = {"@id"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.id}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Parametros_puntos_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.id = IIf(dt.Rows(0).Item("id") Is DBNull.Value, Nothing, dt.Rows(0).Item("id"))
            d.monto_puntos = IIf(dt.Rows(0).Item("monto_puntos") Is DBNull.Value, Nothing, dt.Rows(0).Item("monto_puntos"))
            d.fechainicio = IIf(dt.Rows(0).Item("fechainicio") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechainicio"))
            d.frase = IIf(dt.Rows(0).Item("frase") Is DBNull.Value, Nothing, dt.Rows(0).Item("frase"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.isboleta = IIf(dt.Rows(0).Item("isboleta") Is DBNull.Value, Nothing, dt.Rows(0).Item("isboleta"))
            d.isfactura = IIf(dt.Rows(0).Item("isfactura") Is DBNull.Value, Nothing, dt.Rows(0).Item("isfactura"))
            d.estado = IIf(dt.Rows(0).Item("estado") Is DBNull.Value, Nothing, dt.Rows(0).Item("estado"))
            d.RazonSocial = IIf(dt.Rows(0).Item("RazonSocial") Is DBNull.Value, Nothing, dt.Rows(0).Item("RazonSocial"))
        Else
            d.id = Nothing
            d.monto_puntos = Nothing
            d.fechainicio = Nothing
            d.frase = Nothing
            d.fechacrea = Nothing
            d.usuariocrea = Nothing
            d.isboleta = Nothing
            d.isfactura = Nothing
            d.estado = Nothing
            d.RazonSocial = Nothing
        End If
        Return d
    End Function
#End Region
End Class
