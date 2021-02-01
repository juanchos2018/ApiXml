Imports CapaDatos
Public Class NTbl_Surtidor_Lado
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property idsurtidor As Integer
    Public Property lado As String
    Public Property transact As String
    Public Property descripcion As String
    Public Property rutafile As String
    Public Property usuariocrea As String
    Public Property fechacrea As System.DateTime

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NTbl_Surtidor_Lado)

        Dim parametros() As Object = {"@idsurtidor", "@lado", "@transact", "@descripcion", "@rutafile", "@usuariocrea", "@fechacrea"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.idsurtidor, d.lado, d.transact, d.descripcion, d.rutafile, d.usuariocrea, d.fechacrea}
        sql.EjecutarProcedure("Str_Tbl_Surtidor_Lado_I", parametros, valores, tipoParametro, 7)
    End Sub
    Public Sub Actualizar(d As NTbl_Surtidor_Lado)
        Dim parametros() As Object = {"@idsurtidor", "@lado", "@transact", "@descripcion", "@rutafile", "@usuariocrea", "@fechacrea"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.idsurtidor, d.lado, d.transact, d.descripcion, d.rutafile, d.usuariocrea, d.fechacrea}
        sql.EjecutarProcedure("Str_Tbl_Surtidor_Lado_U", parametros, valores, tipoParametro, 7)
    End Sub
    Public Function Agregar(d As NTbl_Surtidor_Lado, Retornatable As Boolean) As NTbl_Surtidor_Lado

        Dim parametros() As Object = {"@idsurtidor", "@lado", "@transact", "@descripcion", "@rutafile", "@usuariocrea", "@fechacrea"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.idsurtidor, d.lado, d.transact, d.descripcion, d.rutafile, d.usuariocrea, d.fechacrea}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Surtidor_Lado_I_S", parametros, valores, tipoParametro, 7).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idsurtidor = IIf(dt.Rows(0).Item("idsurtidor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsurtidor"))
            d.lado = IIf(dt.Rows(0).Item("lado") Is DBNull.Value, Nothing, dt.Rows(0).Item("lado"))
            d.transact = IIf(dt.Rows(0).Item("transact") Is DBNull.Value, Nothing, dt.Rows(0).Item("transact"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.rutafile = IIf(dt.Rows(0).Item("rutafile") Is DBNull.Value, Nothing, dt.Rows(0).Item("rutafile"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
        Else
            d.idsurtidor = Nothing
            d.lado = Nothing
            d.transact = Nothing
            d.descripcion = Nothing
            d.rutafile = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NTbl_Surtidor_Lado, Retornatable As Boolean) As NTbl_Surtidor_Lado

        Dim parametros() As Object = {"@idsurtidor", "@lado", "@transact", "@descripcion", "@rutafile", "@usuariocrea", "@fechacrea"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.fechacrea = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Surtidor_Lado_U_S", parametros, valores, tipoParametro, 7).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idsurtidor = IIf(dt.Rows(0).Item("idsurtidor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsurtidor"))
            d.lado = IIf(dt.Rows(0).Item("lado") Is DBNull.Value, Nothing, dt.Rows(0).Item("lado"))
            d.transact = IIf(dt.Rows(0).Item("transact") Is DBNull.Value, Nothing, dt.Rows(0).Item("transact"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.rutafile = IIf(dt.Rows(0).Item("rutafile") Is DBNull.Value, Nothing, dt.Rows(0).Item("rutafile"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
        Else
            d.idsurtidor = Nothing
            d.lado = Nothing
            d.transact = Nothing
            d.descripcion = Nothing
            d.rutafile = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NTbl_Surtidor_Lado)
        Dim parametros() As Object = {"@idsurtidor", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar}
        Dim valores() As Object = {d.idsurtidor, d.lado}
        sql.EjecutarProcedure("Str_Tbl_Surtidor_Lado_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idsurtidor", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Surtidor_Lado_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTbl_Surtidor_Lado) As DataTable
        Dim parametros() As Object = {"@idsurtidor", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar}
        Dim valores() As Object = {d.idsurtidor, d.lado}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Surtidor_Lado_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTbl_Surtidor_Lado) As NTbl_Surtidor_Lado
        Dim parametros() As Object = {"@idsurtidor", "@lado"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.VarChar}
        Dim valores() As Object = {d.idsurtidor, d.lado}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_Surtidor_Lado_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idsurtidor = IIf(dt.Rows(0).Item("idsurtidor") Is DBNull.Value, Nothing, dt.Rows(0).Item("idsurtidor"))
            d.lado = IIf(dt.Rows(0).Item("lado") Is DBNull.Value, Nothing, dt.Rows(0).Item("lado"))
            d.transact = IIf(dt.Rows(0).Item("transact") Is DBNull.Value, Nothing, dt.Rows(0).Item("transact"))
            d.descripcion = IIf(dt.Rows(0).Item("descripcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("descripcion"))
            d.rutafile = IIf(dt.Rows(0).Item("rutafile") Is DBNull.Value, Nothing, dt.Rows(0).Item("rutafile"))
            d.usuariocrea = IIf(dt.Rows(0).Item("usuariocrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("usuariocrea"))
            d.fechacrea = IIf(dt.Rows(0).Item("fechacrea") Is DBNull.Value, Nothing, dt.Rows(0).Item("fechacrea"))
        Else
            d.idsurtidor = Nothing
            d.lado = Nothing
            d.transact = Nothing
            d.descripcion = Nothing
            d.rutafile = Nothing
            d.usuariocrea = Nothing
            d.fechacrea = Nothing
        End If
        Return d
    End Function
    ''' <summary>
    ''' Lista los surtidores relacionados por lado y tablas dbf
    ''' </summary>
    ''' <returns></returns>
    Public Function lista_surt_lado() As DataTable
        Dim c As String = " SELECT     s.IdSurtidor, s.item, s.IdCaja, s.Nombre,ts.NroMangueras, sl.lado, sl.Transact, sl.RutaFile,  "
        c += " sl.Descripcion FROM  tbl_Surtidor AS s INNER JOIN tbl_Tipo_Surtidor AS ts ON s.item = ts.Item LEFT OUTER JOIN "
        c += " Tbl_Surtidor_lado AS sl ON s.IdSurtidor = sl.IdSurtidor "
        Return sql.EjecutarConsulta("d", c).Tables(0)

    End Function
#End Region

End Class
