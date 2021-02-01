Imports CapaDatos
Public Class NWebComprobante
    Private sql As New ClsConexion()
    Public Property idcomprobante() As Integer
    Public Property ruc() As String
    Public Property tipo() As String
    Public Property numero() As String
    Public Property serie() As String
    Public Property fechaemision() As DateTime
    Public Property montototal() As [Decimal]
    Public Property archivo() As String
    Public Property ruccliente() As String

    Public Sub Agregar(d As NWebComprobante)

        Dim parametros As Object() = New Object() {"@ruc", "@tipo", "@numero", "@serie", "@fechaemision", "@montototal",
            "@archivo", "@ruccliente"}
        Dim tipoParametro As SqlDbType() = New SqlDbType() {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.[Date], SqlDbType.[Decimal],
            SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores As Object() = New Object() {d.ruc, d.tipo, d.numero, d.serie, d.fechaemision, d.montototal,
            d.archivo, d.ruccliente}
        sql.EjecutarProcedure("Str_comprobante_I", parametros, valores, tipoParametro, 8)
    End Sub
    Public Sub Actualizar(d As NWebComprobante)
        Dim parametros As Object() = New Object() {"@ruc", "@tipo", "@numero", "@serie", "@fechaemision", "@montototal", "@archivo", "@ruccliente"}
        Dim tipoParametro As SqlDbType() = New SqlDbType() {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.[Date], SqlDbType.[Decimal],
            SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores As Object() = New Object() {d.ruc, d.tipo, d.numero, d.serie, d.fechaemision, d.montototal, d.archivo, d.ruccliente}
        sql.EjecutarProcedure("Str_comprobante_U", parametros, valores, tipoParametro, 8)
    End Sub
    Public Sub Eliminar(d As NWebComprobante)
        Dim parametros As Object() = New Object() {"@idcomprobante"}
        Dim tipoParametro As SqlDbType() = New SqlDbType() {SqlDbType.Int}
        Dim valores As Object() = New Object() {d.idcomprobante}
        sql.EjecutarProcedure("Str_comprobante_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_comprobante(d As NWebComprobante) As Boolean
        Dim parametros As Object() = New Object() {"@ruccliente", "@idcomprobante", "@ruc", "@numero", "@serie", "@fechaemision", "@montototal"}
        Dim tipoParametro As SqlDbType() = New SqlDbType() {SqlDbType.VarChar, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Date, SqlDbType.Decimal}
        Dim valores As Object() = New Object() {d.ruccliente, d.idcomprobante, d.ruc, d.numero, d.serie, d.fechaemision, d.montototal}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_comprobante", parametros, valores, tipoParametro, 7)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros As Object() = New Object() {"@idcomprobante"}
        Dim tipoParametro As SqlDbType() = New SqlDbType() {SqlDbType.Int}
        Dim valores As Object() = New Object() {DBNull.Value}
        Dim dt As New DataTable()
        dt = sql.ProcedureSQL("Str_comprobante_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NWebComprobante) As DataTable
        Dim parametros As Object() = New Object() {"@idcomprobante"}
        Dim tipoParametro As SqlDbType() = New SqlDbType() {SqlDbType.Int}
        Dim valores As Object() = New Object() {d.idcomprobante}
        Dim dt As New DataTable()
        dt = sql.ProcedureSQL("Str_comprobante_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NWebComprobante) As NWebComprobante
        Dim parametros As Object() = New Object() {"@idcomprobante"}
        Dim tipoParametro As SqlDbType() = New SqlDbType() {SqlDbType.Int}
        Dim valores As Object() = New Object() {d.idcomprobante}
        Dim dt As New DataTable()
        dt = sql.ProcedureSQL("Str_comprobante_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.ruc = Convert.ToString(dt.Rows(0)("ruc"))
            d.tipo = Convert.ToString(dt.Rows(0)("tipo"))
            d.numero = Convert.ToString(dt.Rows(0)("numero"))
            d.serie = Convert.ToString(dt.Rows(0)("serie"))
            d.montototal = Convert.ToDecimal(dt.Rows(0)("montototal"))
            d.archivo = Convert.ToString(dt.Rows(0)("archivo"))
            d.ruccliente = Convert.ToString(dt.Rows(0)("ruccliente"))
        Else
        End If
        Return d
    End Function


End Class
