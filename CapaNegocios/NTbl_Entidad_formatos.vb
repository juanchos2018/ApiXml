Imports CapaDatos
Public Class NTbl_Entidad_formatos
    Dim sql As New ClsConexion
#Region "Declarations"
    Public Property identidad As String
    Public Property facturaa4 As String
    Public Property boletaa4 As String
    Public Property notacreditoa4 As String
    Public Property notadebitoa4 As String
    Public Property facturaa5 As String
    Public Property boletaa5 As String
    Public Property notacreditoa5 As String
    Public Property notadebitoa5 As String
    Public Property facturaticket As String
    Public Property boletaticket As String
    Public Property notacreditoticket As String
    Public Property notadebitoticket As String

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub


#End Region


#Region "Metodos"
    Public Sub Agregar(d As NTbl_Entidad_formatos)

        Dim parametros() As Object = {"@identidad", "@facturaa4", "@boletaa4", "@notacreditoa4", "@notadebitoa4", "@facturaa5", "@boletaa5", "@notacreditoa5", "@notadebitoa5", "@facturaticket", "@boletaticket", "@notacreditoticket", "@notadebitoticket"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.identidad, d.facturaa4, d.boletaa4, d.notacreditoa4, d.notadebitoa4, d.facturaa5, d.boletaa5, d.notacreditoa5, d.notadebitoa5, d.facturaticket, d.boletaticket, d.notacreditoticket, d.notadebitoticket}
        Sql.EjecutarProcedure("Str_Tbl_Entidad_formatos_I", parametros, valores, tipoParametro, 13)
    End Sub
    Public Sub Actualizar(d As NTbl_Entidad_formatos)
        Dim parametros() As Object = {"@identidad", "@facturaa4", "@boletaa4", "@notacreditoa4", "@notadebitoa4", "@facturaa5", "@boletaa5", "@notacreditoa5", "@notadebitoa5", "@facturaticket", "@boletaticket", "@notacreditoticket", "@notadebitoticket"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.identidad, d.facturaa4, d.boletaa4, d.notacreditoa4, d.notadebitoa4, d.facturaa5, d.boletaa5, d.notacreditoa5, d.notadebitoa5, d.facturaticket, d.boletaticket, d.notacreditoticket, d.notadebitoticket}
        Sql.EjecutarProcedure("Str_Tbl_Entidad_formatos_U", parametros, valores, tipoParametro, 13)
    End Sub
    Public Function Agregar(d As NTbl_Entidad_formatos, Retornatable As Boolean) As NTbl_Entidad_formatos

        Dim parametros() As Object = {"@identidad", "@facturaa4", "@boletaa4", "@notacreditoa4", "@notadebitoa4", "@facturaa5", "@boletaa5", "@notacreditoa5", "@notadebitoa5", "@facturaticket", "@boletaticket", "@notacreditoticket", "@notadebitoticket"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.identidad, d.facturaa4, d.boletaa4, d.notacreditoa4, d.notadebitoa4, d.facturaa5, d.boletaa5, d.notacreditoa5, d.notadebitoa5, d.facturaticket, d.boletaticket, d.notacreditoticket, d.notadebitoticket}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_Tbl_Entidad_formatos_I_S", parametros, valores, tipoParametro, 13).Tables(0)
        If dt.Rows.Count > 0 Then
            d.identidad = IIf(dt.Rows(0).Item("identidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("identidad"))
            d.facturaa4 = IIf(dt.Rows(0).Item("facturaa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaa4"))
            d.boletaa4 = IIf(dt.Rows(0).Item("boletaa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("boletaa4"))
            d.notacreditoa4 = IIf(dt.Rows(0).Item("notacreditoa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("notacreditoa4"))
            d.notadebitoa4 = IIf(dt.Rows(0).Item("notadebitoa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("notadebitoa4"))
            d.facturaa5 = IIf(dt.Rows(0).Item("facturaa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaa5"))
            d.boletaa5 = IIf(dt.Rows(0).Item("boletaa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("boletaa5"))
            d.notacreditoa5 = IIf(dt.Rows(0).Item("notacreditoa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("notacreditoa5"))
            d.notadebitoa5 = IIf(dt.Rows(0).Item("notadebitoa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("notadebitoa5"))
            d.facturaticket = IIf(dt.Rows(0).Item("facturaticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaticket"))
            d.boletaticket = IIf(dt.Rows(0).Item("boletaticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("boletaticket"))
            d.notacreditoticket = IIf(dt.Rows(0).Item("notacreditoticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("notacreditoticket"))
            d.notadebitoticket = IIf(dt.Rows(0).Item("notadebitoticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("notadebitoticket"))
        Else
            d.identidad = Nothing
            d.facturaa4 = Nothing
            d.boletaa4 = Nothing
            d.notacreditoa4 = Nothing
            d.notadebitoa4 = Nothing
            d.facturaa5 = Nothing
            d.boletaa5 = Nothing
            d.notacreditoa5 = Nothing
            d.notadebitoa5 = Nothing
            d.facturaticket = Nothing
            d.boletaticket = Nothing
            d.notacreditoticket = Nothing
            d.notadebitoticket = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NTbl_Entidad_formatos, Retornatable As Boolean) As NTbl_Entidad_formatos

        Dim parametros() As Object = {"@identidad", "@facturaa4", "@boletaa4", "@notacreditoa4", "@notadebitoa4", "@facturaa5", "@boletaa5", "@notacreditoa5", "@notadebitoa5", "@facturaticket", "@boletaticket", "@notacreditoticket", "@notadebitoticket"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.notadebitoticket = Nothing}
        Dim dt As New DataTable
        dt = Sql.ProcedureSQL("Str_Tbl_Entidad_formatos_U_S", parametros, valores, tipoParametro, 39).Tables(0)
        If dt.Rows.Count > 0 Then
            d.identidad = IIf(dt.Rows(0).Item("identidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("identidad"))
            d.facturaa4 = IIf(dt.Rows(0).Item("facturaa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaa4"))
            d.boletaa4 = IIf(dt.Rows(0).Item("boletaa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("boletaa4"))
            d.notacreditoa4 = IIf(dt.Rows(0).Item("notacreditoa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("notacreditoa4"))
            d.notadebitoa4 = IIf(dt.Rows(0).Item("notadebitoa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("notadebitoa4"))
            d.facturaa5 = IIf(dt.Rows(0).Item("facturaa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaa5"))
            d.boletaa5 = IIf(dt.Rows(0).Item("boletaa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("boletaa5"))
            d.notacreditoa5 = IIf(dt.Rows(0).Item("notacreditoa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("notacreditoa5"))
            d.notadebitoa5 = IIf(dt.Rows(0).Item("notadebitoa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("notadebitoa5"))
            d.facturaticket = IIf(dt.Rows(0).Item("facturaticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaticket"))
            d.boletaticket = IIf(dt.Rows(0).Item("boletaticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("boletaticket"))
            d.notacreditoticket = IIf(dt.Rows(0).Item("notacreditoticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("notacreditoticket"))
            d.notadebitoticket = IIf(dt.Rows(0).Item("notadebitoticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("notadebitoticket"))
        Else
            d.identidad = Nothing
            d.facturaa4 = Nothing
            d.boletaa4 = Nothing
            d.notacreditoa4 = Nothing
            d.notadebitoa4 = Nothing
            d.facturaa5 = Nothing
            d.boletaa5 = Nothing
            d.notacreditoa5 = Nothing
            d.notadebitoa5 = Nothing
            d.facturaticket = Nothing
            d.boletaticket = Nothing
            d.notacreditoticket = Nothing
            d.notadebitoticket = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NTbl_Entidad_formatos)
        Dim parametros() As Object = {"@identidad"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.identidad}
        Sql.EjecutarProcedure("Str_Tbl_Entidad_formatos_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_Tbl_Entidad_formatos(d As NTbl_Entidad_formatos) As Boolean
        Dim parametros() As Object = {"@identidad"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.identidad}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = Sql.procedimiento_escalar("Existe_Tbl_Entidad_formatos", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@identidad"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        ' dt = Sql.ProcedureSQL("Str_Tbl_Entidad_formatos_S", parametros, valores, tipoParametro, 1).Tables(0)
        dt = sql.Proc_DataReader("Str_Tbl_Entidad_formatos_S", parametros, valores, tipoParametro, 1)
        Return dt
    End Function
    Public Function Lista(d As NTbl_Entidad_formatos) As DataTable
        Dim parametros() As Object = {"@identidad"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.identidad}
        Dim dt As New DataTable
        dt = sql.Proc_DataReader("Str_Tbl_Entidad_formatos_S", parametros, valores, tipoParametro, 1)
        Return dt
    End Function

    Public Function Registro(d As NTbl_Entidad_formatos) As NTbl_Entidad_formatos
        Dim parametros() As Object = {"@identidad"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim valores() As Object = {d.identidad}
        Dim dt As New DataTable
        dt = sql.Proc_DataReader("Str_Tbl_Entidad_formatos_S", parametros, valores, tipoParametro, 1)
        If dt.Rows.Count > 0 Then
            d.identidad = IIf(dt.Rows(0).Item("identidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("identidad"))
            d.facturaa4 = IIf(dt.Rows(0).Item("facturaa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaa4"))
            d.boletaa4 = IIf(dt.Rows(0).Item("boletaa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("boletaa4"))
            d.notacreditoa4 = IIf(dt.Rows(0).Item("notacreditoa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("notacreditoa4"))
            d.notadebitoa4 = IIf(dt.Rows(0).Item("notadebitoa4") Is DBNull.Value, Nothing, dt.Rows(0).Item("notadebitoa4"))
            d.facturaa5 = IIf(dt.Rows(0).Item("facturaa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaa5"))
            d.boletaa5 = IIf(dt.Rows(0).Item("boletaa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("boletaa5"))
            d.notacreditoa5 = IIf(dt.Rows(0).Item("notacreditoa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("notacreditoa5"))
            d.notadebitoa5 = IIf(dt.Rows(0).Item("notadebitoa5") Is DBNull.Value, Nothing, dt.Rows(0).Item("notadebitoa5"))
            d.facturaticket = IIf(dt.Rows(0).Item("facturaticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("facturaticket"))
            d.boletaticket = IIf(dt.Rows(0).Item("boletaticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("boletaticket"))
            d.notacreditoticket = IIf(dt.Rows(0).Item("notacreditoticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("notacreditoticket"))
            d.notadebitoticket = IIf(dt.Rows(0).Item("notadebitoticket") Is DBNull.Value, Nothing, dt.Rows(0).Item("notadebitoticket"))
        Else
            d.identidad = Nothing
            d.facturaa4 = Nothing
            d.boletaa4 = Nothing
            d.notacreditoa4 = Nothing
            d.notadebitoa4 = Nothing
            d.facturaa5 = Nothing
            d.boletaa5 = Nothing
            d.notacreditoa5 = Nothing
            d.notadebitoa5 = Nothing
            d.facturaticket = Nothing
            d.boletaticket = Nothing
            d.notacreditoticket = Nothing
            d.notadebitoticket = Nothing
        End If
        Return d
    End Function
#End Region

End Class
