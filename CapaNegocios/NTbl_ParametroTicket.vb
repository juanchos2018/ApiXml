Imports CapaDatos
Public Class NTbl_ParametroTicket
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property item As String
    Public Property nombrecomercial As String
    Public Property muestranomcom As Boolean
    Public Property muestradomfis As Boolean
    Public Property muestrasucur As Boolean
    Public Property muestrafondonegro As Boolean
    Public Property muestrafechaimpre As Boolean
    Public Property muestrahoraimpre As Boolean
    Public Property imprimeoriginal As Boolean
    Public Property nombreoriginal As String
    Public Property imprimecopia1 As Boolean
    Public Property nombrecopia1 As String
    Public Property imprimecopia2 As Boolean
    Public Property nombrecopia2 As String
    Public Property imprimeglosa As Boolean
    Public Property glosa As String
    Public Property AplicarReimpresion As Boolean

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As NTbl_ParametroTicket)

        Dim parametros() As Object = {"@item", "@nombrecomercial", "@muestranomcom", "@muestradomfis", "@muestrasucur", "@muestrafondonegro", "@muestrafechaimpre", "@muestrahoraimpre", "@imprimeoriginal", "@nombreoriginal", "@imprimecopia1", "@nombrecopia1", "@imprimecopia2", "@nombrecopia2", "@imprimeglosa", "@glosa", "@aplicarreimpresion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.item, d.nombrecomercial, d.muestranomcom, d.muestradomfis, d.muestrasucur, d.muestrafondonegro, d.muestrafechaimpre, d.muestrahoraimpre, d.imprimeoriginal, d.nombreoriginal, d.imprimecopia1, d.nombrecopia1, d.imprimecopia2, d.nombrecopia2, d.imprimeglosa, d.glosa, d.aplicarreimpresion}
        sql.EjecutarProcedure("Str_Tbl_ParametroTicket_I", parametros, valores, tipoParametro, 17)
    End Sub
    Public Sub Actualizar(d As NTbl_ParametroTicket)
        Dim parametros() As Object = {"@item", "@nombrecomercial", "@muestranomcom", "@muestradomfis", "@muestrasucur", "@muestrafondonegro", "@muestrafechaimpre", "@muestrahoraimpre", "@imprimeoriginal", "@nombreoriginal", "@imprimecopia1", "@nombrecopia1", "@imprimecopia2", "@nombrecopia2", "@imprimeglosa", "@glosa", "@aplicarreimpresion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.item, d.nombrecomercial, d.muestranomcom, d.muestradomfis, d.muestrasucur, d.muestrafondonegro, d.muestrafechaimpre, d.muestrahoraimpre, d.imprimeoriginal, d.nombreoriginal, d.imprimecopia1, d.nombrecopia1, d.imprimecopia2, d.nombrecopia2, d.imprimeglosa, d.glosa, d.aplicarreimpresion}
        sql.EjecutarProcedure("Str_Tbl_ParametroTicket_U", parametros, valores, tipoParametro, 17)
    End Sub
    Public Function Agregar(d As NTbl_ParametroTicket, Retornatable As Boolean) As NTbl_ParametroTicket

        Dim parametros() As Object = {"@item", "@nombrecomercial", "@muestranomcom", "@muestradomfis", "@muestrasucur", "@muestrafondonegro", "@muestrafechaimpre", "@muestrahoraimpre", "@imprimeoriginal", "@nombreoriginal", "@imprimecopia1", "@nombrecopia1", "@imprimecopia2", "@nombrecopia2", "@imprimeglosa", "@glosa", "@aplicarreimpresion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.item, d.nombrecomercial, d.muestranomcom, d.muestradomfis, d.muestrasucur, d.muestrafondonegro, d.muestrafechaimpre, d.muestrahoraimpre, d.imprimeoriginal, d.nombreoriginal, d.imprimecopia1, d.nombrecopia1, d.imprimecopia2, d.nombrecopia2, d.imprimeglosa, d.glosa, d.aplicarreimpresion}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_ParametroTicket_I_S", parametros, valores, tipoParametro, 17).Tables(0)
        If dt.Rows.Count > 0 Then
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.nombrecomercial = IIf(dt.Rows(0).Item("nombrecomercial") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecomercial"))
            d.muestranomcom = IIf(dt.Rows(0).Item("muestranomcom") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestranomcom"))
            d.muestradomfis = IIf(dt.Rows(0).Item("muestradomfis") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestradomfis"))
            d.muestrasucur = IIf(dt.Rows(0).Item("muestrasucur") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrasucur"))
            d.muestrafondonegro = IIf(dt.Rows(0).Item("muestrafondonegro") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrafondonegro"))
            d.muestrafechaimpre = IIf(dt.Rows(0).Item("muestrafechaimpre") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrafechaimpre"))
            d.muestrahoraimpre = IIf(dt.Rows(0).Item("muestrahoraimpre") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrahoraimpre"))
            d.imprimeoriginal = IIf(dt.Rows(0).Item("imprimeoriginal") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimeoriginal"))
            d.nombreoriginal = IIf(dt.Rows(0).Item("nombreoriginal") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreoriginal"))
            d.imprimecopia1 = IIf(dt.Rows(0).Item("imprimecopia1") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimecopia1"))
            d.nombrecopia1 = IIf(dt.Rows(0).Item("nombrecopia1") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecopia1"))
            d.imprimecopia2 = IIf(dt.Rows(0).Item("imprimecopia2") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimecopia2"))
            d.nombrecopia2 = IIf(dt.Rows(0).Item("nombrecopia2") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecopia2"))
            d.imprimeglosa = IIf(dt.Rows(0).Item("imprimeglosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimeglosa"))
            d.glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.aplicarreimpresion = IIf(dt.Rows(0).Item("aplicarreimpresion") Is DBNull.Value, Nothing, dt.Rows(0).Item("aplicarreimpresion"))
        Else
            d.item = Nothing
            d.nombrecomercial = Nothing
            d.muestranomcom = Nothing
            d.muestradomfis = Nothing
            d.muestrasucur = Nothing
            d.muestrafondonegro = Nothing
            d.muestrafechaimpre = Nothing
            d.muestrahoraimpre = Nothing
            d.imprimeoriginal = Nothing
            d.nombreoriginal = Nothing
            d.imprimecopia1 = Nothing
            d.nombrecopia1 = Nothing
            d.imprimecopia2 = Nothing
            d.nombrecopia2 = Nothing
            d.imprimeglosa = Nothing
            d.glosa = Nothing
            d.aplicarreimpresion = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NTbl_ParametroTicket, Retornatable As Boolean) As NTbl_ParametroTicket

        Dim parametros() As Object = {"@item", "@nombrecomercial", "@muestranomcom", "@muestradomfis", "@muestrasucur", "@muestrafondonegro", "@muestrafechaimpre", "@muestrahoraimpre", "@imprimeoriginal", "@nombreoriginal", "@imprimecopia1", "@nombrecopia1", "@imprimecopia2", "@nombrecopia2", "@imprimeglosa", "@glosa", "@aplicarreimpresion"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit, SqlDbType.VarChar, SqlDbType.Bit}
        Dim valores() As Object = {d.aplicarreimpresion = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_ParametroTicket_U_S", parametros, valores, tipoParametro, 51).Tables(0)
        If dt.Rows.Count > 0 Then
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.nombrecomercial = IIf(dt.Rows(0).Item("nombrecomercial") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecomercial"))
            d.muestranomcom = IIf(dt.Rows(0).Item("muestranomcom") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestranomcom"))
            d.muestradomfis = IIf(dt.Rows(0).Item("muestradomfis") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestradomfis"))
            d.muestrasucur = IIf(dt.Rows(0).Item("muestrasucur") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrasucur"))
            d.muestrafondonegro = IIf(dt.Rows(0).Item("muestrafondonegro") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrafondonegro"))
            d.muestrafechaimpre = IIf(dt.Rows(0).Item("muestrafechaimpre") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrafechaimpre"))
            d.muestrahoraimpre = IIf(dt.Rows(0).Item("muestrahoraimpre") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrahoraimpre"))
            d.imprimeoriginal = IIf(dt.Rows(0).Item("imprimeoriginal") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimeoriginal"))
            d.nombreoriginal = IIf(dt.Rows(0).Item("nombreoriginal") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreoriginal"))
            d.imprimecopia1 = IIf(dt.Rows(0).Item("imprimecopia1") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimecopia1"))
            d.nombrecopia1 = IIf(dt.Rows(0).Item("nombrecopia1") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecopia1"))
            d.imprimecopia2 = IIf(dt.Rows(0).Item("imprimecopia2") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimecopia2"))
            d.nombrecopia2 = IIf(dt.Rows(0).Item("nombrecopia2") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecopia2"))
            d.imprimeglosa = IIf(dt.Rows(0).Item("imprimeglosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimeglosa"))
            d.glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.aplicarreimpresion = IIf(dt.Rows(0).Item("aplicarreimpresion") Is DBNull.Value, Nothing, dt.Rows(0).Item("aplicarreimpresion"))
        Else
            d.item = Nothing
            d.nombrecomercial = Nothing
            d.muestranomcom = Nothing
            d.muestradomfis = Nothing
            d.muestrasucur = Nothing
            d.muestrafondonegro = Nothing
            d.muestrafechaimpre = Nothing
            d.muestrahoraimpre = Nothing
            d.imprimeoriginal = Nothing
            d.nombreoriginal = Nothing
            d.imprimecopia1 = Nothing
            d.nombrecopia1 = Nothing
            d.imprimecopia2 = Nothing
            d.nombrecopia2 = Nothing
            d.imprimeglosa = Nothing
            d.glosa = Nothing
            d.aplicarreimpresion = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NTbl_ParametroTicket)
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.item}
        sql.EjecutarProcedure("Str_Tbl_ParametroTicket_D", parametros, valores, tipoParametro, 1)
    End Sub
    Public Function Existe_Tbl_ParametroTicket(d As NTbl_ParametroTicket) As Boolean
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.item}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_Tbl_ParametroTicket", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_ParametroTicket_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NTbl_ParametroTicket) As DataTable
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_ParametroTicket_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NTbl_ParametroTicket) As NTbl_ParametroTicket
        Dim parametros() As Object = {"@item"}
        Dim tipoParametro() As Object = {SqlDbType.VarChar}
        Dim valores() As Object = {d.item}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_Tbl_ParametroTicket_S", parametros, valores, tipoParametro, 1).Tables(0)
        If dt.Rows.Count > 0 Then
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.nombrecomercial = IIf(dt.Rows(0).Item("nombrecomercial") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecomercial"))
            d.muestranomcom = IIf(dt.Rows(0).Item("muestranomcom") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestranomcom"))
            d.muestradomfis = IIf(dt.Rows(0).Item("muestradomfis") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestradomfis"))
            d.muestrasucur = IIf(dt.Rows(0).Item("muestrasucur") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrasucur"))
            d.muestrafondonegro = IIf(dt.Rows(0).Item("muestrafondonegro") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrafondonegro"))
            d.muestrafechaimpre = IIf(dt.Rows(0).Item("muestrafechaimpre") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrafechaimpre"))
            d.muestrahoraimpre = IIf(dt.Rows(0).Item("muestrahoraimpre") Is DBNull.Value, Nothing, dt.Rows(0).Item("muestrahoraimpre"))
            d.imprimeoriginal = IIf(dt.Rows(0).Item("imprimeoriginal") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimeoriginal"))
            d.nombreoriginal = IIf(dt.Rows(0).Item("nombreoriginal") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombreoriginal"))
            d.imprimecopia1 = IIf(dt.Rows(0).Item("imprimecopia1") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimecopia1"))
            d.nombrecopia1 = IIf(dt.Rows(0).Item("nombrecopia1") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecopia1"))
            d.imprimecopia2 = IIf(dt.Rows(0).Item("imprimecopia2") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimecopia2"))
            d.nombrecopia2 = IIf(dt.Rows(0).Item("nombrecopia2") Is DBNull.Value, Nothing, dt.Rows(0).Item("nombrecopia2"))
            d.imprimeglosa = IIf(dt.Rows(0).Item("imprimeglosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("imprimeglosa"))
            d.glosa = IIf(dt.Rows(0).Item("glosa") Is DBNull.Value, Nothing, dt.Rows(0).Item("glosa"))
            d.aplicarreimpresion = IIf(dt.Rows(0).Item("aplicarreimpresion") Is DBNull.Value, Nothing, dt.Rows(0).Item("aplicarreimpresion"))
        Else
            d.item = Nothing
            d.nombrecomercial = Nothing
            d.muestranomcom = Nothing
            d.muestradomfis = Nothing
            d.muestrasucur = Nothing
            d.muestrafondonegro = Nothing
            d.muestrafechaimpre = Nothing
            d.muestrahoraimpre = Nothing
            d.imprimeoriginal = Nothing
            d.nombreoriginal = Nothing
            d.imprimecopia1 = Nothing
            d.nombrecopia1 = Nothing
            d.imprimecopia2 = Nothing
            d.nombrecopia2 = Nothing
            d.imprimeglosa = Nothing
            d.glosa = Nothing
            d.aplicarreimpresion = Nothing
        End If
        Return d
    End Function
#End Region

End Class
