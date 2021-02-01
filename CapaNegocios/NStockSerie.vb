Imports CapaDatos
Public Class NStockSerie
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idalmacen As String
    Public Property idarticulo As String
    Public Property nroserie As String
    Public Property cantidad As Decimal
    Public Property tipodocref As String
    Public Property nrodocref As String
    Public Property fecharef As System.DateTime
    Public Property idagencia As String
    Public Property saldo As Decimal
    Public Property ingreso As Decimal
    Public Property salida As Decimal

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region
#Region "Metodos"
    Public Sub Agregar(d As NStockSerie)

        Dim parametros() As Object = {"@idalmacen", "@idarticulo", "@nroserie", "@cantidad", "@tipodocref", "@nrodocref", "@fecharef", "@idagencia", "@saldo", "@ingreso", "@salida"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal}
        Dim valores() As Object = {d.idalmacen, d.idarticulo, d.nroserie, d.cantidad, d.tipodocref, d.nrodocref, d.fecharef, d.idagencia, d.saldo, d.ingreso, d.salida}
        sql.EjecutarProcedure("Str_StockSerie_I", parametros, valores, tipoParametro, 11)
    End Sub
    Public Sub Actualizar(d As NStockSerie)
        Dim parametros() As Object = {"@idalmacen", "@idarticulo", "@nroserie", "@cantidad", "@tipodocref", "@nrodocref", "@fecharef", "@idagencia", "@saldo", "@ingreso", "@salida"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal}
        Dim valores() As Object = {d.idalmacen, d.idarticulo, d.nroserie, d.cantidad, d.tipodocref, d.nrodocref, d.fecharef, d.idagencia, d.saldo, d.ingreso, d.salida}
        sql.EjecutarProcedure("Str_StockSerie_U", parametros, valores, tipoParametro, 11)
    End Sub
    Public Function Agregar(d As NStockSerie, Retornatable As Boolean) As NStockSerie

        Dim parametros() As Object = {"@idalmacen", "@idarticulo", "@nroserie", "@cantidad", "@tipodocref", "@nrodocref", "@fecharef", "@idagencia", "@saldo", "@ingreso", "@salida"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal}
        Dim valores() As Object = {d.idalmacen, d.idarticulo, d.nroserie, d.cantidad, d.tipodocref, d.nrodocref, d.fecharef, d.idagencia, d.saldo, d.ingreso, d.salida}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_StockSerie_I_S", parametros, valores, tipoParametro, 11).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.nroserie = IIf(dt.Rows(0).Item("nroserie") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroserie"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.tipodocref = IIf(dt.Rows(0).Item("tipodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocref"))
            d.nrodocref = IIf(dt.Rows(0).Item("nrodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocref"))
            d.fecharef = IIf(dt.Rows(0).Item("fecharef") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecharef"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.ingreso = IIf(dt.Rows(0).Item("ingreso") Is DBNull.Value, Nothing, dt.Rows(0).Item("ingreso"))
            d.salida = IIf(dt.Rows(0).Item("salida") Is DBNull.Value, Nothing, dt.Rows(0).Item("salida"))
        Else
            d.idalmacen = Nothing
            d.idarticulo = Nothing
            d.nroserie = Nothing
            d.cantidad = Nothing
            d.tipodocref = Nothing
            d.nrodocref = Nothing
            d.fecharef = Nothing
            d.idagencia = Nothing
            d.saldo = Nothing
            d.ingreso = Nothing
            d.salida = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As NStockSerie, Retornatable As Boolean) As NStockSerie

        Dim parametros() As Object = {"@idalmacen", "@idarticulo", "@nroserie", "@cantidad", "@tipodocref", "@nrodocref", "@fecharef", "@idagencia", "@saldo", "@ingreso", "@salida"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar, SqlDbType.Decimal, SqlDbType.Char, SqlDbType.Char, SqlDbType.DateTime, SqlDbType.Char, SqlDbType.Decimal, SqlDbType.Decimal, SqlDbType.Decimal}
        Dim valores() As Object = {d.salida = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_StockSerie_U_S", parametros, valores, tipoParametro, 33).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.nroserie = IIf(dt.Rows(0).Item("nroserie") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroserie"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.tipodocref = IIf(dt.Rows(0).Item("tipodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocref"))
            d.nrodocref = IIf(dt.Rows(0).Item("nrodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocref"))
            d.fecharef = IIf(dt.Rows(0).Item("fecharef") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecharef"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.ingreso = IIf(dt.Rows(0).Item("ingreso") Is DBNull.Value, Nothing, dt.Rows(0).Item("ingreso"))
            d.salida = IIf(dt.Rows(0).Item("salida") Is DBNull.Value, Nothing, dt.Rows(0).Item("salida"))
        Else
            d.idalmacen = Nothing
            d.idarticulo = Nothing
            d.nroserie = Nothing
            d.cantidad = Nothing
            d.tipodocref = Nothing
            d.nrodocref = Nothing
            d.fecharef = Nothing
            d.idagencia = Nothing
            d.saldo = Nothing
            d.ingreso = Nothing
            d.salida = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As NStockSerie)
        Dim parametros() As Object = {"@idalmacen", "@idarticulo", "@nroserie"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idalmacen, d.idarticulo, d.nroserie}
        sql.EjecutarProcedure("Str_StockSerie_D", parametros, valores, tipoParametro, 3)
    End Sub
    Public Function Existe_StockSerie(d As NStockSerie) As Boolean
        Dim parametros() As Object = {"@idalmacen", "@idarticulo", "@nroserie"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idalmacen, d.idarticulo, d.nroserie}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_StockSerie", parametros, valores, tipoParametro, 3)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idalmacen", "@idarticulo", "@nroserie"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_StockSerie_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As NStockSerie) As DataTable
        Dim parametros() As Object = {"@idalmacen", "@idarticulo", "@nroserie"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idalmacen, d.idarticulo, d.nroserie}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_StockSerie_S", parametros, valores, tipoParametro, 3).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As NStockSerie) As NStockSerie
        Dim parametros() As Object = {"@idalmacen", "@idarticulo", "@nroserie"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.VarChar}
        Dim valores() As Object = {d.idalmacen, d.idarticulo, d.nroserie}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_StockSerie_S", parametros, valores, tipoParametro, 3).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idalmacen = IIf(dt.Rows(0).Item("idalmacen") Is DBNull.Value, Nothing, dt.Rows(0).Item("idalmacen"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.nroserie = IIf(dt.Rows(0).Item("nroserie") Is DBNull.Value, Nothing, dt.Rows(0).Item("nroserie"))
            d.cantidad = IIf(dt.Rows(0).Item("cantidad") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidad"))
            d.tipodocref = IIf(dt.Rows(0).Item("tipodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("tipodocref"))
            d.nrodocref = IIf(dt.Rows(0).Item("nrodocref") Is DBNull.Value, Nothing, dt.Rows(0).Item("nrodocref"))
            d.fecharef = IIf(dt.Rows(0).Item("fecharef") Is DBNull.Value, Nothing, dt.Rows(0).Item("fecharef"))
            d.idagencia = IIf(dt.Rows(0).Item("idagencia") Is DBNull.Value, Nothing, dt.Rows(0).Item("idagencia"))
            d.saldo = IIf(dt.Rows(0).Item("saldo") Is DBNull.Value, Nothing, dt.Rows(0).Item("saldo"))
            d.ingreso = IIf(dt.Rows(0).Item("ingreso") Is DBNull.Value, Nothing, dt.Rows(0).Item("ingreso"))
            d.salida = IIf(dt.Rows(0).Item("salida") Is DBNull.Value, Nothing, dt.Rows(0).Item("salida"))
        Else
            d.idalmacen = Nothing
            d.idarticulo = Nothing
            d.nroserie = Nothing
            d.cantidad = Nothing
            d.tipodocref = Nothing
            d.nrodocref = Nothing
            d.fecharef = Nothing
            d.idagencia = Nothing
            d.saldo = Nothing
            d.ingreso = Nothing
            d.salida = Nothing
        End If
        Return d
    End Function

    Public Function Lista_StockSerie(d As NStockSerie, esserie As Boolean) As DataTable
        Dim parametros() As Object = {"@idalmacen", "@idarticulo", "@IsSerie"}
        Dim tipoParametro() As Object = {SqlDbType.Char, SqlDbType.Char, SqlDbType.Bit}
        Dim valores() As Object = {d.idalmacen, d.idarticulo, esserie}
        Return sql.ProcedureSQL("Str_Stock_Serie", parametros, valores, tipoParametro, 3).Tables(0)
    End Function
#End Region


End Class
