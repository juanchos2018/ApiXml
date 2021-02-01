Imports CapaDatos
Public Class Ntbl_Lic_PlazoEntrega_Obra
    Dim sql As New ClsConexion
#Region "Declarations"

    Public Property idplazoentrega As Integer
    Public Property idobras As Integer
    Public Property item As Integer
    Public Property idarticulo As String
    Public Property cantidadentrega As Integer
    Public Property plazoentrega As System.DateTime
    Public Property iddetproceso As Long
    Public Property lugarentregas_obra_plazo As String
    Public Property idproceso As Long
    Public Property DiasEntre As Integer
    Public Property entrega As String
    Public Property forma As String
    Public Property FechaRecepcion As DateTime


#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

#End Region

#Region "Metodos"
    Public Sub Agregar(d As Ntbl_Lic_PlazoEntrega_Obra)

        Dim parametros() As Object = {"@idobras", "@item", "@idarticulo", "@cantidadentrega", "@plazoentrega", "@iddetproceso", "@lugarentregas_obra_plazo", "@idproceso", "@DiasEntre", "@entrega", "@forma"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Int, SqlDbType.Char, SqlDbType.Int, SqlDbType.DateTime, SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.BigInt, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar}
        Dim valores() As Object = {d.idobras, d.item, d.idarticulo, d.cantidadentrega, d.plazoentrega, d.iddetproceso, d.lugarentregas_obra_plazo, d.idproceso, d.DiasEntre, d.entrega, d.forma}
        sql.EjecutarProcedure("Str_tbl_Lic_PlazoEntrega_Obra_I", parametros, valores, tipoParametro, 11)
    End Sub
    Public Sub Actualizar(d As Ntbl_Lic_PlazoEntrega_Obra)
        Dim parametros() As Object = {"@IdPlazoEntrega", "@idobras", "@item", "@idarticulo", "@cantidadentrega", "@plazoentrega", "@iddetproceso", "@lugarentregas_obra_plazo", "@idproceso", "@DiasEntre", "@entrega", "@forma", "@FechaRecepcion"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Int, SqlDbType.Int, SqlDbType.Char, SqlDbType.Int, SqlDbType.DateTime, SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.BigInt, SqlDbType.Int, SqlDbType.VarChar, SqlDbType.VarChar, SqlDbType.DateTime}
        Dim valores() As Object = {d.idplazoentrega, d.idobras, d.item, d.idarticulo, d.cantidadentrega, d.plazoentrega, d.iddetproceso, d.lugarentregas_obra_plazo, d.idproceso, d.DiasEntre, d.entrega, d.forma, d.FechaRecepcion}
        sql.EjecutarProcedure("Str_tbl_Lic_PlazoEntrega_Obra_U", parametros, valores, tipoParametro, 13)
    End Sub

    Public Function Agregar(d As Ntbl_Lic_PlazoEntrega_Obra, Retornatable As Boolean) As Ntbl_Lic_PlazoEntrega_Obra
        Dim parametros() As Object = {"@idobras", "@item", "@idarticulo", "@cantidadentrega", "@plazoentrega", "@iddetproceso", "@lugarentregas_obra_plazo", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Int, SqlDbType.Char, SqlDbType.Int, SqlDbType.DateTime, SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.BigInt}
        Dim valores() As Object = {d.idobras, d.item, d.idarticulo, d.cantidadentrega, d.plazoentrega, d.iddetproceso, d.lugarentregas_obra_plazo, d.idproceso}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_PlazoEntrega_Obra_I_S", parametros, valores, tipoParametro, 8).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idplazoentrega = IIf(dt.Rows(0).Item("idplazoentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("idplazoentrega"))
            d.idobras = IIf(dt.Rows(0).Item("idobras") Is DBNull.Value, Nothing, dt.Rows(0).Item("idobras"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.cantidadentrega = IIf(dt.Rows(0).Item("cantidadentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidadentrega"))
            d.plazoentrega = IIf(dt.Rows(0).Item("plazoentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("plazoentrega"))
            d.iddetproceso = IIf(dt.Rows(0).Item("iddetproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetproceso"))
            d.lugarentregas_obra_plazo = IIf(dt.Rows(0).Item("lugarentregas_obra_plazo") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarentregas_obra_plazo"))
            d.idproceso = IIf(dt.Rows(0).Item("idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproceso"))
        Else
            d.idplazoentrega = Nothing
            d.idobras = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.cantidadentrega = Nothing
            d.plazoentrega = Nothing
            d.iddetproceso = Nothing
            d.lugarentregas_obra_plazo = Nothing
            d.idproceso = Nothing
        End If
        Return d
    End Function
    Public Function Actualizar(d As Ntbl_Lic_PlazoEntrega_Obra, Retornatable As Boolean) As Ntbl_Lic_PlazoEntrega_Obra

        Dim parametros() As Object = {"@idobras", "@item", "@idarticulo", "@cantidadentrega", "@plazoentrega", "@iddetproceso", "@lugarentregas_obra_plazo", "@idproceso"}
        Dim tipoParametro() As Object = {SqlDbType.Int, SqlDbType.Int, SqlDbType.Char, SqlDbType.Int, SqlDbType.DateTime, SqlDbType.BigInt, SqlDbType.VarChar, SqlDbType.BigInt}
        Dim valores() As Object = {d.idproceso = Nothing}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_PlazoEntrega_Obra_U_S", parametros, valores, tipoParametro, 26).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idplazoentrega = IIf(dt.Rows(0).Item("idplazoentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("idplazoentrega"))
            d.idobras = IIf(dt.Rows(0).Item("idobras") Is DBNull.Value, Nothing, dt.Rows(0).Item("idobras"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.cantidadentrega = IIf(dt.Rows(0).Item("cantidadentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidadentrega"))
            d.plazoentrega = IIf(dt.Rows(0).Item("plazoentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("plazoentrega"))
            d.iddetproceso = IIf(dt.Rows(0).Item("iddetproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetproceso"))
            d.lugarentregas_obra_plazo = IIf(dt.Rows(0).Item("lugarentregas_obra_plazo") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarentregas_obra_plazo"))
            d.idproceso = IIf(dt.Rows(0).Item("idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproceso"))
        Else
            d.idplazoentrega = Nothing
            d.idobras = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.cantidadentrega = Nothing
            d.plazoentrega = Nothing
            d.iddetproceso = Nothing
            d.lugarentregas_obra_plazo = Nothing
            d.idproceso = Nothing
        End If
        Return d
    End Function
    Public Sub Eliminar(d As Ntbl_Lic_PlazoEntrega_Obra)
        Dim parametros() As Object = {"@idplazoentrega", "@IdObras"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.BigInt}
        Dim valores() As Object = {d.idplazoentrega, d.idobras}
        sql.EjecutarProcedure("Str_tbl_Lic_PlazoEntrega_Obra_D", parametros, valores, tipoParametro, 2)
    End Sub
    Public Function Existe_tbl_Lic_PlazoEntrega_Obra(d As Ntbl_Lic_PlazoEntrega_Obra) As Boolean
        Dim parametros() As Object = {"@idplazoentrega"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {d.idplazoentrega}
        Dim resultado As Integer
        Dim bandera As Boolean = False
        resultado = sql.procedimiento_escalar("Existe_tbl_Lic_PlazoEntrega_Obra", parametros, valores, tipoParametro, 1)
        If resultado = 1 Then
            bandera = True
        Else
            bandera = False
        End If
        Return bandera
    End Function
    Public Function Lista() As DataTable
        Dim parametros() As Object = {"@idplazoentrega"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt}
        Dim valores() As Object = {DBNull.Value}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_PlazoEntrega_Obra_S", parametros, valores, tipoParametro, 1).Tables(0)
        Return dt
    End Function
    Public Function Lista(d As Ntbl_Lic_PlazoEntrega_Obra) As DataTable
        Dim parametros() As Object = {"@idplazoentrega", "@IdObras"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.BigInt}
        Dim valores() As Object = {d.idplazoentrega, d.idobras}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_PlazoEntrega_Obra_S", parametros, valores, tipoParametro, 2).Tables(0)
        Return dt
    End Function
    Public Function Registro(d As Ntbl_Lic_PlazoEntrega_Obra) As Ntbl_Lic_PlazoEntrega_Obra
        Dim parametros() As Object = {"@idplazoentrega", "@IdObras"}
        Dim tipoParametro() As Object = {SqlDbType.BigInt, SqlDbType.BigInt}
        Dim valores() As Object = {d.idplazoentrega, d.idobras}
        Dim dt As New DataTable
        dt = sql.ProcedureSQL("Str_tbl_Lic_PlazoEntrega_Obra_S", parametros, valores, tipoParametro, 2).Tables(0)
        If dt.Rows.Count > 0 Then
            d.idplazoentrega = IIf(dt.Rows(0).Item("idplazoentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("idplazoentrega"))
            d.idobras = IIf(dt.Rows(0).Item("idobras") Is DBNull.Value, Nothing, dt.Rows(0).Item("idobras"))
            d.item = IIf(dt.Rows(0).Item("item") Is DBNull.Value, Nothing, dt.Rows(0).Item("item"))
            d.idarticulo = IIf(dt.Rows(0).Item("idarticulo") Is DBNull.Value, Nothing, dt.Rows(0).Item("idarticulo"))
            d.cantidadentrega = IIf(dt.Rows(0).Item("cantidadentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("cantidadentrega"))
            d.plazoentrega = IIf(dt.Rows(0).Item("plazoentrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("plazoentrega"))
            d.iddetproceso = IIf(dt.Rows(0).Item("iddetproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("iddetproceso"))
            d.lugarentregas_obra_plazo = IIf(dt.Rows(0).Item("lugarentregas_obra_plazo") Is DBNull.Value, Nothing, dt.Rows(0).Item("lugarentregas_obra_plazo"))
            d.idproceso = IIf(dt.Rows(0).Item("idproceso") Is DBNull.Value, Nothing, dt.Rows(0).Item("idproceso"))
            d.DiasEntre = IIf(dt.Rows(0).Item("DiasEntre") Is DBNull.Value, Nothing, dt.Rows(0).Item("DiasEntre"))
            d.entrega = IIf(dt.Rows(0).Item("entrega") Is DBNull.Value, Nothing, dt.Rows(0).Item("entrega"))
            d.forma = IIf(dt.Rows(0).Item("forma") Is DBNull.Value, Nothing, dt.Rows(0).Item("forma"))
            d.FechaRecepcion = IIf(dt.Rows(0).Item("FechaRecepcion") Is DBNull.Value, Nothing, dt.Rows(0).Item("FechaRecepcion"))

        Else
            d.idplazoentrega = Nothing
            d.idobras = Nothing
            d.item = Nothing
            d.idarticulo = Nothing
            d.cantidadentrega = Nothing
            d.plazoentrega = Nothing
            d.iddetproceso = Nothing
            d.lugarentregas_obra_plazo = Nothing
            d.idproceso = Nothing
            d.DiasEntre = Nothing
            d.entrega = Nothing
            d.forma = Nothing
        End If
        Return d
    End Function
#End Region

End Class
