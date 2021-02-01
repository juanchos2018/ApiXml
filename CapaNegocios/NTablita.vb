Imports CapaDatos
Public Class NTablita
    Dim sql As New ClsConexion

    Private _fecha As DateTime
    Private _Fecha1 As String
    Private _fecha2 As Date
    Public Property Fecha As DateTime
        Get
            Return _fecha
        End Get
        Set(value As DateTime)
            _fecha = value
        End Set
    End Property

    Public Property Fecha1 As String
        Get
            Return _Fecha1
        End Get
        Set(value As String)
            _Fecha1 = value
        End Set
    End Property

    Public Property Fecha2 As Date
        Get
            Return _fecha2
        End Get
        Set(value As Date)
            _fecha2 = value
        End Set
    End Property


    Public Sub agregar(n As NTablita)
        Dim params() As Object = {"@Fecha"}
        Dim tipoParametro() As Object = {SqlDbType.DateTime}
        Dim vals() As Object = {n.Fecha}
        sql.EjecutarProcedure("prc_tbl", params, vals, tipoParametro, 1)
    End Sub
    Public Sub agregar1(n As NTablita)
        Dim params() As Object = {"@Fecha"}
        Dim tipoParametro() As Object = {SqlDbType.Char}
        Dim vals() As Object = {n.Fecha1}
        sql.EjecutarProcedure("prc_tbl1", params, vals, tipoParametro, 1)
    End Sub
    Public Sub agregar2(n As NTablita)
        Dim params() As Object = {"@Fecha"}
        Dim tipoParametro() As Object = {SqlDbType.Date}
        Dim vals() As Object = {n.Fecha2}
        sql.EjecutarProcedure("prc_tbl2", params, vals, tipoParametro, 1)
    End Sub
End Class
