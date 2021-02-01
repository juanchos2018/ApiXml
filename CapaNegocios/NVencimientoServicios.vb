Public Class NVencimientoServicios

    Public Property ruc As String
    Public Property razonsocial As String
    Public Property FechaVencimientoOSE As DateTime?
    Public Property FechaVencimientoHosting As DateTime?
    Public Property FechaVencimientoCertificado As DateTime?
    Public Property fechavencimientodni As DateTime?
    Public Property DiasAviso As Byte
    Public Property vencidoOSE As Boolean
    Public Property vencidoCertificado As Boolean
    Public Property vencidoHosting As Boolean
    Public Property vencidodni As Boolean
    Public Property mensajeOSE As String
    Public Property mensajeCertificado As String
    Public Property mensajeHosting As String
    Public Property mensajedni As String
End Class
Public Class NEmpresaVencimiento

    Public Property ruc As String
    Public Property aliass As String
    Public Property mensajeVencimiento As String
    Public Property mensajeAdvertencia As String
    Public Property idservicio As String
    Public Property descripcion As String
    Public Property obligatorio As String
    Public Property vencimiento As String
    Public Property diasaviso As String
    Public Property diasvencido As Int16
    Public Property vencido As String
    Public Property advertencia As String

End Class
Public Class NMensajeServicios
    Public Property mensaje As String
    Public Property cantidad As Int16
    Public Property errores As Boolean
    Public Property data As New List(Of NEmpresaVencimiento)
End Class

