Imports System.IO
Imports System.Xml.Serialization

Public Class ReadCdrXml
    Public Function ReadCDR(ByVal Ruta As String, NameCDR As String) As String()
        Dim x As New XmlSerializer(GetType(ApplicationResponseType))
        Dim objStreamReader As New StreamReader(Ruta & "\" & NameCDR)
        Dim Valores(10) As String
        Dim SunatCDR As New ApplicationResponseType()
        SunatCDR = x.Deserialize(objStreamReader)
        objStreamReader.Close()
        Dim Descrip As New DescriptionType, AdquirienteType As New PartyIdentificationType
        Valores(0) = SunatCDR.ID.Value
        Valores(1) = SunatCDR.ResponseDate.Value
        Valores(2) = SunatCDR.IssueTime.Value
        Valores(3) = SunatCDR.ResponseDate.Value
        Valores(4) = SunatCDR.ResponseTime.Value
        Dim notas() As NoteType
        If SunatCDR.Note IsNot Nothing Then
            For Each nota As NoteType In SunatCDR.Note
                Valores(10) = Valores(10) + nota.Value & "/"
            Next
        End If
        Dim a(0) As DocumentResponseType
        Dim ax As New DocumentResponseType
        Try
            ax = SunatCDR.DocumentResponse.GetValue(0)
            Valores(5) = ax.Response.ResponseCode.Value
            Valores(6) = ax.Response.ReferenceID.Value
            Descrip = ax.Response.Description.GetValue(0)
            Valores(7) = Replace(Descrip.Value, ",", "-")
            Valores(8) = ax.DocumentReference.ID.Value
            AdquirienteType = ax.RecipientParty.PartyIdentification.GetValue(0)
            Valores(9) = AdquirienteType.ID.Value
        Catch ex As Exception

        End Try

        Return Valores
    End Function
    Public Function ReadCDRbinario(xml_byte As Byte()) As String()

        Dim x As New XmlSerializer(GetType(ApplicationResponseType))

        Dim objStreamReader As New MemoryStream(xml_byte)
        Dim Valores(10) As String
        Dim SunatCDR As New ApplicationResponseType()
        SunatCDR = x.Deserialize(objStreamReader)
        objStreamReader.Close()
        Dim Descrip As New DescriptionType, AdquirienteType As New PartyIdentificationType
        Valores(0) = SunatCDR.ID.Value
        Valores(1) = SunatCDR.ResponseDate.Value
        Valores(2) = SunatCDR.IssueTime.Value
        Valores(3) = SunatCDR.ResponseDate.Value
        Valores(4) = SunatCDR.ResponseTime.Value
        Dim notas() As NoteType
        If SunatCDR.Note IsNot Nothing Then
            For Each nota As NoteType In SunatCDR.Note
                Valores(10) = Valores(10) + nota.Value & "/"
            Next
        End If
        Dim a(0) As DocumentResponseType
        Dim ax As New DocumentResponseType
        ax = SunatCDR.DocumentResponse.GetValue(0)
        Valores(5) = ax.Response.ResponseCode.Value
        Valores(6) = ax.Response.ReferenceID.Value
        Descrip = ax.Response.Description.GetValue(0)
        Valores(7) = Replace(Descrip.Value, ",", "-")
        Valores(8) = ax.DocumentReference.ID.Value
        Try
            AdquirienteType = ax.RecipientParty.PartyIdentification.GetValue(0)
            Valores(9) = AdquirienteType.ID.Value
        Catch ex As Exception

        End Try

        Return Valores
    End Function
End Class
