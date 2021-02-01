using ClsSigmaWs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;

namespace WebApi.Models
{
    public class ClsSendSunat2
    {
        private ServiceReference1.billServiceClient wService;
        //private 
        // Dim service As WebApi
        public ClsSendSunat2()
        {
            wService = new ServiceReference1.billServiceClient();
            ServicePointManager.UseNagleAlgorithm = true;
            ServicePointManager.Expect100Continue = false;
            ServicePointManager.CheckCertificateRevocationList = true;
        }
        public ClsSendSunat2(string endpointurl, string Ruc, string UserName, string pws)
        {
            // ********* bloque
            ServicePointManager.UseNagleAlgorithm = true;
            ServicePointManager.Expect100Continue = false;
            ServicePointManager.CheckCertificateRevocationList = true;
            if (endpointurl != "")
            {
                var behavior = new PasswordDigestBehavior(Ruc + UserName, pws);
                wService = new ServiceReference1.billServiceClient("BillServicePort", endpointurl);
                wService.Endpoint.Behaviors.Add(behavior);
            }
        }
        public void openWs()
        {
            wService.Open();
        }
        public void CerrarWS()
        {
            wService.Close();
        }

        public byte[] EnviarDocumentoBynary(byte[] archivo, string FileNameXml)
        {
            byte[] returnbyte = null;
            try
            {
                wService.Open();
                returnbyte = wService.sendBill(FileNameXml + ".zip", archivo, null/* TODO Change to default(_) if this is not a reference type */);
                wService.Close();
            }
            catch (Exception ex)
            {
            }
            return returnbyte;
        }
        public byte[] EnviarDocumento(byte[] archivo, string FileNameXml)
        {
            byte[] returnbyte = null;
            try
            {
                wService = new ServiceReference1.billServiceClient();
                returnbyte = wService.sendBill(FileNameXml + ".zip", archivo, "");
            }
            catch (Exception ex)
            {
                returnbyte = Encoding.ASCII.GetBytes(ex.Message);
            }
            return returnbyte;
        }

    }
}