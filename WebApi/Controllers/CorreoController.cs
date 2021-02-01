using CapaDatos;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Web.Http;

namespace WebApi.Controllers
{
    public class CorreoController : ApiController
    {

        ClsConexion go_sql = new ClsConexion();
        public IHttpActionResult Index([FromBody] Datos o)
        {                                
          
            //variables necesarias 
            //  string serie;           
            // string numero;
            string tipo = o.tipo;
            //bd
            string ruc=o.ruc;
            string periodo=o.periodo;
            //tabla comprabante    
            // de otra manera de obtener la ruta del documento
            string [] datos = o.documento.Split('-');
            string ru = datos[0];
            string sub= Path.GetExtension(o.documento);                 
            string ti = sub.Replace(".", "");
            string rutafisica = ru + "//" + ti + "//" + o.documento;
           
            // ClsConexion go_sql = new ClsConexion(ru,periodo);
            string txt = "select SMTP,Puerto,SSL,Credencial,MasterMail,Pws,CC,CCopiaMail,Asunto,CuerpoMail from MailServer";                      
            DataTable dt_mail = new DataTable();
            dt_mail = go_sql.EjecutarConsulta("se", txt).Tables[0];
          
            if (o.correo==null || o.correo=="")
            {
                return BadRequest("Campo correo requerido");
            }
            if (ruc==null || ruc =="")
            {
                return BadRequest("Campo ruc requerido");
            }
            if (periodo==null || periodo =="")
            {
                return BadRequest("Campo periodo requerido");
            }
            string documento = o.documento;
            string documento2 = o.documento2;
            string ruta = ruc + "//" + tipo + "//" + documento;

            string path = System.Web.HttpContext.Current.Server.MapPath("~/sigma/" + ruta);
            string ruta2 = ruc + "//zip"+ "//" + documento2;
            string path2 = System.Web.HttpContext.Current.Server.MapPath("~/sigma/" + ruta2);
            if (dt_mail.Rows.Count>0)
            {
                var toAddress = new MailAddress(o.correo, "To Name");
                DataRow fila = dt_mail.Rows[0];
                var smtp = new SmtpClient
                {
                    Host = fila["SMTP"].ToString(),
                    Port = int.Parse(fila["Puerto"].ToString()),
                    EnableSsl =Boolean.Parse(fila["ssl"].ToString()),
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fila["MasterMail"].ToString(),fila["Pws"].ToString()),//quien lo envia
                };
                var fromAddress = new MailAddress(fila["MasterMail"].ToString(),fila["Asunto"].ToString());
                if (File.Exists(path))
                {
                    using (var message = new MailMessage(fromAddress, toAddress) //quien lo envia  y  a quien se lo envia
                    {
                        Subject = "Comprobante Electronico ",
                        Body    = fila["CuerpoMail"].ToString(),
                    })
                    {
                        message.Attachments.Add(new Attachment(path));   //pdf
                        message.Attachments.Add(new Attachment(path2));  //zip
                        smtp.Send(message);
                    }
                }
                else
                {
                    return BadRequest("No exite el documento");
                }
            }
            else
            {
                return BadRequest("No hay Datos");
            }                  
            return Ok("Enviado con Exito");
              
        }


    }
    public class Datos
    {
        public string correo { get; set; }
        public string ruc { get; set; }
        public string periodo { get; set; }     
            
        public string documento { get; set; }
        public string documento2 { get; set; }
        public string tipo { get; set; }      
        public string nombre_empresa { get; set; }
      
    }
}
