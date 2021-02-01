using CapaDatos;
using CapaNegocios;
using ClsSigmaWs;
using efacturacionClsNuevo;
using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web.Http;
using WebApi.Models;

namespace WebApi.Controllers
{
    [RoutePrefix("api")]
    public class EnvioController : ApiController
    {
        ClsConexion go_sql = new ClsConexion();
        [Route("Envio/Sunat")]
        public IHttpActionResult Index([FromBody] Factura o)
        {          //sdfsf

          //  asa
            NPtentidad entidad = new NPtentidad();
            entidad.identidad = "001";
            entidad = entidad.item(entidad);
            string pws = "";
           // pws = Desencriptar(entidad.pws);
          ///  var a =ServiceReference1.billService
            if (entidad.identidad != "")
            {
                string odt = "";
                string OComprobante = "";
                string FileNamexml = "";
                string ls_webservices="";
                //20532580936
                odt = o.TdSunat;               
                OComprobante = o.Serie + "-" + o.NumeroDocumento;  
                FileNamexml = entidad.ruc + "-" + odt + "-" + OComprobante;
                //ClsSigmaWs.ClsSendSunat  clssunat = new ClsSigmaWs.ClsSendSunat(ls_webservices, entidad.ruc, entidad.user_sol, Desencriptar(entidad.pws_sol));
                ClsSendSunat2 clssunat = new ClsSendSunat2(ls_webservices, entidad.ruc, entidad.user_sol, Desencriptar(entidad.pws_sol));
                string ru = System.Web.HttpContext.Current.Server.MapPath("~/sigma/" + entidad.ruc+"//"+"zip//");
                //string ru = urlGet + "/zip/";
                string strRetorno;
                try
                {                   
                    byte[] bytes = System.IO.File.ReadAllBytes(ru + FileNamexml + ".zip");                 
                    byte[] cdr = clssunat.EnviarDocumento(bytes, FileNamexml);
                    string ruta = System.Web.HttpContext.Current.Server.MapPath("~/sigma/XmlEnviados");
                    File.WriteAllBytes(ruta + "\\R-" + FileNamexml + ".zip", cdr);

                   // FileStream stream = new FileStream(ruta + "\\R-" + FileNamexml + ".zip", FileMode.Create, FileAccess.ReadWrite);
                   // stream.Write(cdr, 0, cdr.Length);
                   // stream.Close();
                    //string fullpath = "";
                    var serverPath = "";
                   // serverPath = System.Web.HttpContext.Current.Server.MapPath("~/sigma/" + nameRuc1 + "/pdf/");
                   // fullpath = Path.Combine(ruta, Path.GetFileName(fileName));
                   // file.SaveAs(fullpath);
                    try
                    {
                        ReadCdrXml ReadCdr = new ReadCdrXml();                       
                        NTbl_CDR tbl_cdr = new NTbl_CDR();
                        string[] valores;
                        try
                        {
                           // valores = ReadCdr.ReadCDRbinario(Unzip(cdr));
                           // {
                           //     var withBlock = tbl_cdr;
                           //     withBlock.NroIdSunat = valores[0]; withBlock.FechaRecepcion = DateTime.Parse(valores[1]); withBlock.HoraRecepcion = valores[2];
                           //     withBlock.FechaCRD = DateTime.Parse(valores[3]); withBlock.HoraCDR = valores[4]; withBlock.Nota = valores[10]; withBlock.NroDocEnviado = valores[6];
                           //     withBlock.CodRecepcion = valores[5]; withBlock.Descriponerror = valores[7]; withBlock.NroDocFirmado = valores[8];
                           //     withBlock.IdAquiriente = valores[9];
                           //   
                           // }
                           // strRetorno = valores[7]; // "Archivo se ha enviado con Exito"
                        }
                        catch (Exception)
                        {
                            throw;
                        }
                    }
                    catch (Exception)
                    {
                        throw;
                    }   
                    //string ruta = System.Web.HttpContext.Current.Server.MapPath("~/sigma/"+"XmlEnviados"); 
                    // https://e-beta.sunat.gob.pe/ol-ti-itcpfegem-beta/billService?wsdl
                }
                catch (Exception)
                {
                    throw;
                }
                 //   Dim clssunat As New ClsSigmaWs.ClsSendSunat(ls_webservices, entidad.ruc, entidad.user_sol, Desencriptar(entidad.pws_sol))
            }            
            return Ok("Enviado con Exito");
        }

        // public byte[] ExtraerByte(byte[] byt)
        // {
        //     string sampleZipFile = @"C:\temp\myzip.zip";
        //     string result;
        //     MemoryStream ms = new MemoryStream(byt);
        //     MemoryStream msxml = new MemoryStream();
        //   
        //     using (MemoryStream memory = new MemoryStream())
        //     {
        //         using (ZipFile zip = ZipFile.Read(ms))
        //         {
        //             ZipEntry ae = zip["myfile.txt"];
        //            
        //             ae.Extract(memory);
        //         }
        //        
        //         using (StreamReader reader = new StreamReader(memory))
        //         {
        //             memory.Seek(0, SeekOrigin.Begin);
        //             result = reader.ReadToEnd();
        //         }
        //     }
        //     
        //     return msxml.ToArray();
        // }
        public byte[] Zip(string textToZip, byte[] byt)
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var zipArchive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                {
                    var demoFile = zipArchive.CreateEntry("zipped.txt");
                    using (var entryStream = demoFile.Open())
                    {
                        using (var streamWriter = new StreamWriter(entryStream))
                        {
                            streamWriter.Write(textToZip);
                        }
                    }
                }
                return memoryStream.ToArray();
            }
        }
        public byte[] Unzip(byte[] zippedBuffer)
        {
            using (var zippedStream = new MemoryStream(zippedBuffer))
            {
                using (var archive = new ZipArchive(zippedStream))
                {
                    var entry = archive.Entries.FirstOrDefault();

                    if (entry != null)
                    {
                        using (var unzippedEntryStream = entry.Open())
                        {
                            using (var ms = new MemoryStream())
                            {
                                unzippedEntryStream.CopyTo(ms);
                                var unzippedArray = ms.ToArray();
                                return unzippedArray.ToArray();
                                //  return Encoding.Default.GetString(unzippedArray);
                            }
                        }
                    }

                    return null;
                }
            }
        }
        public byte[] ExtrarToByte(byte[] a)
        {
            MemoryStream ms = new MemoryStream(a);
            MemoryStream msxml = new MemoryStream();
            //using (ZipFile zip = ZipFile.Read(ms))
            //{
            //    ZipEntry e;
            //    foreach (var e in zip)
            //        e.Extract(msxml);
            //}             
            return msxml.ToArray();
        }
        public string Desencriptar(string aString)
        {
            string st = "";
            int i;
            for (i = 0; i <= aString.Length - 1; i++)
                st += Denc(char.Parse(aString.Substring(i, 1)));
            return st;
        }
        private char Denc(char aChar)
        {
            char ctem;
            bool minuscula = false;
            if (char.IsLower(aChar))
            {
                minuscula = true;
                aChar = char.ToUpper(aChar);
            }
            ctem = '-';
            switch (aChar)
            {
                case object _ when aChar == 'Y':
                    {
                        ctem = 'A';
                        break;
                    }

                case object _ when aChar == 'S':
                    {
                        ctem = 'B';
                        break;
                    }

                case object _ when aChar == 'A':
                    {
                        ctem = 'C';
                        break;
                    }

                case object _ when aChar == 'R':
                    {
                        ctem = 'D';
                        break;
                    }

                case object _ when aChar == 'X':
                    {
                        ctem = 'E';
                        break;
                    }

                case object _ when aChar == 'B':
                    {
                        ctem = 'F';
                        break;
                    }

                case object _ when aChar == 'T':
                    {
                        ctem = 'G';
                        break;
                    }

                case object _ when aChar == 'F':
                    {
                        ctem = 'H';
                        break;
                    }

                case object _ when aChar == 'H':
                    {
                        ctem = 'I';
                        break;
                    }

                case object _ when aChar == 'L':
                    {
                        ctem = 'J';
                        break;
                    }

                case object _ when aChar == 'O':
                    {
                        ctem = 'K';
                        break;
                    }

                case object _ when aChar == 'P':
                    {
                        ctem = 'L';
                        break;
                    }

                case object _ when aChar == 'Ñ':
                    {
                        ctem = 'M';
                        break;
                    }

                case object _ when aChar == 'C':
                    {
                        ctem = 'N';
                        break;
                    }

                case object _ when aChar == 'D':
                    {
                        ctem = 'Ñ';
                        break;
                    }

                case object _ when aChar == 'G':
                    {
                        ctem = 'O';
                        break;
                    }

                case object _ when aChar == 'I':
                    {
                        ctem = 'P';
                        break;
                    }

                case object _ when aChar == 'W':
                    {
                        ctem = 'Q';
                        break;
                    }

                case object _ when aChar == 'Z':
                    {
                        ctem = 'R';
                        break;
                    }

                case object _ when aChar == 'K':
                    {
                        ctem = 'S';
                        break;
                    }

                case object _ when aChar == 'V':
                    {
                        ctem = 'T';
                        break;
                    }

                case object _ when aChar == 'E':
                    {
                        ctem = 'U';
                        break;
                    }

                case object _ when aChar == 'M':
                    {
                        ctem = 'V';
                        break;
                    }

                case object _ when aChar == 'N':
                    {
                        ctem = 'W';
                        break;
                    }

                case object _ when aChar == 'J':
                    {
                        ctem = 'X';
                        break;
                    }

                case object _ when aChar == 'Q':
                    {
                        ctem = 'Y';
                        break;
                    }

                case object _ when aChar == 'U':
                    {
                        ctem = 'Z';
                        break;
                    }

                case object _ when aChar == '(':
                    {
                        ctem = '0';
                        break;
                    }

                case object _ when aChar == '*':
                    {
                        ctem = '1';
                        break;
                    }

                case object _ when aChar == '[':
                    {
                        ctem = '2';
                        break;
                    }

                case object _ when aChar == ')':
                    {
                        ctem = '3';
                        break;
                    }

                case object _ when aChar == '$':
                    {
                        ctem = '4';
                        break;
                    }

                case object _ when aChar == '#':
                    {
                        ctem = '5';
                        break;
                    }

                case object _ when aChar == '.':
                    {
                        ctem = '6';
                        break;
                    }

                case object _ when aChar == ']':
                    {
                        ctem = '7';
                        break;
                    }

                case object _ when aChar == '+':
                    {
                        ctem = '8';
                        break;
                    }

                case object _ when aChar == '{':
                    {
                        ctem = '9';
                        break;
                    }

                case object _ when aChar == '9':
                    {
                        ctem = '&';
                        break;
                    }

                case object _ when aChar == '&':
                    {
                        ctem = '*';
                        break;
                    }

                case object _ when aChar == '6':
                    {
                        ctem = '+';
                        break;
                    }

                case object _ when aChar == '4':
                    {
                        ctem = '.';
                        break;
                    }

                case object _ when aChar == '8':
                    {
                        ctem = '8';
                        break;
                    }

                case object _ when aChar == '2':
                    {
                        ctem = '2';
                        break;
                    }

                case object _ when aChar == '3':
                    {
                        ctem = '3';
                        break;
                    }

                case object _ when aChar == '-':
                    {
                        ctem = '-';
                        break;
                    }

                case object _ when aChar == '5':
                    {
                        ctem = '5';
                        break;
                    }

                case object _ when aChar == '7':
                    {
                        ctem = '7';
                        break;
                    }

                case object _ when aChar == '0':
                    {
                        ctem = '0';
                        break;
                    }

                case object _ when aChar == '?':
                    {
                        ctem = '$';
                        break;
                    }

                case object _ when aChar == '@':
                    {
                        ctem = '#';
                        break;
                    }

                case object _ when aChar == '}':
                    {
                        ctem = '-';
                        break;
                    }

                case object _ when aChar == '1':
                    {
                        ctem = '@';
                        break;
                    }

                case object _ when aChar == '%':
                    {
                        ctem = '%';
                        break;
                    }

                default:
                    {
                        ctem = aChar;
                        break;
                    }
            }
            if (minuscula == true)
                ctem = char.ToLower(ctem);
            return ctem;
        }
    }
    public class Factura
    {
        public string TdSunat { get; set; }
        public string Serie { get; set; }
        public string NumeroDocumento { get; set; }
    }



}
