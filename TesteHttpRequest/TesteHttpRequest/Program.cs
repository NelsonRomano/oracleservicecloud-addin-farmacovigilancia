using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace WebServiceCall {

    public class RightnowWSSuperaTeste {

        public static string realizarQueryCampo(string query) {
            var _url = "https://superarx.custhelp.com/cgi-bin/superarx.cfg/services/soap";
            var _action = "https://superarx.custhelp.com/cgi-bin/superarx.cfg/services/soap?op=QueryCSV";
            XmlDocument soapEnvelopeXml = CreateSoapEnvelopeCampo(query);
            HttpWebRequest webRequest = CreateWebRequest(_url, _action);
            InsertSoapEnvelopeIntoWebRequest(soapEnvelopeXml, webRequest);
            IAsyncResult asyncResult = webRequest.BeginGetResponse(null, null);
            asyncResult.AsyncWaitHandle.WaitOne();
            string soapResult;
            using (WebResponse webResponse = webRequest.EndGetResponse(asyncResult)) {
                using (StreamReader rd = new StreamReader(webResponse.GetResponseStream())) {
                    soapResult = rd.ReadToEnd();
                }
                return soapResult;
            }
        }

        public static string realizarQueryTabela(string query) {
            var _url = "https://superarx.custhelp.com/cgi-bin/superarx.cfg/services/soap";
            var _action = "https://superarx.custhelp.com/cgi-bin/superarx.cfg/services/soap?op=QueryCSV";
            XmlDocument soapEnvelopeXml = CreateSoapEnvelopeTabela(query);
            HttpWebRequest webRequest = CreateWebRequest(_url, _action);
            InsertSoapEnvelopeIntoWebRequest(soapEnvelopeXml, webRequest);
            IAsyncResult asyncResult = webRequest.BeginGetResponse(null, null);
            asyncResult.AsyncWaitHandle.WaitOne();
            string soapResult;
            using(WebResponse webResponse = webRequest.EndGetResponse(asyncResult)) {
                using(StreamReader rd = new StreamReader(webResponse.GetResponseStream())) {
                    soapResult = rd.ReadToEnd();
                }
                return soapResult;
            }
        }

        private static HttpWebRequest CreateWebRequest(string url, string action) {
            HttpWebRequest webRequest = (HttpWebRequest) WebRequest.Create(url);
            webRequest.Headers.Add("SOAPAction", action);
            webRequest.ContentType = "text/xml;charset=\"utf-8\"";
            webRequest.Accept = "text/xml";
            webRequest.Method = "POST";
            return webRequest;
        }

        private static XmlDocument CreateSoapEnvelopeCampo(string query) {
            XmlDocument soapEnvelop = new XmlDocument();
            soapEnvelop.LoadXml(popularRequisicaoCampo(query));
            return soapEnvelop;
        }

        private static XmlDocument CreateSoapEnvelopeTabela(string query) {
            XmlDocument soapEnvelop = new XmlDocument();
            soapEnvelop.LoadXml(popularRequisicaoTabela(query));
            return soapEnvelop;
        }

        private static void InsertSoapEnvelopeIntoWebRequest(XmlDocument soapEnvelopeXml, HttpWebRequest webRequest) {
            using (Stream stream = webRequest.GetRequestStream()) {
                soapEnvelopeXml.Save(stream);
            }
        }

        public static string popularRequisicaoCampo(string query) {
            string xml = "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:v1=\"urn:messages.ws.rightnow.com/v1_3\">"
                            + "<soapenv:Header>"
                            + "   <ns7:ClientInfoHeader xmlns:ns7=\"urn:messages.ws.rightnow.com/v1_3\" soapenv:mustUnderstand=\"0\">"
                            + "       <ns7:AppID>Basic Query CSV</ns7:AppID>"
                            + "   </ns7:ClientInfoHeader>"
                            + "   <wsse:Security xmlns:wsse=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd\" mustUnderstand=\"1\">"
                            + "       <wsse:UsernameToken>"
                            + "           <wsse:Username>integrador</wsse:Username>"
                            + "           <wsse:Password Type=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText\">Supera123@</wsse:Password>"
                            + "       </wsse:UsernameToken>"
                            + "   </wsse:Security>"
                            + "</soapenv:Header>"
                            + "   <soapenv:Body>"
                            + "      <v1:QueryCSV>"
                            + "         <v1:Query>" + query + "</v1:Query>"
                            + "         <v1:PageSize>1</v1:PageSize>"
                            + "         <v1:Delimiter>;</v1:Delimiter>"
                            + "         <v1:ReturnRawResult>false</v1:ReturnRawResult>"
                            + "         <v1:DisableMTOM>true</v1:DisableMTOM>"
                            + "      </v1:QueryCSV>"
                            + "   </soapenv:Body>"
                            + "</soapenv:Envelope>";
            return xml;
        }

        public static string popularRequisicaoTabela(string query) {
            string xml = "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:v1=\"urn:messages.ws.rightnow.com/v1_3\">"
                            + "<soapenv:Header>"
                            + "   <ns7:ClientInfoHeader xmlns:ns7=\"urn:messages.ws.rightnow.com/v1_3\" soapenv:mustUnderstand=\"0\">"
                            + "       <ns7:AppID>Basic Query CSV</ns7:AppID>"
                            + "   </ns7:ClientInfoHeader>"
                            + "   <wsse:Security xmlns:wsse=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd\" mustUnderstand=\"1\">"
                            + "       <wsse:UsernameToken>"
                            + "           <wsse:Username>integrador</wsse:Username>"
                            + "           <wsse:Password Type=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText\">Supera123@</wsse:Password>"
                            + "       </wsse:UsernameToken>"
                            + "   </wsse:Security>"
                            + "</soapenv:Header>"
                            + "   <soapenv:Body>"
                            + "      <v1:QueryCSV>"
                            + "         <v1:Query>" + query + "</v1:Query>"
                            + "         <v1:PageSize>-1</v1:PageSize>"
                            + "         <v1:Delimiter>;</v1:Delimiter>"
                            + "         <v1:ReturnRawResult>false</v1:ReturnRawResult>"
                            + "         <v1:DisableMTOM>true</v1:DisableMTOM>"
                            + "      </v1:QueryCSV>"
                            + "   </soapenv:Body>"
                            + "</soapenv:Envelope>";
            return xml;
        }

    }

}