using System;
using System.Threading.Tasks;
using System.Net;
using Titanium.Web.Proxy;
using Titanium.Web.Proxy.EventArguments;
using Titanium.Web.Proxy.Models;
using System.Windows.Forms;

namespace WindowsFormsApp3
{
    public class TWP
    {
        public string redirect;
        public string[] items { get; set; }
        ProxyServer proxyServer = new ProxyServer();
        public void CreateProxySrvr()
        {
            proxyServer.CertificateManager.TrustRootCertificate(true);
            proxyServer.BeforeRequest += OnRequest;
            proxyServer.BeforeResponse += OnResponse;
            proxyServer.ServerCertificateValidationCallback += OnCertificateValidation;
            proxyServer.ClientCertificateSelectionCallback += OnCertificateSelection;
            var explicitEndPoint = new ExplicitProxyEndPoint(IPAddress.Any, 8000, true);
            proxyServer.AddEndPoint(explicitEndPoint);
            proxyServer.Start();
            var transparentEndPoint = new TransparentProxyEndPoint(IPAddress.Any, 8001, true)
            {
                GenericCertificateName = "google.com"
            };

            proxyServer.AddEndPoint(transparentEndPoint);
            foreach (var endPoint in proxyServer.ProxyEndPoints)
                Console.WriteLine("Listening on '{0}' endpoint at Ip {1} and port: {2} ",
                    endPoint.GetType().Name, endPoint.IpAddress, endPoint.Port);
            proxyServer.SetAsSystemHttpProxy(explicitEndPoint);
            proxyServer.SetAsSystemHttpsProxy(explicitEndPoint);
            logger.addLog("proxy has been created successfully \n");
        }
        private async Task OnBeforeTunnelConnectRequest(object sender, TunnelConnectSessionEventArgs e)
        {
            string hostname = e.HttpClient.Request.RequestUri.Host;

            if (hostname.Contains("dropbox.com"))
            {
                e.DecryptSsl = false;
            }
        }
        
        public async Task OnRequest(object sender, SessionEventArgs e)
        {

            
            Console.WriteLine(e.HttpClient.Request.Url);
            try
            {
                /* if (e.HttpClient.Request.RequestUri.AbsoluteUri.Contains("google.com"))
                 {
                     e.Ok("<!DOCTYPE html>" +
                           "<html><body><h1>" +
                           "Website Blocked" +
                           "</h1>" +
                           "<p>Blocked by titanium web proxy.</p>" +
                           "</body>" +
                           "</html>");
                 }*/

                //Redirect example
                foreach (string s in items)
                {
                    if (e.HttpClient.Request.RequestUri.AbsoluteUri.Contains(s))
                    {

                        e.Redirect($"{redirect}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
                logger.addLog("Exception Occurred");
            }

        }
        public async Task OnResponse(object sender, SessionEventArgs e)
        {
            var responseHeaders = e.HttpClient.Response.Headers;
            if (e.HttpClient.Request.Method == "GET" || e.HttpClient.Request.Method == "POST")
            {
                if (e.HttpClient.Response.StatusCode == 200)
                {
                    if (e.HttpClient.Response.ContentType != null && e.HttpClient.Response.ContentType.Trim().ToLower().Contains("text/html"))
                    {
                        byte[] bodyBytes = await e.GetResponseBody();
                        e.SetResponseBody(bodyBytes);

                        string body = await e.GetResponseBodyAsString();
                        e.SetResponseBodyString(body);
                    }
                }
            }

            if (e.UserData != null)
            {
                var request = (Titanium.Web.Proxy.Http.Request)e.UserData;
            }
        }
        public Task OnCertificateValidation(object sender, CertificateValidationEventArgs e)
        {
            if (e.SslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
                e.IsValid = true;
            return Task.FromResult(0);
        }
        public Task OnCertificateSelection(object sender, CertificateSelectionEventArgs e)
        {
            return Task.FromResult(0);
        }
        public void StopProxySrvr()
        {
            proxyServer.BeforeRequest -= OnRequest;
            proxyServer.BeforeResponse -= OnResponse;
            proxyServer.ServerCertificateValidationCallback -= OnCertificateValidation;
            proxyServer.ClientCertificateSelectionCallback -= OnCertificateSelection;
            proxyServer.Stop();
            logger.addLog("proxy has been closed successfully \n\n");
        }
    }
}
