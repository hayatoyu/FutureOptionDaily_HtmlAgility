using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;
using System.Net;


namespace FutureOptionDaily_HtmlAgility.Class
{
    public class BrowserSession_Big5
    {
        private bool isPost;
        private HtmlDocument htmlDoc;

        public CookieCollection Cookies { get; set; }

        public FormElementCollection FormElements { get; set; }
        //public HtmlNodeCollection FormElements;

        public string Get(string url)
        {
            isPost = false;
            CreateWebRequestObject().Load(url);
            return htmlDoc.DocumentNode.InnerHtml;
        }

        public string Post(string url)
        {
            isPost = true;
            CreateWebRequestObject().Load(url, "POST");
            
            return htmlDoc.DocumentNode.InnerHtml;
        }

        private HtmlWeb CreateWebRequestObject()
        {
            HtmlWeb web = new HtmlWeb();
            web.AutoDetectEncoding = false;
            web.OverrideEncoding = Encoding.GetEncoding("big5");
            web.UseCookies = true;
            web.PreRequest = new HtmlWeb.PreRequestHandler(OnPreRequest);
            web.PostResponse = new HtmlWeb.PostResponseHandler(OnAfterResponse);
            web.PreHandleDocument = new HtmlWeb.PreHandleDocumentHandler(OnPreHandleDocument);
            return web;
        }

        protected bool OnPreRequest(HttpWebRequest request)
        {
            AddCookiesTo(request);
            if (isPost)
                AddPostDataTo(request);
            return true;
        }

        protected void OnAfterResponse(HttpWebRequest request, HttpWebResponse response)
        {
            SaveCookiesFrom(response);
        }

        protected void OnPreHandleDocument(HtmlDocument document)
        {
            SaveHtmlDocument(document);
        }

        private void AddCookiesTo(HttpWebRequest request)
        {
            if (Cookies != null && Cookies.Count > 0)
                request.CookieContainer.Add(Cookies);
        }

        private void AddPostDataTo(HttpWebRequest request)
        {
            string payload = FormElements.AssemblyPostPayload();
            byte[] buffer = Encoding.UTF8.GetBytes(payload.ToCharArray());
            request.ContentLength = buffer.Length;
            request.ContentType = "application/x-www-form-urlencoded";
            System.IO.Stream reqStream = request.GetRequestStream();
            reqStream.Write(buffer, 0, buffer.Length);
        }

        private void AddPostDataTo(HttpWebRequest request, Encoding en)
        {
            string payload = FormElements.AssemblyPostPayload();
            byte[] buffer = en.GetBytes(payload.ToCharArray());
            request.ContentLength = buffer.Length;
            request.ContentType = "application/x-www-form-urlencoded";
            System.IO.Stream reqStream = request.GetRequestStream();
            reqStream.Write(buffer, 0, buffer.Length);
        }

        private void SaveCookiesFrom(HttpWebResponse response)
        {
            if (response.Cookies.Count > 0)
            {
                if (Cookies == null)
                    Cookies = new CookieCollection();
                Cookies.Add(response.Cookies);
            }
        }

        private void SaveHtmlDocument(HtmlDocument doc)
        {
            htmlDoc = doc;
            FormElements = new FormElementCollection(htmlDoc);
            //FormElements = htmlDoc.DocumentNode.Descendants("form");
        }
    }
}
