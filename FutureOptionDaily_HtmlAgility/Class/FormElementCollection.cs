using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using HtmlAgilityPack;

namespace FutureOptionDaily_HtmlAgility.Class
{
    public class FormElementCollection : List<HtmlNode>
    {
        public FormElementCollection(HtmlDocument htmlDoc)
        {
            var inputs = htmlDoc.DocumentNode.Descendants("input");
            AddRange(inputs);
            var menus = htmlDoc.DocumentNode.Descendants("select");
            AddRange(menus);
            var textareas = htmlDoc.DocumentNode.Descendants("textarea");
            AddRange(textareas);
            //foreach(var element in inputs)
            //{
            //    //string name = element.GetAttributeValue("name", "undefined");
            //    //string value = element.GetAttributeValue("value", "");
            //    //if (!name.Equals("undefined"))
            //    //    Add(name, value);
                
            //}
        }

        public string AssemblyPostPayload()
        {
            StringBuilder stbr = new StringBuilder();
            foreach(var element in this)
            {
                //string value = System.Web.HttpUtility.UrlEncode(element.Value);
                //stbr.Append("&" + element.Key + "=" + value);
                string value = System.Web.HttpUtility.UrlEncode(element.GetAttributeValue("value", ""));
                string name = element.GetAttributeValue("name", "undefined");
                stbr.Append("&" + name + "=" + value);
            }
            return stbr.ToString().Substring(1);
        }

        public void SetSelectValue(HtmlNode select,string value)
        {
            // 取消 selected
            foreach (var option in select.ChildNodes)
            {
                if (string.IsNullOrEmpty(option.GetAttributeValue("selected", "undefined")))
                {
                    option.Attributes.Remove("selected");
                    break;
                }
            }

            // 重新設置selected
            foreach (var option in select.ChildNodes)
            {
                if (option.GetAttributeValue("value", "").Equals(value))
                {
                    option.Attributes.Add("selected", "");
                    break;
                }
            }
        }
    }
}
