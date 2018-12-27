using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace FontImage
{
    public class SVGColor
    {
        private string _strFolder;
        public SVGColor(string strFolder)
        {
            _strFolder = strFolder;
        }

        public void ChangeColor()
        {
            foreach(string strSvg in Directory.GetFiles(_strFolder,"*.svg", SearchOption.AllDirectories))
            {
                XmlDocument doc = new XmlDocument();
                XmlNamespaceManager nsMgr = new XmlNamespaceManager(doc.NameTable);
                nsMgr.AddNamespace("ns", "http://www.w3.org/2000/svg");
                doc.Load(strSvg);
                XmlNode nodeSvg = doc.SelectSingleNode("/ns:svg", nsMgr);
                XmlNode nodeStyle = nodeSvg.SelectSingleNode("/ns:style", nsMgr);
                if (nodeStyle == null)
                {
                    nodeStyle = doc.CreateElement("style");
                    nodeStyle = nodeSvg.InsertBefore(nodeStyle, nodeSvg.ChildNodes[0]);
                    XmlAttribute attribute = doc.CreateAttribute("type");
                    attribute.Value = "text/css";
                    nodeStyle.Attributes.Append(attribute);
                }

                nodeStyle.InnerText = ".st0{fill:#D5261B;}";
                doc.Save(strSvg);
            }
        }
    }
}
