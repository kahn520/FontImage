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

        public void ChangeColor(string strOld, string strNew)
        {
            foreach(string strSvg in Directory.GetFiles(_strFolder,"*.svg", SearchOption.AllDirectories))
            {
                string strText = File.ReadAllText(strSvg);
                if (strText.Contains(strOld))
                {
                    strText = strText.Replace(strOld, strNew);
                    File.WriteAllText(strSvg, strText);
                }
            }
        }
    }
}
