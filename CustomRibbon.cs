using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.IO;
using System.Resources;
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Campagna
{
    [ComVisible(true)]
    public class CustomRibbon : ExcelRibbon
    {
        private readonly Application _excel = (Application)ExcelDnaUtil.Application;
        private readonly IRibbonUI _thisRibbon;
        private bool _isStartup = true;
        private readonly ProgressBar _progressBar = new ProgressBar();

        public override string GetCustomUI(string ribbonId)
        {
            string ribbonXml = GetCustomRibbonXML();
            return ribbonXml;
        }

        private string GetCustomRibbonXML()
        {
            string ribbonXml;
            var thisAssembly = typeof(CustomRibbon).Assembly;
            var resourceName = typeof(CustomRibbon).Namespace + ".CustomRibbon.xml";

            using (Stream stream = thisAssembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                ribbonXml = reader.ReadToEnd() ?? throw new MissingManifestResourceException(resourceName);
            }

            return ribbonXml;
        }

        public string GetSelectedItemID(IRibbonControl control)
        {
            if (_isStartup == false)
            {
                return control.Id;
            }
            else
            {
                _isStartup = false;
                Globals._numberIterations = 1000;
                return "ID1000";
            }

        }

        public void GetEditBoxText(IRibbonControl control, string returnedVal)
        {
            
            try
            {
                Globals._randomSeed = int.Parse(returnedVal);
            }
            catch
            {
                Globals._randomSeed = 0;
            }
                
        }
            
        public void OnActionButton(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "generateResults":
                    using (var controller = new ExcelController(_excel, _thisRibbon, _progressBar))
                    {
                        if (Globals._numberIterations == 0)
                            return;
                        if (Globals._randomSeed != 0)
                        {
                            Globals._rand = new Random(Globals._randomSeed);
                        }
                        controller.GetResults();
                        Globals._rand = new Random();
                    }
                    break;
                default:
                    break;
            }

        }

        public void OnActionDropDown(IRibbonControl control, string selectedId, int index)
        {
            switch (control.Id)
            {
                case "dropDownIterations":
                    Globals._numberIterations = int.Parse(selectedId.Substring(2, selectedId.Length - 2));
                    break;
                default:
                    break;
            }
        }

    }
}