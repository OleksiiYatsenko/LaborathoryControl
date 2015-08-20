using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LaborathoryControl.Enum;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace LaborathoryControl.Model
{
    public class TextDocumentWorker
    {
        private TextRedactors _textRedactor;
        private List<Data> data;
        private Calculation calc;


        public TextDocumentWorker(IEnumerable<Data> values, Calculation calc)
        {
            data = new List<Data>(values);
            this.calc = calc;
            _textRedactor = WhatsInstaled();
        }

        private TextRedactors WhatsInstaled()
        {
            using (RegistryKey microsoft = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Microsoft"))
            {
                if (microsoft != null)
                {
                    RegistryKey word = microsoft.OpenSubKey("Word");

                    if (word != null)
                    {
                        return TextRedactors.Word;
                    }
                }
            }
            string baseKey;
            if (Marshal.SizeOf(typeof(IntPtr)) == 8) 
                baseKey = @"SOFTWARE\Wow6432Node\OpenOffice.org\";
            else
                baseKey = @"SOFTWARE\OpenOffice.org\";
            string key = baseKey + @"Layers\URE\1";

            RegistryKey OpenOffice = Registry.CurrentUser.OpenSubKey(key);
                if (OpenOffice == null)
                    OpenOffice = Registry.LocalMachine.OpenSubKey(key);
                string urePath = OpenOffice.GetValue("UREINSTALLLOCATION") as string;
                if (!string.IsNullOrEmpty(urePath))
                {
                    OpenOffice.Close();
                    return TextRedactors.OpenOffice;
                }
            return TextRedactors.None;
        }

        public void MakeDocument()
        {
            switch(_textRedactor)
            {
                case TextRedactors.Word:
                    {
                        MSWordWorker.WorkWithMsWord(data, calc);
                        return;
                    }
                case TextRedactors.OpenOffice:
                    {
                        return;
                    }
                default:
                    return;
            }
        }
    }
}
