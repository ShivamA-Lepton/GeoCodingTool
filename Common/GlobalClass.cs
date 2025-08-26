using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReverseGeoCoding.Common
{
    public static class GlobalClass
    {
        public static string InputFilepath = string.Empty;
        public static string OutputFilepath = string.Empty;
        public static string PanelMessage = string.Empty;
        public static bool BangValue = true;
        public static string ext = string.Empty;

        public static string messagevalue = string.Empty;
        public static FormChangeEvent ChangeForm = new FormChangeEvent();
        public delegate void FormSelectIndex(int e);
        public class FormChangeEvent
        {
            public event FormSelectIndex ChangeEvent;
            public void OnChangeForm(int e)
            { if (ChangeEvent != null) { ChangeEvent(e); } }
        }

        public static double progressVaue = 0;

    }
}
