using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    internal class ExcelResources
    {
        // Singleton Implementation
        private ExcelResources() { }
        private static ExcelResources _instance;
        public static ExcelResources GetInstance()
        {
            if (_instance == null)
            {
                _instance = new ExcelResources();
            }
            return _instance;
        }

        // Color dictionary for table decorations
        public Dictionary<string, Color> Colors { get; private set; } = new Dictionary<string, Color>()
        {
            {"DarkBlue", Color.FromArgb(75, 177, 250) },
            {"DarkOrange", Color.FromArgb(250, 148, 75) },
            {"DarkRed", Color.FromArgb(250, 92, 75) },
            {"DarkGray", Color.FromArgb(230, 229, 225) },
            {"DarkGreen", Color.FromArgb(181, 250, 97) },
            {"Blue", Color.FromArgb(148, 216, 255) },
            {"Orange", Color.FromArgb(255, 187, 148) },
            {"Green", Color.FromArgb(202, 255, 138) },
            {"Yellow", Color.FromArgb(250, 228, 102) },
            {"Gray", Color.FromArgb(230, 229, 225) },
            {"Red", Color.FromArgb(255, 157, 148) }
        };
    }
}
