using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DataProcessing.Views
{
    /// <summary>
    /// Interaction logic for TextDialog.xaml
    /// </summary>
    public partial class TextDialog : UserControl
    {
        public TextDialog()
        {
            InitializeComponent();

            this.Loaded += (s, e) =>
            {
                Keyboard.Focus(this);
                txbInput.Focus();
            };
        }
    }
}
