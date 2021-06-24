using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace InfoBrowser
{
    public partial class Browser : Form
    {
        public Browser()
        {
            InitializeComponent();
            string dir = Directory.GetCurrentDirectory();
            viewer.Url = new Uri(String.Format("file:///{0}/index.html", dir));
            Text = viewer.DocumentTitle;
        }
    }
}
