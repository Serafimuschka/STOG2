using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Resources;
using System.Collections;
using GeneratorV2;

namespace MainForm
{
	public partial class Window : Form
	{
		public static WindowHelper helper = new WindowHelper();
		public Window()
		{
			InitializeComponent();
			try
			{
				helper.LoadDataFromResources();
				CreateMenu();	
			}
			catch (Exception e)
			{

			}
		}

		// Creates a menu of main form.
		// Returns nothing.
		// Limited using. Call once in form lifetime.
		void CreateMenu()
        {
			/*
			 * Notification for developers
			 * In case of codepage fault, there is a list of menu items:
			 * User <Account, Verification>
			 * About
			 */
			ToolStripMenuItem item = new ToolStripMenuItem("Пользователь");
			item.DropDownItems.Add("Учётная запись");
			item.DropDownItems.Add("Верификация");
			MainMenu.Items.Add(item);

			item = new ToolStripMenuItem("О программе");
            item.Click += About_OnClick;
			MainMenu.Items.Add(item);
        }

		// About menu button click event
        private void About_OnClick(object sender, EventArgs e)
        {
			
        }
    }
	public class WindowHelper
	{
		public Dictionary<string, string> directions;
		public Dictionary<string, string> highSchools;
		static readonly string directionsPath = "Directions.resx";
		static readonly string hSchoolsPath = "Schools.resx";

		// Loads the specific data from .resx files.
		// Returns nothing. Fills a Dictionary<string, string> fields with data.
		// May be unsafe. Use only in the try-catch blocks.
		public void LoadDataFromResources()
		{
			directions = new Dictionary<string, string>();
			highSchools = new Dictionary<string, string>();

			ResXResourceReader reader = new ResXResourceReader(directionsPath);
			foreach (DictionaryEntry d in reader)
			{
				directions.Add(d.Key.ToString(), d.Value.ToString());
			}

			reader = new ResXResourceReader(hSchoolsPath);
			foreach (DictionaryEntry d in reader)
			{
				highSchools.Add(d.Key.ToString(), d.Value.ToString());
			}
		}

		// Transmitts the dictionary data to the ComboBox control.
		// Returns nothing. Specified purpose.
		// Limited using. Call once in form lifetime.
		public void GetDirList(ComboBox cb)
		{
			IDictionaryEnumerator dirEnumerator = directions.GetEnumerator();
			while (dirEnumerator.MoveNext())
			{
				cb.Items.Add(dirEnumerator.Value);
			}
			cb.SelectedIndex = 0;
		}
	}
}
