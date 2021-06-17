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
using Microsoft.Win32;

namespace MainForm
{
	public partial class Window : Form
	{
		static WindowHelper helper = new WindowHelper();
		public Window()
		{
			InitializeComponent();
			try
			{
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
			item.DropDownItems[0].Click += Account_Click;
			MainMenu.Items.Add(item);

			item = new ToolStripMenuItem("О программе");
			item.Click += About_OnClick;
			MainMenu.Items.Add(item);
		}

		// About menu button click event
		private void About_OnClick(object sender, EventArgs e)
		{
			
		}

		private void Account_Click(object sender, EventArgs e)
		{
			//Process.Start("GeneratorV2.exe", "--title 1");
			AccountModify form = new AccountModify();
			form.Show();
		}
	}

	public struct UserInfo
	{
		public string uiForename;
		public string uiSurname;
		public string uiPatronymic;
		public bool uiGender;
		public string uiGroup;
		public string uiCourse;
		public string uiDirection;
		public string uiCode;
		public string uiHighSchool;

		public void setGender()
		{
			if (uiSurname != null)
			{
				if (uiSurname[uiSurname.Length - 1] == 'а')
					uiGender = false;
				else uiGender = true;
			}
		}
	}

	public struct DocumentInfo
	{
		public string diDiscipline;
		public string diTheme;
		public string diPrepod;
		public string diPrepodInfo;
		public string diPrepodIniz;
		public string diYear;
		public int diGeneratorMode;
		public int diGeneratorSubMode;
	}

	public class WindowHelper
	{
		public Dictionary<string, string> directions;
		public Dictionary<string, string> highSchools;
		static readonly string directionsPath = "Directions.resx";
		static readonly string hSchoolsPath = "Schools.resx";

		public UserInfo ui = new UserInfo();
		public DocumentInfo di = new DocumentInfo();

		RegistryKey hkcu = Registry.CurrentUser;
		static readonly string rgNull = "registryNullReference";
		static readonly string rgUserForename = "rgUserForename";
		static readonly string rgUserSurname = "rgUserSurname";
		static readonly string rgUserPatronymic = "rgUserPatronymic";
		static readonly string rgUserGender = "rgUserGender";
		static readonly string rgUserGroup = "rgUserGroup";
		static readonly string rgUserCourse = "rgUserCourse";
		static readonly string rgUserHighSchool = "rgUserHighSchool";
		static readonly string rgUserDirectionName = "rgUserDirectionName";
		static readonly string rgUserDirectionCode = "rgUserDirectionCode";
		static readonly string rgDocLastPrepod = "rgDocLastPrepod";
		static readonly string rgDocLastPrepodInfo = "rgDocLastPrepodInfo";
		static readonly string rgDocLastPrepodIniz = "rgDocLastPrepodIniz";
		static readonly string rgDocLastTheme = "rgDocLastTheme";
		static readonly string rgDocLastDisc = "rgDocLastDisc";
		static readonly string rgDocYear = "rgDocYear";
		static readonly string rgDocLastMode = "rgDocLastMode";
		static readonly string rgDocLastSubMode = "rgDocLastSubMode";

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
				cb.Items.Add
				(
					dirEnumerator.Key.ToString() + ' ' + dirEnumerator.Value
				);
			}
			cb.SelectedIndex = 0;
		}

		// Transmitts the dictionary data to the ComboBox control.
		// Returns nothing. Specified purpose.
		// Limited using. Call once in form lifetime.
		public void GetHsList(ComboBox cb)
		{
			IDictionaryEnumerator dirEnumerator = highSchools.GetEnumerator();
			while (dirEnumerator.MoveNext())
			{
				cb.Items.Add(dirEnumerator.Value);
			}
			cb.SelectedIndex = 0;
		}

		// Transfer data from the registry.
		// Returns nothing. Unlimited use.
		public void RegistryDownloadData()
		{
			RegistryKey stog = hkcu.OpenSubKey("STOGv2");

			if (stog == null) RegistryUpdateData();
			else
			{
				ui.uiForename = stog.GetValue(rgUserForename).ToString();
				ui.uiSurname = stog.GetValue(rgUserSurname).ToString();
				ui.uiPatronymic = stog.GetValue(rgUserPatronymic).ToString();

				ui.uiGender = Convert.ToBoolean
				(
					stog.GetValue(rgUserGender).ToString()
				);
				ui.uiGroup = stog.GetValue(rgUserGroup).ToString();
				ui.uiCourse = stog.GetValue(rgUserCourse).ToString();
				ui.uiCode = stog.GetValue(rgUserDirectionCode).ToString();
				ui.uiDirection = stog.GetValue(rgUserDirectionName).ToString();
				ui.uiHighSchool = stog.GetValue(rgUserHighSchool).ToString();

				di.diYear = stog.GetValue(rgDocYear).ToString();
				di.diGeneratorMode = Convert.ToInt32
				(
					stog.GetValue(rgDocLastMode).ToString()
				);
				di.diGeneratorSubMode = Convert.ToInt32
				(
					stog.GetValue(rgDocLastSubMode)
				);

				di.diDiscipline = stog.GetValue(rgDocLastDisc).ToString();
				di.diTheme = stog.GetValue(rgDocLastTheme).ToString();
				di.diPrepod = stog.GetValue(rgDocLastPrepod).ToString();
				di.diPrepodInfo = stog.GetValue(rgDocLastPrepodInfo).ToString();
				di.diPrepodIniz = stog.GetValue(rgDocLastPrepodIniz).ToString();

				stog.Close();
			}
		}

		// Transfer data to the registry.
		// Returns nothing. Unlimited use.
		public void RegistryUpdateData()
		{
			RegistryKey stog = hkcu.OpenSubKey("STOGv2", true);

			if (stog == null)
			{
				stog = hkcu.CreateSubKey("STOGv2");

				stog.SetValue(rgUserForename, rgNull);
				stog.SetValue(rgUserSurname, rgNull);
				stog.SetValue(rgUserPatronymic, rgNull);

				ui.setGender();
				stog.SetValue(rgUserGender, 1);
				stog.SetValue(rgUserGroup, rgNull);
				stog.SetValue(rgUserCourse, rgNull);
				stog.SetValue(rgUserDirectionCode, rgNull);
				stog.SetValue(rgUserDirectionName, rgNull);
				stog.SetValue(rgUserHighSchool, rgNull);

				stog.SetValue(rgDocYear, rgNull);
				stog.SetValue(rgDocLastMode, 1);
				stog.SetValue(rgDocLastSubMode, 1);

				stog.SetValue(rgDocLastDisc, rgNull);
				stog.SetValue(rgDocLastTheme, rgNull);
				stog.SetValue(rgDocLastPrepod, rgNull);
				stog.SetValue(rgDocLastPrepodInfo, rgNull);
				stog.SetValue(rgDocLastPrepodIniz, rgNull);

				stog.Close();
			}
			else
			{
				stog.SetValue(rgUserForename, ui.uiForename);
				stog.SetValue(rgUserSurname, ui.uiSurname);
				stog.SetValue(rgUserPatronymic, ui.uiPatronymic);

				ui.setGender();
				stog.SetValue(rgUserGender, ui.uiGender);
				stog.SetValue(rgUserGroup, ui.uiGroup);
				stog.SetValue(rgUserCourse, ui.uiCourse);
				stog.SetValue(rgUserDirectionCode, ui.uiCode);
				stog.SetValue(rgUserDirectionName, ui.uiDirection);
				stog.SetValue(rgUserHighSchool, ui.uiHighSchool);

				stog.SetValue(rgDocYear, di.diYear);
				stog.SetValue(rgDocLastMode, di.diGeneratorMode);
				stog.SetValue(rgDocLastSubMode, di.diGeneratorSubMode);

				stog.SetValue(rgDocLastDisc, di.diDiscipline);
				stog.SetValue(rgDocLastTheme, di.diTheme);
				stog.SetValue(rgDocLastPrepod, di.diPrepod);
				stog.SetValue(rgDocLastPrepodInfo, di.diPrepodInfo);
				stog.SetValue(rgDocLastPrepodIniz, di.diPrepodIniz);

				stog.Close();
			}
		}

		public void RegistryClose()
        {
			this.hkcu.Close();
        }
	}
}
