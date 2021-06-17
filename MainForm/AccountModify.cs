using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Resources;

namespace MainForm
{
	public partial class AccountModify : Form
	{
		static WindowHelper helper = new WindowHelper();

		string __sto_surn = "nullptr";
		string __sto_forn = "nullptr";
		string __sto_patn = "nullptr";
		string __sto_grp = "nullptr";
		string __sto_dir = "nullptr";
		string __sto_dco = "nullptr";
		string __sto_hsc = "nullptr";
		string __sto_crs = "nullptr";

		public AccountModify()
		{
			InitializeComponent();
			helper.RegistryDownloadData();
			helper.LoadDataFromResources();
			helper.GetDirList(directionsBox);
			helper.GetHsList(hschoolBox);
			courseBox.SelectedIndex = 0;

			if (helper.ui.uiSurname.Length > 0)
				__sto_surn = helper.ui.uiSurname;
			if (helper.ui.uiForename.Length > 0)
				__sto_forn = helper.ui.uiForename;
			if (helper.ui.uiPatronymic.Length > 0)
				__sto_patn = helper.ui.uiPatronymic;
			if (helper.ui.uiGroup.Length > 0)
				__sto_grp = helper.ui.uiGroup;
			if (helper.ui.uiDirection.Length > 0)
				__sto_dir = helper.ui.uiDirection;
			if (helper.ui.uiCode.Length > 0)
				__sto_dco = helper.ui.uiCode;
			if (helper.ui.uiHighSchool.Length > 0)
				__sto_hsc = helper.ui.uiHighSchool;
			if (helper.ui.uiCourse.Length > 0)
				__sto_crs = helper.ui.uiCourse;

			surname.Text = __sto_surn;
			forename.Text = __sto_forn;
			patronymic.Text = __sto_patn;
			group.Text = __sto_grp;

			directionsBox.Text = 
				(helper.ui.uiCode + ' ' + helper.ui.uiDirection);
			hschoolBox.Text = helper.ui.uiHighSchool;
			courseBox.SelectedIndex = (Convert.ToInt32(helper.ui.uiCourse) - 1);

			FormUpdate();
		}

		void FormUpdate()
        {
			resultLabel.Text = __sto_hsc + '\n';
			resultLabel.Text += __sto_dco + ' ' + __sto_dir + '\n';
			resultLabel.Text += __sto_surn + ' ' + __sto_forn + ' ' +
				__sto_patn + '\n';
			resultLabel.Text += "Группа " + __sto_grp + ", " +
				__sto_crs + " курс";
			this.Update();
        }

		private void abortButton_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void saveButton_Click(object sender, EventArgs e)
		{
			helper.RegistryUpdateData();
			helper.RegistryClose();
			this.Close();
		}

		private void directionSelected(object sender, EventArgs e)
		{
			__sto_dco = directionsBox.Text.Substring(0, 8);
			__sto_dir = directionsBox.Text.Substring(9);
			helper.ui.uiCode = __sto_dco;
			helper.ui.uiDirection = __sto_dir;
			FormUpdate();
		}

		private void hsSelected(object sender, EventArgs e)
		{
			__sto_hsc = hschoolBox.Text;
			helper.ui.uiHighSchool = __sto_hsc;
			FormUpdate();
		}

		private void snChanged(object sender, EventArgs e)
		{
			__sto_surn = surname.Text;
			helper.ui.uiSurname = __sto_surn;
			FormUpdate();
		}

		private void fnChanged(object sender, EventArgs e)
		{
			__sto_forn = forename.Text;
			helper.ui.uiForename = __sto_forn;
			FormUpdate();
		}

		private void pnChanged(object sender, EventArgs e)
		{
			__sto_patn = patronymic.Text;
			helper.ui.uiPatronymic = __sto_patn;
			FormUpdate();
		}

		private void gpChanged(object sender, EventArgs e)
		{
			__sto_grp = group.Text;
			helper.ui.uiGroup = __sto_grp;
			FormUpdate();
		}

        private void courseSelected(object sender, EventArgs e)
        {
			__sto_crs = courseBox.Text;
			helper.ui.uiCourse = __sto_crs;
			FormUpdate();
        }
    }
}
