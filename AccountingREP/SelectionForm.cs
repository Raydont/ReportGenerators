using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex.Reporting;

namespace AccountingREPReport
{
	public partial class SelectionForm : Form
	{
		private IReportGenerationContext _context;
		public НастройкиОтчета reportParameters = new НастройкиОтчета();	 

		public SelectionForm()
		{
			InitializeComponent();
		}

		public void Init(IReportGenerationContext context)
		{
			_context = context;
		}

		private void checkBoxQuarter_CheckedChanged(object sender, EventArgs e)
		{
			comboBoxQuarter.Enabled = checkBoxQuarter.Checked ? true : false;
			dtpStartPeriod.Enabled = checkBoxQuarter.Checked ? false : true;
			dtpEndPeriod.Enabled = checkBoxQuarter.Checked ? false : true;

			TrySetQuarter();
		}

		private void SelectionForm_Load(object sender, EventArgs e)
		{
			comboBoxQuarter.SelectedIndex = 0;
			comboBoxQuarter.Enabled = false;
		}

		private void comboBoxQuarter_SelectedValueChanged(object sender, EventArgs e)
		{
			TrySetQuarter();
		}

		private void buttonОК_Click(object sender, EventArgs e)
		{
			DialogResult = DialogResult.OK;
		}

		private void TrySetQuarter()
		{
			if (checkBoxQuarter.Checked)
			{
				var квартал = comboBoxQuarter.SelectedIndex + 1;
				reportParameters.НачалоПериода = new DateTime(DateTime.Now.Year, 3 * квартал - 2, 1);
				reportParameters.КонецПериода = reportParameters.НачалоПериода.AddMonths(3).AddDays(-1);
				
				dtpStartPeriod.Value = reportParameters.НачалоПериода;
				dtpEndPeriod.Value = reportParameters.КонецПериода;
			}
		}
	}
}