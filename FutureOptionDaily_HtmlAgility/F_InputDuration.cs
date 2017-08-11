using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FutureOptionDaily_HtmlAgility
{
    public partial class F_InputDuration : Form
    {
        Form1 F;
        public F_InputDuration()
        {
            InitializeComponent();
        }

        public F_InputDuration(Form1 F)
        {
            InitializeComponent();
            this.F = F;
        }

        private void btn_Confirm_Click(object sender, EventArgs e)
        {
            DateTime StartDate = dateTimePicker_Start.Value;
            DateTime EndDate = dateTimePicker_EndDate.Value;
            StringBuilder stbr = new StringBuilder();

            if(F.duration == null)
                F.duration = new List<DateTime>();

            while (StartDate <= EndDate)
            {
                if(!F.duration.Exists(x => x.Equals(StartDate)))
                    F.duration.Add(StartDate);
                StartDate = StartDate.AddDays(1);
            }

            if (F.duration.Count > 0)
            {
                foreach (DateTime d in F.duration)
                {
                    stbr.AppendLine(d.ToShortDateString());
                }
                F.rtbSelectedDate.Text = stbr.ToString();
            }
            this.Close();
        }
    }
}
