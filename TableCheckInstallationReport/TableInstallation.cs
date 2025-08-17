using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TableCheckInstallationReport
{

    public class TableInstallation
    {
        public string NumberChain;
        public List<Chain> ChainData = new List<Chain>(); 
    }
    public class Chain
    {
        public int StepX;
        public int StepY;
        public string PosDenotation;

        public void Add(int x, int y, string pos, ReportParameters reportParameters)
        {
            StepX = Convert.ToInt32(Convert.ToDouble(x - reportParameters.X) / 1250);
            StepY = Convert.ToInt32(Convert.ToDouble(y - reportParameters.Y) / 1250);
            var inx = pos.IndexOf("-");
            try
            {
                if (pos.Contains("VIA"))
                    PosDenotation = "VIA";
                else
                    PosDenotation = pos.Remove(inx).Trim() + ":" + pos.Remove(0, inx + 1).Trim();
            }
            catch
            {
                MessageBox.Show("Невозможно обработать ячейку, значение - " + pos +". Обратитесь в бюро 911!", "Ошибка столбца поз.обознач.!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        public void Add(double x, double y, string pos, ReportParameters reportParameters)
        {
            StepX = Convert.ToInt32(Convert.ToDouble(x * 1000 - reportParameters.X) / 1250);
            StepY = Convert.ToInt32(Convert.ToDouble(y * 1000 - reportParameters.Y) / 1250);
            PosDenotation = pos;
        }
    }

}
