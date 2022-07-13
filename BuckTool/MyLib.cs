using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;

namespace BuckTool
{
    public class MyLib
    {

        public static List<double> DGData(DataGridView dataGrid)
        {
            List<double> data = new List<double>();

            for(int row_idx = 0; row_idx < dataGrid.RowCount; row_idx++)
            {
                double start = Convert.ToDouble(dataGrid[0, row_idx].Value);
                double step = Convert.ToDouble(dataGrid[1, row_idx].Value);
                double stop = Convert.ToDouble(dataGrid[2, row_idx].Value);
                double res = 0;
                for (int idx = 0; res < stop; idx++)
                {
                    res = start + step * idx;
                    data.Add(res);
                }
            }
            return data;
        }



    }
}
