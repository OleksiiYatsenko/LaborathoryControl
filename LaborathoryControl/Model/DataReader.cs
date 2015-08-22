using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace LaborathoryControl.Model
{
    public class DataReader
    {
        public IEnumerable<Data> Read()
        {
            List<Data> dt = new List<Data>();
            using (StreamReader reader = new StreamReader("DataTri.txt"))
            {
                while (!reader.EndOfStream)
                {
                    string str = reader.ReadLine().Replace('.', ',');

                    Data d = new Data(double.Parse(str));
                    dt.Add(d);
                }
            }
            return dt;
        }
    }
}
