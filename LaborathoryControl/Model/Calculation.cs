using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace LaborathoryControl.Model
{
    public class Calculation : NotificationEntity
    {
        private ObservableCollection<Data> data;

        private double _Average;
        private double _TMax;
        private double _TMin;
        private double _Variation;
        private double _Variance;
        private double[] _contrArr;

        public double Average
        {
            get { return _Average; }
            set
            {
                if(_Average != value)
                {
                    _Average = value;
                    OnPropertyChanged();
                }
            }
        }
        public double TMax
        {
            get { return _TMax; }
            set
            {
                if(_TMax != value)
                {
                    _TMax = value;
                    OnPropertyChanged();
                }
            }
        }
        public double TMin
        {
            get { return _TMin; }
            set
            {
                if(_TMin != value)
                {
                    _TMin = value;
                    OnPropertyChanged();
                }
            }
        }
        public double Variance
        {
            get { return _Variance; }
            set
            {
                if(_Variance != value)
                {
                    _Variance = value;
                    OnPropertyChanged();
                }
            }
        }
        public double Variation
        {
            get { return _Variation; }
            set
            {
                if(_Variation != value)
                {
                    _Variation = value;
                    OnPropertyChanged();
                }
            }
        }
        public double[] ContrArr
        {
            get { return _contrArr; }
            set
            {
                if(_contrArr != value)
                {
                    _contrArr = value;
                    OnPropertyChanged();
                }
            }
        }

        public Calculation(ObservableCollection<Data> arr)
        {
            data = arr;
            ContrArr = new double[6];

            AverageCalculation();
            Dispersion();
            FactorOfVariation();
            TmaxCalculation();
            TminCalculation();
            ContrMap();
        }

        private void AverageCalculation()
        {
            if (data.Count == 0)
                return;

            double sum = 0;
            foreach(Data elem in data)
            {
                sum += elem.Value.Value;
            }
            Average = Math.Round(sum / data.Count, 4);
        }

        private double Dispersion()
        {
            double intermediate = 0;
            int negative = -1;
            for (int i = 0; i < data.Count; i++)
            {
                data[i].Deviation = (Average - data[i].Value.Value) * negative;
                data[i].SquaredDeviation = Math.Round(Math.Pow((Average - data[i].Value.Value), 2), 4);
                intermediate += data[i].SquaredDeviation;
            }
            //Variance = Math.Round(Math.Sqrt(intermediate / data.Count), 4);
            Variance = Math.Round(intermediate, 4);
            return Variance;
        }

        private void FactorOfVariation()
        {
            if(Average != 0)
                Variation = (Variance / Average) * 100;
        }

        private void TmaxCalculation()
        {
            if(Variance != 0)
                TMax = (Maximum() - Average) / Variance;
        }

        private void TminCalculation()
        {
            if(Variance != 0)
                TMin = (Minimum() - Average) / Variance;
        }

        private void ContrMap()
        {
            for (int i = 1, j = 4; i < 4; i++)
            {
                ContrArr[i - 1] = Average + i * Variance;
                ContrArr[j - 1] = Average - i * Variance;
                j++;
            }
        }

        private double Maximum()
        {
            double maximum = 0;
            foreach(Data obj in data)
            {
                if (maximum < obj.Value.Value)
                    maximum = obj.Value.Value;
            }
            return maximum;
        }

        private double Minimum()
        {
            double min = data[0].Value.Value;
            foreach(Data obj in data)
            {
                if (obj.Value < min)
                    min = obj.Value.Value;
            }
            return min;
        }
    }
}
