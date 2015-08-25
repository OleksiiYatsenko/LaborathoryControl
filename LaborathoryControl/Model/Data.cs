using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LaborathoryControl.Model
{
    public class Data : NotificationEntity
    {
        private static int counter;

        private int _number;
        private double? _value;
        private double _deviation;
        private double _squaredDeviation;
        private DateTime _date;

        public DateTime Date
        {
            get { return _date; }
            set
            {
                _date = value;
                OnPropertyChanged();
            }
        }
        public int Number
        {
            get { return _number; }
            set
            {
                if(value != _number)
                {
                    _number = value;
                    OnPropertyChanged();
                    
                }
            }
        }
        public double? Value
        {
            get { return _value; }
            set
            {
                if(_value != value)
                {
                    _value = value;
                    OnPropertyChanged();
                }
            }
        }
        public double Deviation
        {
            get { return _deviation; }
            set
            {
                if(_deviation != value)
                {
                    _deviation = value;
                    OnPropertyChanged();
                }
            }
        }
        public double SquaredDeviation
        {
            get { return _squaredDeviation; }
            set
            {
                if(_squaredDeviation != value)
                {
                    _squaredDeviation = value;
                    OnPropertyChanged();
                }
            }
        }

        static Data()
        {
            counter = 1;
        }

        public Data()
        {
            Number = counter++;
        }

        public Data(double value)
        {
            Number = counter++;
            this.Value = value;
        }
    }
}
