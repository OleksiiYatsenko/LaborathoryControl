﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LaborathoryControl.Model
{
    public class Data : NotificationEntity
    {
        private int _number;
        private double _value;
        private double _deviation;
        private double _squaredDeviation;

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
        public double Value
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

        public Data(int number)
        {
            Number = number;
        }

        public override bool Equals(object obj)
        {

            return this.Value.Equals((obj as Data).Value);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}