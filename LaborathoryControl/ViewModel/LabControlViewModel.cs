using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LaborathoryControl.Model;
using System.Collections.ObjectModel;
using System.Windows.Input;
using GalaSoft.MvvmLight.CommandWpf;
using OxyPlot;
using OxyPlot.Series;

namespace LaborathoryControl.ViewModel
{
    public class LabControlViewModel : NotificationEntity
    {
        private ObservableCollection<Data> _quarterValues;
        public ObservableCollection<Data> QuarterValues
        {
            get { return _quarterValues; }
            set
            {
                if(_quarterValues != null &&_quarterValues != value)
                {
                    _quarterValues = value;
                    OnPropertyChanged();
                }
            }
        }

        private Calculation _calculation;
        public Calculation Calculation
        {
            get { return _calculation; }
            set
            {
                if(_calculation != value)
                {
                    _calculation = value;
                    OnPropertyChanged();
                }
            }
        }

        private PlotModel _Model;
        public PlotModel Model 
        { 
            get { return _Model; }
            private set 
            {
                _Model = value;
                OnPropertyChanged();
            } 
        }

        public ICommand StartCommand { get; set; }
        public ICommand GenerateWordDocCommand { get; set; }
        public ICommand CloseCommand { get; set; }

        public LabControlViewModel()
        {
            _quarterValues = new ObservableCollection<Data>();
            for (int i = 0; i < 20;)
                QuarterValues.Add(new Data(++i));
            StartCommand = new RelayCommand(Start);
            GenerateWordDocCommand = new RelayCommand(GenerateMSWordDoc);
            CloseCommand = new RelayCommand(Close);
        }

        private void Start()
        {
            Calculation = new Calculation(QuarterValues);

            GetPlotModel();
        }

        void GetPlotModel()
        {
            var series = new LineSeries { Title = "Данные по анализам", MarkerType = MarkerType.Circle };
            foreach(Data d in QuarterValues)
            {
                series.Points.Add(new DataPoint(d.Number, d.Value));
            }
            PlotModel tmp = new PlotModel();
            tmp.Series.Add(series);
            this.Model = tmp;
        }

        private void GenerateMSWordDoc()
        {
            TextDocumentWorker tdw = new TextDocumentWorker(QuarterValues, Calculation);
            tdw.MakeDocument();
        }

        private void Close()
        {
            App.Current.MainWindow.Close();
        }
    }
}
