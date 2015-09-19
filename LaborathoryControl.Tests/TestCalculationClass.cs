using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using LaborathoryControl.ViewModel;
using LaborathoryControl.Model;

namespace LaborathoryControl.Tests
{
    [TestClass]
    public class TestCalculationClass
    {
        private LabControlViewModel SetUp()
        {
            var lab = new LabControlViewModel();

            lab.QuarterValues[0].Value = 1.28;
            lab.QuarterValues[1].Value = 1.31;
            lab.QuarterValues[2].Value = 1.48;
            lab.QuarterValues[3].Value = 1.29;
            lab.QuarterValues[4].Value = 1.30;
            lab.QuarterValues[5].Value = 1.47;
            lab.QuarterValues[6].Value = 1.29;
            lab.QuarterValues[7].Value = 1.34;
            lab.QuarterValues[8].Value = 1.29;
            lab.QuarterValues[9].Value = 1.30;
            lab.QuarterValues[10].Value = 1.47;
            lab.QuarterValues[11].Value = 1.28;
            lab.QuarterValues[12].Value = 1.32;
            lab.QuarterValues[13].Value = 1.37;
            lab.QuarterValues[14].Value = 1.30;
            lab.QuarterValues[15].Value = 1.28;
            lab.QuarterValues[16].Value = 1.31;
            lab.QuarterValues[17].Value = 1.33;
            lab.QuarterValues[18].Value = 1.35;
            lab.QuarterValues[19].Value = 1.28;

            lab.Calculation = new Calculation(lab.QuarterValues);

            return lab;
        }

        [TestMethod]
        public void CalculationValuesTest()
        {
            var LabControl = SetUp();
            
            Assert.AreEqual(LabControl.Calculation.Average, 1.332);
            Assert.AreEqual(LabControl.Calculation.Variance, 0.0657);
            Assert.AreEqual(LabControl.Calculation.Variation, 4.93);
            Assert.AreEqual(LabControl.Calculation.TMin, -0.7915);
            Assert.AreEqual(LabControl.Calculation.TMax, 2.2527);

            double[] contrArrayChackValues = new double[] { 1.3977, 1.4634, 1.5291, 1.2663, 1.2006, 1.1349 };
            for (int i = 0; i < LabControl.Calculation.ContrArr.Length; i++)
                Assert.AreEqual(LabControl.Calculation.ContrArr[i], contrArrayChackValues[i]);
        }

        [TestMethod]
        public void DataValuesTest()
        {
            var LabControl = SetUp();
            int len = LabControl.QuarterValues.Count - 1;

            Assert.IsTrue(LabControl.QuarterValues[len].Value.HasValue);

            Assert.AreEqual(LabControl.QuarterValues[len].Value.Value, 26.64);

            Assert.AreEqual(LabControl.QuarterValues[len].SquaredDeviation, 0.082);
        }
    }
}
