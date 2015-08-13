using System;
using System.Linq;

namespace forlab
{
    class Calculation
    {

        float Average, Tmax, Tmin, Variation, Variance;
            

        public float AverageCalculation(float []myArray)
        {
           Average = myArray.Sum() / myArray.Length;   
           return Average;
 
        }


        public float Dispersion(float[] myArray,ref float[] deviationArray,ref float[] squareDeviationArray)
        {
            double intermediate = 0;
            int negative=-1;
            for(int i = 0; i < myArray.Length; i++)
            {
                deviationArray[i] = (Average - myArray[i]) * negative;
                squareDeviationArray[i] = System.Convert.ToSingle(Math.Pow((Average - myArray[i]), 2));
                intermediate += Math.Pow( (Average - myArray[i]), 2);
            }
            Variance = System.Convert.ToSingle(Math.Sqrt(intermediate / (myArray.Length - 1)));
            return Variance;
        }

        public float FactorOfVariation()
        {

            Variation = (Variance / Average) * 100;

            return Variation;
        }

        public float TmaxCalculation(float []myArray)
        {
            float max;
            max = myArray[19];
            Tmax = (max - Average) / Variance;
            return Tmax;
        }

        public float TminCalculation(float []myArray)
        {
            float min;
            min = myArray[0];
            Tmin = (min - Average) / Variance;
            return Tmin;
        }

        public void ContrMap(ref float[] contrArray)
        {
            for(int i = 1, j=4; i<4; i++)
            {
                contrArray[i-1] = Average + i * Variance;
                contrArray[j-1] = Average - i * Variance;
                j ++;
            }
                  
        }
    }
}
