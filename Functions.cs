using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;


namespace Campagna.Functions
{

    public static class Input
    {

        [ExcelFunction(Description = "Generates  a triangular distribution with given minimum, most likely and max", IsVolatile = true, IsMacroType =true)]
        public static object CampagnaInputTriangular(double min, double ml, double max)
        {
            double dblLowerRange = ml - min;
            double dblHigherRange = max - ml;
            double dblTotalRange = max - min;
            double U;
            double k;

            if ((min > ml) | (max < ml))
                return ExcelError.ExcelErrorNum;

            if ((min == ml) & (ml == max))
                return ml;

            object caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;

            U = (double)GetRandomNumber(caller, Globals._currentIteration);

            if (U <= (dblLowerRange / dblTotalRange))
                k = min + Math.Sqrt(U * dblLowerRange * dblTotalRange);
            else
                k = max - Math.Sqrt((1 - U) * dblHigherRange * dblTotalRange);
            return k;

        }

        [ExcelFunction(Description = "Generates a Bernoulli distribution with a given probability ", IsVolatile = true, IsMacroType =true)]
        public static object CampagnaInputBernoulli(double probability)
        {
            double U;

            if ((probability < 0) || (probability > 1))
                return ExcelError.ExcelErrorNum;
            else
            {
                U = Globals._rand.NextDouble();
                if (U > probability)
                    return 0;
                else
                    return 1;
            }

        }

        public static object GetRandomNumber(object range, int k)
        {
            if (Globals._randomSeed == 0 || Globals._randomNumbers.Count==0)
            {
                return Globals._rand.NextDouble();
            }
            else
            {

                string address = (string)XlCall.Excel(XlCall.xlfReftext, range, true);
                List<Slice> sliceList = Globals._randomNumbers.GetSliceList(address);

                return sliceList[k - 1].Data;

            }

        }
    }

    public static class Output
    {

        [ComVisible(true)]
        [ExcelFunction(Description = "Returns the percentile of the given distribution", IsMacroType = true)]
        public static object CampagnaOutputPercentile([ExcelArgument(AllowReference = true)] object range, double percentile)
        {
            if (Globals._isSimulating == true)
                return ExcelError.ExcelErrorGettingData;
            if (Globals._sliceData.Count == 0)
                return ExcelError.ExcelErrorValue;

            try
            {
                string address = (string)XlCall.Excel(XlCall.xlfReftext, range, true);
                List<Slice> sliceList = Globals._sliceData.GetSliceList(address);
                sliceList.Sort();

                int N = sliceList.Count;
                double n = (N - 1) * percentile + 1;
                if (n == 1d) return sliceList[0].Data;
                else if (n == N) return sliceList[N - 1].Data;
                else
                {
                    int k = (int)n;
                    double d = n - k;
                    return sliceList[k - 1].Data + d * (sliceList[k].Data - sliceList[k - 1].Data);
                }

            }
            catch (ArgumentNullException)
            {
                return ExcelError.ExcelErrorValue;
            }

        }

        [ComVisible(true)]
        [ExcelFunction(Description = "Returns the kth data slice of the distribution in order that they were sampled.", IsMacroType = true)]
        public static object CampagnaOutputSingleSlice([ExcelArgument(AllowReference = true)] object range, int k)
        {

            if (Globals._isSimulating == true)
                return ExcelError.ExcelErrorGettingData;
            if (Globals._sliceData.Count == 0)
                return ExcelError.ExcelErrorValue;

            try
            {
                string address = (string)XlCall.Excel(XlCall.xlfReftext, range, true);
                List<Slice> sliceList = Globals._sliceData.GetSliceList(address);

                return sliceList[k - 1].Data;
            }
            catch (ArgumentNullException)
            {
                return ExcelError.ExcelErrorValue;
            }
        }

        [ComVisible(true)]
        [ExcelFunction(Description = "Calculates the arithmetic mean of the given distribution", IsMacroType = true)]
        public static object CampagnaOutputMean([ExcelArgument(AllowReference = true)] object range)
        {
            if (Globals._isSimulating == true)
                return ExcelError.ExcelErrorGettingData;
            if (Globals._sliceData.Count == 0)
                return ExcelError.ExcelErrorValue;

            double total = 0;

            try
            {
                string address = (string)XlCall.Excel(XlCall.xlfReftext, range, true);
                List<Slice> sliceList = Globals._sliceData.GetSliceList(address);

                foreach (Slice sliceDatum in sliceList)
                {
                    total += sliceDatum.Data;
                }

                return total / sliceList.Count;

            }
            catch (ArgumentNullException)
            {
                return ExcelError.ExcelErrorValue;
            }

        }

        [ComVisible(true)]
        [ExcelFunction(Description = "Outputs various properties of the last analysis that has been run", IsVolatile = true)]
        public static object CampagnaOutputProperties(int property)
        {
            switch (property)
            {
                case 1:
                    if (Globals._numberIterations != 0)
                    {
                        return Globals._numberIterations;
                    }
                    return ExcelError.ExcelErrorValue;

                case 2:
                    if (Globals._start != DateTime.MinValue)
                    {
                        return Globals._start;
                    }
                    return Globals._start;

            }
            return ExcelError.ExcelErrorValue;
        }
    }
}