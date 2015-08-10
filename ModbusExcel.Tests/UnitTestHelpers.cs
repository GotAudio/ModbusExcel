using System;
using System.Collections.Generic;
using System.Globalization;
using System.Collections;

namespace ModbusExcel.Tests
{
    class UnitTestHelpers
    {
        /// <summary>
        /// Helper class to compare 2 values are within a certain range.
        /// </summary>
        public class DoubleComparer : IComparer<Double>
        {
            public Double MarginOfError { get; private set; }

            public DoubleComparer(double marginOfError)
            {
                MarginOfError = marginOfError;
            }

            public int Compare(double x, double y)  // x = expected, y = actual
            {
                var margin = x - y;
                if (margin <= MarginOfError)
                    return 0;
                return new Comparer(CultureInfo.CurrentUICulture).Compare(x, y);
            }
        }

    }
}
