using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Campagna
{

    internal class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // Versions before v1.1.0 required only a call to Register() in the AutoOpen().
            // The name was changed (and made obsolete) to highlight the pair of function calls now required.
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }

    internal class Slice : IComparable<Slice>

    {
        public int Index { get; set; }
        public double Data { get; set; }

        public Slice(int index, double data)
        {
            this.Index = index;
            this.Data = data;
        }

        // Default comparer for Slice type.
        public int CompareTo(Slice compareSlice)
        {
            // A null value means that this object is greater.
            if (compareSlice == null)
                return 1;
            else
                return this.Data.CompareTo(compareSlice.Data);
        }

        public override int GetHashCode()
        {
            return Index;
        }

    }

    internal class SliceList: List<Slice>
    {

    }

    internal class SliceDictionary: Dictionary<string, SliceList>
    {
        public SliceList GetSliceList(string cellName)
        {

            if (this.TryGetValue(cellName, out SliceList sliceList))
            {
                return sliceList;
            }
            else
            {
                throw new ArgumentNullException();
            }

        }
    }

    internal class ParentWnd : IWin32Window
    {
        public IntPtr Handle
        {
            get
            {
                // get current excel window through Excel DNA
                return (IntPtr)(ExcelDnaUtil.Application as Application).Hwnd;
            }
        }
    }

    internal class Globals
    {

        // These dictate which formulas are identified as input or output functions
        public static string _outputRegexPattern = "CampagnaOutput";
        public static string _inputRegexPattern = "CampagnaInput";

        // These deal with simulations
        public static readonly int _dataSteps = 100;
        public static int _numberIterations;
        public static int _currentIteration;
        public static bool _isSimulating = false;
        public static DateTime _start;

        // These deal with the random number generator
        public static int _randomSeed;
        public static Random _rand = new Random();

        // set up Dictionary for storage of the data
        public static SliceDictionary _sliceData = new SliceDictionary();
        public static SliceDictionary _randomNumbers = new SliceDictionary();

    }
}
