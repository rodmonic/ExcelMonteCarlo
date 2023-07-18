using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;

namespace Campagna
{

    class ExcelController : IDisposable
    {
        private readonly IRibbonUI _modelingRibbon;
        private readonly Application _excel;
        private readonly Dictionary<string, Excel.Range> _dataCells = new Dictionary<string, Excel.Range>();
        private readonly ProgressBar _progressBar;
        
        public ExcelController(Application excel, IRibbonUI modelingRibbon, ProgressBar progress)
        {
            // assign class properties
            _modelingRibbon = modelingRibbon;
            _excel = excel;
            _progressBar = progress;

            // set up global variables
            Globals._isSimulating = true;
            Globals._start = DateTime.Now;

            // set up Excel for first use
            (ExcelDnaUtil.Application as Application).Interactive = false;
            _excel.ScreenUpdating = false;
            _excel.Calculate();
            _excel.Calculation = Excel.XlCalculation.xlCalculationManual;

            // show progress bar
            _progressBar.Show(new ParentWnd());
            _progressBar.Update(0, 0, "setting up", "...");
        }

        public void GetResults()
        {
            // Start with full recalcuation
            _excel.Calculate();
            try
            {
                
                // get the dependency dictionary
                _progressBar.Update(10, 0, "Getting Dependencies", "...");
                GetDependencies();

                if (_dataCells.Count == 0)
                {
                    throw new ArgumentNullException();
                }

                // populate data sheet
                _progressBar.Update(75, 0, "Populating Data", "...");
                PopulateRandomSliceDictionary();
                PopulateData();


            }
            catch (Exception ex)
            {
                if (ex is ArgumentNullException || ex is OperationCanceledException)
                {

                }
                else
                {
                    throw;
                }
            }

        }

        private void GetDependencies()
        {
            Excel.Range allCells;
            double count = 0;
            double total = _excel.Worksheets.Count;
            double percentage;

            foreach (Excel.Worksheet sht in _excel.Worksheets)
            {
                percentage = (count / total) * 100;
                count++;
                _progressBar.Update(10, (int)percentage, "Getting Dependencies", "Finding formulas in " + sht.Name);

                try
                {
                    allCells = sht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                    foreach (Excel.Range cell in allCells)
                    {
                        MatchCollection mc;

                        // if formula matches global Output RegexPattern then add dependecies to global dictionary
                        mc = Regex.Matches(cell.Formula, Globals._outputRegexPattern);
                        if (mc.Count > 0)
                        {
                            FindDependencies(cell);
                        }
                            

                        // if formula matches global input RegexPattern then add to Random number slice dictionary
                        mc = Regex.Matches(cell.Formula, Globals._inputRegexPattern);
                        if (mc.Count > 0)
                        {
                            SliceList initialList = new SliceList();
                            Globals._randomNumbers.Add(cell.Address[true, true, Excel.XlReferenceStyle.xlA1, true, false], initialList);
                        }
                    }
                }

                // catch if there are no special cells
                catch (System.Runtime.InteropServices.COMException)
                {
                    continue;
                }

            }

        }

        private void FindDependencies(Excel.Range cell)
        {
            string inAddress = cell.Address[true, true, Excel.XlReferenceStyle.xlA1, true, false];
            int sheetIdx = _excel.Sheets[cell.Parent.Name].index;
            long pCount = 0;
            long qCount = 0;

            // Show precendents and then navigate to the start cell
            _excel.Sheets[sheetIdx].Activate();
            Excel.Range returnSelection = _excel.Selection;
            cell.ShowPrecedents();
            cell.NavigateArrow(true, 1);

            // Loop through precendents until back at the beginning
            do
            {
                // Go to first precedent
                pCount++;
                cell.NavigateArrow(true, pCount);

                // if the precedent is on a different sheet
                if (_excel.ActiveSheet.Name != returnSelection.Parent.Name)
                {
                    do
                    {
                        // Loop through external precents
                        qCount++;
                        cell.NavigateArrow(true, pCount, qCount);

                        AddToDependencyDictionary(_excel.Selection);

                        // Try to move to next external link and exit loop if it doens't exist
                        try
                        {
                            cell.NavigateArrow(true, pCount, qCount + 1);
                        }
                        catch
                        {
                            break;
                        }

                    } while (true);
                    cell.NavigateArrow(true, pCount + 1);
                }
                else
                {
                    // If precedent isn't already in List then add in
                    AddToDependencyDictionary(_excel.Selection);

                    cell.NavigateArrow(true, pCount + 1);
                }
            } while (inAddress != _excel.ActiveCell.Address[true, true, Excel.XlReferenceStyle.xlA1, true, false]);

            cell.Parent.ClearArrows();
            returnSelection.Parent.Activate();
            returnSelection.Select();

        }

        private void AddToDependencyDictionary(Excel.Range cell )
        {
            string address = cell.Address[true, true, Excel.XlReferenceStyle.xlA1, true, false];
            string formula = cell.Formula;
            if (formula.StartsWith("=") || formula.StartsWith("+") || formula.StartsWith("-"))
            {
                if (_dataCells.ContainsKey(address) == false)
                    _dataCells.Add(address, cell);
            }

        }

        private void PopulateRandomSliceDictionary()
        {
            if (Globals._randomSeed == 0)
                return;

            foreach (var item in Globals._randomNumbers)
            {
                
                for (int i = 1; i <= Globals._numberIterations; i++)
                {
                    Slice dataSlice = new Slice(i, Globals._rand.NextDouble());
                    item.Value.Add(dataSlice);
                }
            }
        }

        private void PopulateData()
        {
            double percentage;
            Globals._sliceData.Clear();

            _progressBar.Update(75, 0, "Populating Data", $"Slice 0 of {Globals._numberIterations}");

            // set up dictionary
            foreach (var item in _dataCells)
            {
                SliceList initialList = new SliceList();
                Globals._sliceData.Add(item.Key, initialList);
            }

            for (int i = 1; i <= Globals._numberIterations; i++)
            {
                Globals._currentIteration = i;
                _excel.Calculate();

                if ((i % Globals._dataSteps) == 0)
                {
                    percentage = ((double)i / Globals._numberIterations) * 100;
                    _progressBar.Update(75, (int)percentage, "Populating Data", $"Slice {i} of {Globals._numberIterations}");
                }
                foreach (var item in _dataCells)
                {
                    try
                    {
                        Slice dataSlice = new Slice(i, (double)(item.Value.Value));
                        Globals._sliceData[item.Key].Add(dataSlice);
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        
                    }
                    
                }
                
            }

        }

        public void Dispose()
        {
            // change globals
            Globals._isSimulating = false;
            Globals._randomNumbers = new SliceDictionary();

            // reset excel functionality
            _excel.Calculate();
            _excel.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            _excel.ScreenUpdating = true;
            (ExcelDnaUtil.Application as Application).Interactive = true;

            // hide progreess bar
            _progressBar.Hide();

        }
    }
}
