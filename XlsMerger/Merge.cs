using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows;
using OfficeOpenXml;

namespace XlsMerger
{
    class Merge: INotifyPropertyChanged
    {
        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected void NotifyPropertyChange(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

        private int percent = 0;
        public int Percent
        {
            get { return percent; }
            set
            {
                percent = value;
                NotifyPropertyChange("Percent");
            }
        }

        static bool Check()
        {
            return true;
        }

        private bool Check_Filled(ExcelRange block)
        {
            IEnumerator enumerator = block.GetEnumerator();
            while(enumerator.MoveNext())
            {
                object item = enumerator.Current;
                if ((item as ExcelRangeBase).Value != null) return true;
            }
            return false;
        }

        private bool Convert(string filePath)
        {
            GemBox.Spreadsheet.SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            GemBox.Spreadsheet.ExcelFile ef = GemBox.Spreadsheet.ExcelFile.Load(filePath);
            ef.Save(filePath + 'x');
            return true;
        }

        public void Merge_xlsx(List<string> inputFiles, string outputPath,string filename, List<int> fixed_rows, List<int> main_rows, int sheet_num = 1, bool ignoreEmpty = true)
        {
            List<ExcelRange> merge_data = new List<ExcelRange>();
            Nullable<int> max_len = null; // indicate the column number
            ExcelWorksheet modelsheet = null;
            int counter = 0;
            foreach (var filePath in inputFiles)
            {
                ++counter;
                FileInfo finfo;
                // first, convert xls to xlsx if neccessary
                if (filePath.EndsWith(".xls"))
                {
                    Convert(filePath);
                    finfo = new FileInfo(filePath+'x');
                }

                else finfo = new FileInfo(filePath);
                var package = new ExcelPackage(finfo);
                
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheet_num];
                    if (modelsheet == null)
                        modelsheet = package.Workbook.Worksheets[sheet_num];
                    if (max_len == null)
                    {
                        max_len = 0;
                        // let's determine the max len first
                        foreach (var row in fixed_rows)
                        {
                            for (int i = 1; i <= 100; i++)
                            {
                                //var tmp = worksheet.Cells[row, i].Value;
                                if (worksheet.Cells[row, i].Value != null)
                                    if (max_len != null && max_len < i) max_len = i;
                            }
                        }

                    }
                    // select the data to be collected
                    foreach(var mainrow in main_rows)
                {
                    var tmp = worksheet.Cells[mainrow, 1, mainrow, (int)max_len].Value;
                    if ((!ignoreEmpty) || Check_Filled(worksheet.Cells[mainrow, 1, mainrow, (int)max_len]))
                        merge_data.Add(worksheet.Cells[mainrow, 1, mainrow, (int)max_len]);
                }
                Percent = (int)((float)counter / (inputFiles.Count) * 100);
            }
            // now, create the merged document
            
            var newFile = new FileInfo(outputPath + "\\"+filename);
            
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(outputPath + "\\" + filename);
            }
            using (var package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(modelsheet.Name, modelsheet);
                // add entries from the `main_row`
                int curr_row = main_rows.Min();
                foreach (var entry in merge_data)
                {
                    worksheet.Cells[curr_row, 1, curr_row, (int)max_len].Value = entry.Value;
                    curr_row++;
                }
                package.Save();
            }
        }
    }
}
