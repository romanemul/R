using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace RRD
{
    public class ExcelTools
    {
        private Excel.Application _app = new Excel.Application();
        private Excel.Workbook _wb;
        public Excel.Worksheet _ws;
        public Excel.Range _tempRange1;
        public Excel.Range _tempRange2;
        public System.Data.DataTable _WorksheetDataTable = new System.Data.DataTable();

        public int LastUsedRow;
        public int LastUsedColumn;
        public int HWNd;

        public System.Data.DataTable dt = new System.Data.DataTable();

        public Workbook Wb { 
            get 
                => _wb; 
            set 
            {
                if (value.Worksheets[1] != null)
                    _ws = value.Worksheets[1];
                    _wb = value; 
            } 
        }

        private Excel.Application RunApp() 
        {
            _app.AskToUpdateLinks = false;
            _app.Visible = true;
            _app.DisplayAlerts = false;
            HWNd = _app.Hwnd;

            return _app;
        }

        public ExcelTools()
        {
        
        }
        public ExcelTools(string FilePath)
        {
            RunApp();
            OpenFile(FilePath);
        }

        private void UsedRange()
        {
            _tempRange1 = _ws.UsedRange;
            LastUsedColumn = _tempRange1.Columns.Count;
            LastUsedRow = _tempRange1.Rows.Count;        
        }

        public void SetWorksheet(int WorksheetNumber)
        {
            _ws = Wb.Worksheets[WorksheetNumber];
            UsedRange();
            //ExcelToDataTable();
        }

        public void SetWorksheet(string WorksheetName)
        {
            _ws = Wb.Worksheets[WorksheetName];
            UsedRange();
            //ExcelToDataTable();
        }
        public void DataTableToWorksheet(System.Data.DataTable dt, string WorksheetName = "Sheet1", int StartRow = 0, int StartColumn = 0)
        {
            Microsoft.Office.Interop.Excel.Application _app = new Excel.Application();
            _app.Visible = true;
            _app.DisplayAlerts = false;
                        
            Microsoft.Office.Interop.Excel.Workbook Wb = _app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet Ws;

            if (WorksheetName == "Sheet1") 
            {
                Ws = Wb.Worksheets[WorksheetName];
            }
            else 
            {
                Ws = Wb.Worksheets.Add(Wb.Worksheets[1]);
                Ws.Name = WorksheetName;
            }


            object[,] arr = Make2DObject(dt,true);
            Excel.Range r = MakeRange(1,1,dt,Ws,true);

            r.Value = FillArray(dt,arr, 0, 0);
        }
       

        private object[,] FillArray(System.Data.DataTable dt, object[,] arr, int StartRowSource, int StartColumnSource)
        {
            for (int j = StartColumnSource; j <= dt.Columns.Count-1; j++)
            {
                arr[0, j] = dt.Columns[j].ColumnName;
                for (int i = StartRowSource; i <= dt.Rows.Count-1; i++)
                {
                    arr[i+1, j] = dt.Rows[i][j];
                }
            }
            return arr;            
        }


        private void MakeWorksheet()
        {
        
        }


        private Excel.Range MakeRange(int StartRow,int StartColumn, System.Data.DataTable dt ,Worksheet Ws, Boolean HasColumnName)
        {
            Excel.Range r = Ws.get_Range(this.RCtoAbsolute(StartRow, StartColumn), Type.Missing);
            Excel.Range s;

            if (HasColumnName) 
            {
            
                s = Ws.Cells[dt.Rows.Count+1, dt.Columns.Count];
            }
            else 
            {
                s = Ws.Cells[dt.Rows.Count, dt.Columns.Count];
            }
            
            r = (Excel.Range)Ws.get_Range(r, s);

            return r;
        }

        public object[,] Make2DObject(System.Data.DataTable Dt, Boolean HasColumnNames)
        {
            object[,] arr = new object[0,0];

            if (HasColumnNames) 
            {                
                arr = new object[Dt.Rows.Count+1, Dt.Columns.Count];
            }
            else
            {
                arr = new object[Dt.Rows.Count, Dt.Columns.Count];
            }
            return arr;
        }

        private object[,] Make2DObject(int Rows, int Columns, Boolean HasColumnNames)
        {
            object[,] arr = new object[0, 0];

            if (HasColumnNames)
            {
                arr = new object[Rows + 1, Columns];
            }
            else
            {
                arr = new object[Rows, Columns];
            }
            return arr;            
        }



        public void DataTableToWorksheetOnPosition(System.Data.DataTable dt, string WorksheetName, int PosStartRow, int PosStartColumn)
        {
            
        }

        private System.Data.DataTable Connector(string FilePath, string Command)
        {
            System.Data.DataTable tmpDataTable = new System.Data.DataTable();

            using (ExcelConnector excelConnector = new ExcelConnector(FilePath, true))
            {
                tmpDataTable = excelConnector.Select(Command);
            }
            return tmpDataTable;
        }


        public System.Data.DataTable WorksheetToDataTable(string FilePath, string WorksheetName, Boolean HasColumnName)
        {
            
            System.Data.DataTable tmpDataTable = new System.Data.DataTable();            
            
            using (ExcelConnector excelConnector = new ExcelConnector(FilePath, HasColumnName))
            {
                tmpDataTable = excelConnector.Select("SELECT * FROM [" + WorksheetName + "$]");
            }

            return tmpDataTable;
        }


        public System.Data.DataTable WorksheetToDataTableWithCommand(string FilePath, string Command)
        {

            System.Data.DataTable tmpDataTable = new System.Data.DataTable();

            using (ExcelConnector excelConnector = new ExcelConnector(FilePath, true))
            {
                tmpDataTable = this.Connector(FilePath,Command);
            }
            return tmpDataTable;
        }

        // ACE OLEDB is used for this method so there cannot be used LastUsedRow / LastUsedColumn property.

        public System.Data.DataTable WorksheetToDataTable(string FilePath, string WorksheetName, int Row)
        {
            System.Data.DataTable tmpDataTable = new System.Data.DataTable();

            using (ExcelConnector excelConnector = new ExcelConnector(FilePath, false))
            {
                tmpDataTable = excelConnector.Select("SELECT * FROM [" + WorksheetName + "$]");

                try
                {
                    int i = 0;
                    foreach (var item in tmpDataTable.Rows[Row - 1].ItemArray)
                    {
                        if (item == null || item == DBNull.Value)
                        {
                            break;
                        }

                        tmpDataTable.Columns[i].ColumnName = item.ToString();
                        i++;
                    }
                }
                catch (System.ArgumentException exception)
                {
                    return tmpDataTable;
                }
                catch
                {
                }
            }

            System.Data.DataTable tempDT = tmpDataTable.Clone();
            tempDT.Clear();

            for (int i = Row; i < tmpDataTable.Rows.Count; i++)
            {
                DataRow dr = tempDT.Rows.Add(tmpDataTable.Rows[i].ItemArray);
            }
            return tempDT;
        }

        
        private void SearchValue() 
        {            
        
        }

        private Excel.Workbook OpenFile(string filePath)
        {
            // otevri jen pro cteni
            return _wb = _app.Workbooks.Open(filePath,null, true);
        }

        public System.Data.DataTable WorksheetPartToDataTable(string Path, string WorksheetName, int StartRow, int StartColumn, int EndRow, int EndColumn)
        {
            System.Data.DataTable TmpDataTable = new System.Data.DataTable();
            if(StartRow > EndRow || StartColumn > EndColumn) 
            {
                return TmpDataTable;
            }

            string Command = "SELECT * FROM [" + WorksheetName + "$" + RCtoAbsolute(StartRow, StartColumn) + ":" + RCtoAbsolute(EndRow, EndColumn) + "]";
            TmpDataTable = this.Connector(Path, Command);
            return TmpDataTable; 
        }

        public System.Data.DataTable WorksheetPartToDataTable(string Path, string WorksheetName, string StartRangeR1C1, string EndRangeR1C1) 
        {
            System.Data.DataTable TmpDataTable = new System.Data.DataTable();
            
            string Command = "SELECT * FROM [" + WorksheetName + "$" + StartRangeR1C1 + ":" + EndRangeR1C1 + "]";
            TmpDataTable = this.Connector(Path, Command);
            return TmpDataTable;
        }

        private string RCtoAbsolute(int Row,int Column)
        {
            return this.RCToAlphabet(Column) + Row.ToString();        
        }

        public string RCToAlphabet(int columnNumber)
        {
            //https://stackoverflow.com/questions/181596/how-to-convert-a-column-number-e-g-127-into-an-excel-column-e-g-aa

            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        
        public string SearchValue(int indexOfWorkSheet,string SearchedTerm,double MaxRow, double MaxColumn)
        {
            string tempvalue = string.Empty;
            string address = string.Empty;

            if (_wb.Worksheets[indexOfWorkSheet] is null)
            {
                return "";
            }
            else
            {
                Excel.Worksheet tmpWb = _wb.Worksheets[indexOfWorkSheet];


                for (int j = 1; j < MaxColumn; j++)
                {
                    for (int i = 1; i < MaxRow; i++)
                    {
                        Excel.Range r = tmpWb.Cells[i, j];
                        string val = Convert.ToString(r.Value);
                        try { 
                            if (val.Contains(SearchedTerm))
                            {
                                Excel.Range rr = tmpWb.Cells[(i++), j];
                                tempvalue = tmpWb.Cells[(i++), j].Value2();
                                    //tempvalue = Convert.ToString(rr.Value);
                                //return i.ToString() + ", " + j.ToString();
                                return tempvalue;
                            }
                        }
                        catch (NullReferenceException e)
                        {
                            continue;
                        }
                    }
                }
            }
            return "";
        }



    }
}
