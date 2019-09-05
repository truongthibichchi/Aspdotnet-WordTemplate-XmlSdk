using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Data;

namespace WordTemplate.Helpers
{
    class Helper_WordBarChart : Helper_WordBase
    {
        protected ChartPart chart_part;
        protected string worksheet_name = "Sheet1";
        protected bool rotated_table;

        /// <summary>
        /// Constructor that find and initialize the first chart found in the content control
        /// </summary>
        /// <param name="content_control">SdtElement instance represent the content control</param>
        public Helper_WordBarChart(SdtElement content_control)
        {
            rotated_table = StaticValues.barchart_rotatedtable;
            if (!Init(content_control))
                return;

            // Get the first ChartPart instance found in the content control 
            XElement chart_content_control = Helper_WordBase.GetContentControlByTag(Helper_WordBase.GetContentControlXMLBySdtElement(content_control), StaticValues.content_cc_name);
            string chart_id = (string)chart_content_control.Descendants(StaticValues.c + "chart").Attributes(StaticValues.r + "id").FirstOrDefault<XAttribute>().Value;
            chart_part = (ChartPart)WordTemplateManager.document.MainDocumentPart.GetPartById(chart_id);

            if (chart_part == null)
                StaticValues.logs += "[Error][WordChart]Can't initialize required data using the input content control" + Environment.NewLine;
        }

        #region Public method used to run the task from outside
        /// <summary>
        /// Update chart using the required table. Call this method after class initialization.
        /// </summary>
        public void UpdateChartFromSQL()
        {
            if (chart_part == null || sql_query == null)
                return;

            System.Data.DataTable result = GetDataFromDatabaseUsingSQL();
            DeleteAllChartPreviousData();

            // Insert data into chart - Note that change the flow may lead to unexpected error
            if (rotated_table == false)
            {
                InsertColumnNameIntoChart(result);
                InsertRowNameIntoChart(result);
                InsertDataIntoChart(result);
            }
            else
            {
                InsertColumnNameIntoChart_Rotated(result);
                InsertRowNameIntoChart_Rotated(result);
                InsertDataIntoChart_Rotated(result);
            }
            
            //ModifyChartRange(result);
        }
        #endregion

        #region Inner functions
        /// <summary>
        /// Add column name from data table into chart
        /// </summary>
        /// <param name="table">The input datatable</param>
        protected void InsertColumnNameIntoChart(System.Data.DataTable table)
        {
            if (table == null)
                return;

            for (int i = 1; i < table.Columns.Count; i++)
            {
                string col_name = table.Columns[i].ColumnName.ToString();
                UpdateChart(GetColumnIndexByNum(i), 1, col_name, true);
            }
        }

        /// <summary>
        /// Add column name from data table into chart (rotated mode)
        /// </summary>
        /// <param name="table">The input datatable</param>
        protected void InsertColumnNameIntoChart_Rotated(System.Data.DataTable table)
        {
            if (table == null)
                return;

            for (int i = 1; i < table.Columns.Count; i++)
            {
                string col_name = table.Columns[i].ColumnName.ToString();
                UpdateChart("A", (uint)GetRowIndexByNum(i - 1), col_name, true);
            }
        }

        /// <summary>
        /// Add row name from data table into chart
        /// </summary>
        /// <param name="table">The input datatable</param>
        protected void InsertRowNameIntoChart(System.Data.DataTable table)
        {
            if (table == null || table.Columns.Count == 0)
                return;

            for (int i = 0; i < table.Rows.Count; i++)
                UpdateChart("A", (uint)(i + 2), table.Rows[i][0].ToString(), true);
        }

        /// <summary>
        /// Add row name from data table into chart (rotated)
        /// </summary>
        /// <param name="table">The input datatable</param>
        protected void InsertRowNameIntoChart_Rotated(System.Data.DataTable table)
        {
            if (table == null || table.Columns.Count == 0)
                return;
            for (int i = 0; i < table.Rows.Count; i++)
                UpdateChart(GetColumnIndexByNum(i + 1), 1, table.Rows[i][0].ToString(), true);
        }

        /// <summary>
        /// Insert data into table
        /// </summary>
        /// <param name="table">The input datatable</param>
        protected void InsertDataIntoChart(System.Data.DataTable table)
        {
            if (table == null)
                return;
            for (int i = 0; i < table.Rows.Count; i++)
                for (int k = 1; k < table.Columns.Count; k++)
                {
                    string data = table.Rows[i][k].ToString();
                    int num = Convert.ToInt32(data);
                    UpdateChart(GetColumnIndexByNum(k), (uint)GetRowIndexByNum(i), num.ToString(), false);
                }
        }

        /// <summary>
        /// Insert data into table
        /// </summary>
        /// <param name="table">The input datatable</param>
        protected void InsertDataIntoChart_Rotated(System.Data.DataTable table)
        {
            if (table == null)
                return;
            for (int i = 0; i < table.Rows.Count; i++)
                for (int k = 1; k < table.Columns.Count; k++)
                {
                    string data = table.Rows[i][k].ToString();
                    int num = Convert.ToInt32(data);
                    UpdateChart(GetColumnIndexByNum(i + 1), (uint)GetRowIndexByNum(k - 1), num.ToString(), false);
                }
        }

        /// <summary>
        /// Modify the worksheet (embedded into chart) range values
        /// </summary>
        /// <param name="result">The DataTable instance represent a table</param>
        protected void ModifyChartRange(System.Data.DataTable result)
        {
            ChartSpace chartspace = chart_part.ChartSpace;
            for (int i = 1; i < result.Columns.Count; i++)
            {
                BarChartSeries barchart_series = chart_part.ChartSpace.Descendants<BarChartSeries>().Where(s => string.Compare(s.InnerText, worksheet_name + "!$" + GetColumnIndexByNum(i) + "$1", true) > 0).First();
                DocumentFormat.OpenXml.Drawing.Charts.Values val = barchart_series.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().FirstOrDefault();
                NumberReference nr = val.Descendants<NumberReference>().First();
                DocumentFormat.OpenXml.Drawing.Charts.Formula f = nr.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>().First();

                f.Text = worksheet_name + "!$" + GetColumnIndexByNum(i) + "$2:$" + GetColumnIndexByNum(i) + "$" + GetRowIndexByNum(result.Rows.Count - 1);
            }
        }

        /// <summary>
        /// Update a value in a worksheet embedded into the word chart
        /// </summary>
        /// <param name="column_index">Corresponds to the column index that needs to be modified in chart spreadsheet (Ex: A, B, C, ...)</param>
        /// <param name="row_index">Corresponds to the column index that needs to be modified in excel </param>
        /// <param name="new_value">Corresponds to the new value we need to insert to the cell </param>
        /// <param name="is_axis_value">Is the updated cell the axis? </param>
        protected virtual void UpdateChart(string column_index, uint row_index, string new_value, bool is_axis_value)
        {

            bool update_worksheet_complete = UpdateChartWorksheet(column_index, row_index, new_value, is_axis_value);
            if (update_worksheet_complete)
            {
                // The data in worksheet is updated
                // But we also need to make change to XML structure in order to update the word chart
                if (row_index == 1)     // Modify series
                    ModifyChartXML_Series(column_index, row_index, new_value);
                else if (column_index == "A")
                    ModifyChartXML_Categories(column_index, row_index, new_value);
                else
                    ModifyChartXML_Data(column_index, row_index, new_value);
            }
        }

        /// <summary>
        /// Modify the value in the worksheet embedded in chart
        /// </summary>
        /// <param name="column_index">Corresponds to the column index that needs to be modified in chart spreadsheet (Ex: A, B, C, ...)</param>
        /// <param name="row_index">Corresponds to the column index that needs to be modified in excel </param>
        /// <param name="new_value">Corresponds to the new value we need to insert to the cell </param>
        /// <param name="is_axis_value">Is the updated cell the axis? </param>
        /// <returns>Is the task completed successfully?</returns>
        protected virtual bool UpdateChartWorksheet(string column_index, uint row_index, string new_value, bool is_axis_value)
        {
            if (chart_part == null)
                return false;

            Stream stream = chart_part.EmbeddedPackagePart.GetStream();
            // Open the internal spreadsheet doc for the chart
            using (SpreadsheetDocument wordSSDoc = SpreadsheetDocument.Open(stream, true))
            {
                // Navigate to the sheet where the chart data is located
                WorkbookPart workBookPart = wordSSDoc.WorkbookPart;
                Sheet theSheet = workBookPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == worksheet_name).FirstOrDefault();
                if (theSheet != null)
                {
                    // Update data in worksheet
                    Worksheet ws = ((WorksheetPart)workBookPart.GetPartById(theSheet.Id)).Worksheet;

                    // Get the cell which needs to be updated
                    Cell theCell = GetCellInWorksheet(column_index, row_index, ws);

                    // Update the cell value
                    theCell.CellValue = new CellValue(new_value);
                    if (is_axis_value)
                    {
                        // We are updating the Series text
                        theCell.DataType = new EnumValue<CellValues>(CellValues.String);
                    }
                    else
                    {
                        // We are updating a numeric chart value
                        theCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    }

                    ws.Save();
                    return true;
                }
                return false;
            }
        }

        /// <summary>
        /// Modify/Add series into chart XML
        /// </summary>
        /// <param name="column_index">Corresponds to the column index that needs to be modified in chart spreadsheet (Ex: A, B, C, ...)</param>
        /// <param name="row_index">Corresponds to the column index that needs to be modified in excel </param>
        /// <param name="new_value">Corresponds to the new value we need to insert to the cell </param>
        protected virtual void ModifyChartXML_Series(string column_index, uint row_index, string new_value)
        {
            BarChartSeries barchart_series = chart_part.ChartSpace.Descendants<BarChartSeries>().Where(s => string.Compare(s.InnerText, worksheet_name + "!$" + column_index + "$1", true) > 0).FirstOrDefault();
            if (barchart_series != null)    // There exists data on the series --> We only need to modify it
            {
                SeriesText st = barchart_series.Descendants<SeriesText>().FirstOrDefault();
                StringReference sr = st.Descendants<StringReference>().First();
                StringCache sc = sr.Descendants<StringCache>().First();
                StringPoint sp = sc.Descendants<StringPoint>().First();
                NumericValue nv = sp.Descendants<NumericValue>().First();
                nv.Text = new_value;
            }
            else    // No such series exists --> Consider create a new series
            {
                /*
                // Find location in XML to append the BarChartSeries
                Chart chart = chart_part.ChartSpace.Descendants<Chart>().FirstOrDefault();
                PlotArea plot = chart.PlotArea;
                
                // Create new BarChartSeries
                barchart_series = new BarChartSeries();
                uint index = (uint)plot.Descendants<BarChartSeries>().ToList().Count;

                barchart_series.Append(new Index() { Val = index });
                barchart_series.Append(new Order() { Val = index });

                SeriesText seriesText = new SeriesText();
                seriesText.Append(new NumericValue() { Text = new_value });

                barchart_series.Append(seriesText);
                

                // Append data
                Bar3DChart bar_3dchart = plot.Descendants<Bar3DChart>().FirstOrDefault();
                if (bar_3dchart != null)       // Chart is 3D
                    bar_3dchart.Append(barchart_series);
                else    // Chart is not 3d
                {
                    BarChart barchart = plot.Descendants<BarChart>().FirstOrDefault();
                    barchart.Append(barchart_series);
                }
                
                // Append other settings
                BarChartSeries barchart_series_template = chart_part.ChartSpace.Descendants<BarChartSeries>().LastOrDefault();

                CategoryAxisData cateAxisData = new CategoryAxisData();
                StringReference string_ref = new StringReference();
                string_ref.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = barchart_series.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>().FirstOrDefault().Text});
                StringCache string_cache = new StringCache();
                string_cache.Append(new PointCount() { Val = count });
                */
            }
        }

        /// <summary>
        /// Modify/Add categories into chart XML
        /// </summary>
        /// <param name="column_index">Corresponds to the column index that needs to be modified in chart spreadsheet (Ex: A, B, C, ...)</param>
        /// <param name="row_index">Corresponds to the column index that needs to be modified in excel </param>
        /// <param name="new_value">Corresponds to the new value we need to insert to the cell </param>
        protected virtual void ModifyChartXML_Categories(string column_index, uint row_index, string new_value)
        {
            foreach (BarChartSeries barchart_series in chart_part.ChartSpace.Descendants<BarChartSeries>().ToList())
            {
                CategoryAxisData category_axis_data = barchart_series.Descendants<CategoryAxisData>().FirstOrDefault();
                if (category_axis_data == null)
                {
                    // If no StringReference --> Clone one from the 1st (usually we go in this when we create a new BarChartSeries)
                    BarChartSeries template_barchartseries = chart_part.ChartSpace.Descendants<BarChartSeries>().FirstOrDefault();
                    CategoryAxisData template_categoryaxisdata = template_barchartseries.Descendants<CategoryAxisData>().FirstOrDefault();
                    CategoryAxisData new_categoryaxisdata = new CategoryAxisData(template_categoryaxisdata.OuterXml);
                    barchart_series.Append(new_categoryaxisdata);
                }
                else
                {
                    StringReference sr = category_axis_data.Descendants<StringReference>().FirstOrDefault();
                    // If there is a StringReference --> Update its values
                    StringCache sc = sr.Descendants<StringCache>().First();
                    try
                    {
                        StringPoint sp = sc.Descendants<StringPoint>().ElementAt((int)row_index - 2);
                        NumericValue nv = sp.Descendants<NumericValue>().First();
                        nv.Text = new_value;
                    }
                    catch (Exception)
                    {
                        // Create new data and append to previous XML
                        sc.PointCount.Val = sc.PointCount.Val + 1;
                        NumericValue nv = new NumericValue(new_value);
                        StringPoint sp = new StringPoint(nv);
                        sp.Index = (uint)sc.Descendants<StringPoint>().ToList().Count;
                        sc.Append(sp);

                        // Change fomula range
                        DocumentFormat.OpenXml.Drawing.Charts.Formula f = sr.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>().FirstOrDefault();
                        f.Text = worksheet_name + "!$A$2:$A$" + GetRowIndexByNum((int)row_index - 2).ToString();
                    }
                }
                
                
            } 
        }

        /// <summary>
        /// Modify/Add data into chart XML
        /// </summary>
        /// <param name="column_index">Corresponds to the column index that needs to be modified in chart spreadsheet (Ex: A, B, C, ...)</param>
        /// <param name="row_index">Corresponds to the column index that needs to be modified in excel </param>
        /// <param name="new_value">Corresponds to the new value we need to insert to the cell </param>
        protected virtual void ModifyChartXML_Data(string column_index, uint row_index, string new_value)
        {
            BarChartSeries barchart_series = chart_part.ChartSpace.Descendants<BarChartSeries>().Where(s => string.Compare(s.InnerText, worksheet_name + "!$" + column_index + "$1", true) > 0).First();
            DocumentFormat.OpenXml.Drawing.Charts.Values v = barchart_series.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().FirstOrDefault();
            NumberReference nr = v.Descendants<NumberReference>().First();
            NumberingCache nc = nr.Descendants<NumberingCache>().First();
            
            try
            {
                NumericPoint np = nc.Descendants<NumericPoint>().ElementAt((int)row_index - 2);
                NumericValue nv = np.Descendants<NumericValue>().First();
                nv.Text = new_value;
            }
            catch (Exception)
            {
                // Create new data and append to previous XML
                nc.PointCount.Val = nc.PointCount.Val + 1;
                NumericValue nv = new NumericValue(new_value);
                NumericPoint np = new NumericPoint(nv);
                np.Index = (uint)nc.Descendants<NumericPoint>().ToList().Count;
                nc.Append(np);

                // Change fomula range
                DocumentFormat.OpenXml.Drawing.Charts.Formula f = nr.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>().FirstOrDefault();
                f.Text = worksheet_name + "!$" + column_index + "$2:$" + column_index + "$" + GetRowIndexByNum((int)row_index - 2).ToString();
            }
           
        }

        protected void DeleteAllChartPreviousData()
        {
            //TODO: Implement this function later
        }

        #endregion

        #region Helper functions
        /// <summary>
        /// Get a cell (and insert if doesn't exists) from a chart worksheet
        /// </summary>
        /// <param name="column_name">Corresponds to the column index(Ex: A, B, C, ...)</param>
        /// <param name="row_index">Corresponds to the row index</param>
        /// <param name="worksheet">The worksheet where we need to find/insert cell </param>
        /// <returns> The required cell</returns>
        protected Cell GetCellInWorksheet(string column_name, uint row_index, Worksheet worksheet)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = column_name + row_index;
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == row_index).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == row_index).First();
            }
            else
            {
                row = new Row() { RowIndex = row_index };
                sheetData.Append(row);
            }

            if (row.Elements<Cell>().Where(c => c.CellReference.Value == column_name + row_index).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                return newCell;
            }
        }

        /// <summary>
        /// Convert column index used in data table to column name used in Excel
        /// </summary>
        /// <param name="idx">DataTable column index</param>
        /// <returns></returns>
        protected string GetColumnIndexByNum(int idx)
        {
            int quotient = (idx) / 26;

            if (quotient > 0)
                return GetColumnIndexByNum(quotient - 1) + (char)((idx % 26) + 'A');
            else
                return "" + (char)((idx % 26) + 'A');
        }

        /// <summary>
        /// Convert row index used in data table to column name used in Excel
        /// </summary>
        /// <param name="idx">DataTable row index</param>
        /// <returns></returns>
        protected int GetRowIndexByNum(int idx)
        {
            return idx + 2;
        }
        #endregion

    }
}

