using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace WordTemplate.Helpers
{
    class Helper_WordPieChart : Helper_WordBarChart
    {
        public Helper_WordPieChart(SdtElement content_control) : base(content_control)
        {
            //worksheet_name = "Blad1";
        }

        #region Inner functions
        /// <summary>
        /// Modify/Add series into chart XML
        /// </summary>
        /// <param name="column_index">Corresponds to the column index that needs to be modified in chart spreadsheet (Ex: A, B, C, ...)</param>
        /// <param name="row_index">Corresponds to the column index that needs to be modified in excel </param>
        /// <param name="new_value">Corresponds to the new value we need to insert to the cell </param>
        protected override void ModifyChartXML_Series(string column_index, uint row_index, string new_value)
        {
            PieChartSeries piechart_series = chart_part.ChartSpace.Descendants<PieChartSeries>().Where(s => string.Compare(s.InnerText, worksheet_name + "!$" + column_index + "$1", true) > 0).FirstOrDefault();
            if (piechart_series != null)    // There exists data on the series --> We only need to modify it
            {
                SeriesText st = piechart_series.Descendants<SeriesText>().FirstOrDefault();
                StringReference sr = st.Descendants<StringReference>().First();
                StringCache sc = sr.Descendants<StringCache>().First();
                StringPoint sp = sc.Descendants<StringPoint>().First();
                NumericValue nv = sp.Descendants<NumericValue>().First();
                nv.Text = new_value;
            }
            else    // No such series exists --> Consider create a new series
            {
                /*
                // Find location in XML to append the PieChartSeries
                Chart chart = chart_part.ChartSpace.Descendants<Chart>().FirstOrDefault();
                PlotArea plot = chart.PlotArea;
                
                // Create new PieChartSeries
                barchart_series = new PieChartSeries();
                uint index = (uint)plot.Descendants<PieChartSeries>().ToList().Count;

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
                PieChartSeries barchart_series_template = chart_part.ChartSpace.Descendants<PieChartSeries>().LastOrDefault();

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
        protected override void ModifyChartXML_Categories(string column_index, uint row_index, string new_value)
        {
            foreach (PieChartSeries piechart_series in chart_part.ChartSpace.Descendants<PieChartSeries>().ToList())
            {
                CategoryAxisData category_axis_data = piechart_series.Descendants<CategoryAxisData>().FirstOrDefault();
                if (category_axis_data == null)
                {
                    // If no StringReference --> Clone one from the 1st (usually we go in this when we create a new PieChartSeries)
                    PieChartSeries template_barchartseries = chart_part.ChartSpace.Descendants<PieChartSeries>().FirstOrDefault();
                    CategoryAxisData template_categoryaxisdata = template_barchartseries.Descendants<CategoryAxisData>().FirstOrDefault();
                    CategoryAxisData new_categoryaxisdata = new CategoryAxisData(template_categoryaxisdata.OuterXml);
                    piechart_series.Append(new_categoryaxisdata);
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
        protected override void ModifyChartXML_Data(string column_index, uint row_index, string new_value)
        {
            PieChartSeries piechart_Series = chart_part.ChartSpace.Descendants<PieChartSeries>().Where(s => string.Compare(s.InnerText, worksheet_name + "!$" + column_index + "$1", true) > 0).First();
            DocumentFormat.OpenXml.Drawing.Charts.Values v = piechart_Series.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().FirstOrDefault();
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
        #endregion
    }
}
