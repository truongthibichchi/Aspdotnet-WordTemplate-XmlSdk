using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordTemplate.Helpers
{
    class Helper_WordLineChart : Helper_WordBarChart
    {
        public Helper_WordLineChart(SdtElement content_control) : base(content_control)
        {
        }

        /// <summary>
        /// Modify/Add series into chart XML
        /// </summary>
        /// <param name="column_index">Corresponds to the column index that needs to be modified in chart spreadsheet (Ex: A, B, C, ...)</param>
        /// <param name="row_index">Corresponds to the column index that needs to be modified in excel </param>
        /// <param name="new_value">Corresponds to the new value we need to insert to the cell </param>
        protected override void ModifyChartXML_Series(string column_index, uint row_index, string new_value)
        {
            LineChartSeries linechart_series = chart_part.ChartSpace.Descendants<LineChartSeries>().Where(s => string.Compare(s.InnerText, worksheet_name + "!$" + column_index + "$1", true) > 0).FirstOrDefault();
            if (linechart_series != null)    // There exists data on the series --> We only need to modify it
            {
                SeriesText st = linechart_series.Descendants<SeriesText>().FirstOrDefault();
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
        protected override void ModifyChartXML_Categories(string column_index, uint row_index, string new_value)
        {
            foreach (LineChartSeries linechart_series in chart_part.ChartSpace.Descendants<LineChartSeries>().ToList())
            {
                CategoryAxisData category_axis_data = linechart_series.Descendants<CategoryAxisData>().FirstOrDefault();
                if (category_axis_data == null)
                {
                    // If no StringReference --> Clone one from the 1st (usually we go in this when we create a new BarChartSeries)
                    BarChartSeries template_barchartseries = chart_part.ChartSpace.Descendants<BarChartSeries>().FirstOrDefault();
                    CategoryAxisData template_categoryaxisdata = template_barchartseries.Descendants<CategoryAxisData>().FirstOrDefault();
                    CategoryAxisData new_categoryaxisdata = new CategoryAxisData(template_categoryaxisdata.OuterXml);
                    linechart_series.Append(new_categoryaxisdata);
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
            LineChartSeries linechart_series = chart_part.ChartSpace.Descendants<LineChartSeries>().Where(s => string.Compare(s.InnerText, worksheet_name + "!$" + column_index + "$1", true) > 0).First();
            DocumentFormat.OpenXml.Drawing.Charts.Values v = linechart_series.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().FirstOrDefault();
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
    }
}
