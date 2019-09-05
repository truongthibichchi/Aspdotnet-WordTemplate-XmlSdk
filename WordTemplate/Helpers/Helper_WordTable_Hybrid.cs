using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;

namespace WordTemplate.Helpers
{
    class Helper_WordTable_Hybrid : Helper_WordTable_Columns
    {
        public Helper_WordTable_Hybrid(SdtElement content_control) : base(content_control)
        {
        }

        /// <summary>
        /// Add data to the table
        /// </summary>
        public override void AddDataToTable()
        {
            if (table == null)
                return;

            List<TableRow> table_row_list = table.Elements<TableRow>().ToList();

            DataTable sql_result = GetDataFromDatabaseUsingSQL();

            AddDataToTable(table_row_list, sql_result);
        }

        /// <summary>
        /// Method for create/add data into a table row with properties
        /// </summary>
        /// <param name=word_rows">Destination rows that we will write and add to Word</param>
        /// <param name=data">A table contains input data taken from databse</param
        /// <returns> The TableRow instance that has applied properties and values</returns>
        protected void AddDataToTable(List<TableRow> table_row_list, DataTable data)
        {
            int column_count = 0;
            foreach (TableRow table_row in table_row_list)
            {
                int cell_count = table_row.Descendants<TableCell>().Count();
                for (int i = 0; i < cell_count; i++)
                {
                    bool is_data_exists = (column_count < data.Columns.Count && data.Rows[0][column_count].ToString() != null);
                    string append_text = (is_data_exists ? data.Rows[0][column_count].ToString() : StaticValues.no_data_found_str);

                    // Check for image urls
                    if (Helper_WordPicture.IsUrl(append_text) && Helper_WordPicture.IsImageUrl(append_text))
                    {
                        image_urls.Add(append_text);
                        column_count++;
                        continue;
                    }

                    // Get data from 1st row only
                    if (table_row.Descendants<TableCell>().ElementAt(i).InnerText.ToLower() == StaticValues.value_placeholder_name)
                    {
                        Run run = new Run(new Text(append_text));
                        ChangeCellData(table_row, i, run);
                        
                        column_count++;
                    }
                }
            }

            AddImagesInTable();
        }
    }
}

