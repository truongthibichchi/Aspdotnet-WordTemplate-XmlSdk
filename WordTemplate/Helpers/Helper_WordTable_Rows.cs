using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;

namespace WordTemplate.Helpers
{
    class Helper_WordTable_Rows : Helper_WordTable_Columns
    {
        public Helper_WordTable_Rows(SdtElement content_control) : base(content_control)
        {
        }

        /// <summary>
        /// Add data to the table
        /// </summary>
        /// <param name="???">???</param>
        public override void AddDataToTable()
        {
            if (table == null)
                return;

            List<TableRow> table_row_list = table.Elements<TableRow>().ToList();

            DataTable sql_result = GetDataFromDatabaseUsingSQL();

            int count = 0;
            foreach (TableRow table_row in table_row_list)
            {
                if (count < sql_result.Columns.Count)
                {
                    // Get all data of a column
                    List<string> data = new List<string>(sql_result.Rows.Count);
                    foreach (DataRow row in sql_result.Rows)
                        data.Add((row[count].ToString()));
                    // Add
                    CreateRow(table_row, data);
                }
                else
                    // Add dummy values
                    CreateRow(table_row, null);

                count++;
            }

            AddImagesInTable();
        }

        /// <summary>
        /// Method for create/add data into a table row with properties
        /// </summary>
        /// <param name=word_row">Destination row that we will write and add to Word</param>
        /// <param name=data">A list of string contains data we want to add into this row</param
        /// <returns> The TableRow instance that has applied properties and values</returns>
        protected TableRow CreateRow(TableRow word_row, List<string>data)
        {
            //TODO: Implement code to allow user to add more columns
            int cell_count = word_row.Descendants<TableCell>().Count();
            for (int i = 1; i < cell_count; i++)
            {
                bool is_data_exists = (data != null && i - 1 < data.Count);
                string append_text = (is_data_exists ? data[i - 1] : StaticValues.no_data_found_str);
                
                /*
                // Check for image urls
                if (Helper_WordPicture.IsUrl(append_text) && Helper_WordPicture.IsImageUrl(append_text))
                {
                    image_urls.Add(append_text);
                    continue;
                }
                */
                
                Run run = new Run(new Text(append_text));

                ChangeCellData(word_row, i, run);
            }
            return word_row;
        }
    }
}
