using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml.Linq;

namespace WordTemplate.Helpers
{
    class Helper_WordTable_Columns : Helper_WordBase
    {
        protected Table table;
        protected List<string> image_urls;   // Store url to show pictures (if there are any)

        /// <summary>
        /// Constructor that find and initialize the first table found in the content control
        /// </summary>
        /// <param name="content_control)">SdtElement instance represent the content control</param>
        public Helper_WordTable_Columns(SdtElement content_control)
        {
            image_urls = new List<string>();
            if (!Init(content_control))
                return;

            //Get the first Table instance found in the content control 
            table = content_control.Descendants<Table>().First();

            if (table == null)
                StaticValues.logs += "[Error]Can't get table in the content control" + Environment.NewLine;
        }

        /// <summary>
        /// Get properties of a cell in an table row of an table
        /// </summary>
        /// <param name="row_copy">A TableRow instance represent a row in a table</param>
        /// <param name="cell_index">The index of the cell we want to retrieve properties</param>
        /// <returns> Properties of a cell in a row of a table in RunProperties instance</returns>
        protected RunProperties GetRunPropertyFromTableCell(TableRow row_copy, int cell_index)
        {
            RunProperties runProperties = new RunProperties();
            string fontname;
            string fontSize;
            try
            {
                // Get font from table
                fontname =
                    row_copy.Descendants<TableCell>()
                       .ElementAt(cell_index)
                       .GetFirstChild<Paragraph>()
                       .GetFirstChild<ParagraphProperties>()
                       .GetFirstChild<ParagraphMarkRunProperties>()
                       .GetFirstChild<RunFonts>()
                       .Ascii;
            }
            catch
            {
                // Apply default font
                fontname = StaticValues.default_font;
            }
            try
            {
                fontSize =
                       row_copy.Descendants<TableCell>()
                          .ElementAt(cell_index)
                          .GetFirstChild<Paragraph>()
                          .GetFirstChild<ParagraphProperties>()
                          .GetFirstChild<ParagraphMarkRunProperties>()
                          .GetFirstChild<FontSize>()
                          .Val;
            }
            catch
            {
                // Apply default font size
                fontSize = StaticValues.default_font_size;
            }
            runProperties.AppendChild(new RunFonts() { Ascii = fontname });
            runProperties.AppendChild(new FontSize() { Val = fontSize });

            //TODO: Get more properties if possible ???
            return runProperties;
        }

        /// <summary>
        /// Add data to the table
        /// </summary>
        public virtual void AddDataToTable()
        {
            if (table == null)
                return;

            TableRow table_row = table.Elements<TableRow>().Last();

            DataTable sql_result = GetDataFromDatabaseUsingSQL();

            foreach (DataRow row in sql_result.Rows)
            {
                TableRow row_copy = (TableRow)table_row.CloneNode(true);
                row_copy = CreateRow(row_copy, row);
                table.AppendChild(row_copy);
            }
            table.RemoveChild(table_row);
            AddImagesInTable();
        }

        /// <summary>
        /// Method for create/add data into a table row with properties
        /// </summary>
        /// <param name=word_row">Destination row that we will write and add to Word</param>
        /// <param name=data_row">A row contains input data taken from databse</param
        /// <returns> The TableRow instance that has applied properties and values</returns>
        protected TableRow CreateRow(TableRow word_row, DataRow data_row)
        {
            int cell_count = word_row.Descendants<TableCell>().Count();
            for (int i = 0; i < cell_count; i++)
            {
                bool is_data_exists = (i < data_row.Table.Columns.Count && data_row[i].ToString() != null);
                string append_text = (is_data_exists ? data_row[i].ToString() : StaticValues.no_data_found_str);

                /*
                // Check for image urls
                if (Helper_WordPicture.IsUrl(append_text) && Helper_WordPicture.IsImageUrl(append_text))
                {
                    image_urls.Add(append_text);
                    continue;
                } 
                */

                Run run = new Run(new Text(append_text));
                RunProperties run_properties = GetRunPropertyFromTableCell(word_row, i);
                run.PrependChild<RunProperties>(run_properties);

                ChangeCellData(word_row, i, run);
            }
            return word_row;
        }

        /// <summary>
        /// Update data in a cell (TableCell instance) in a row (TableRow instance)
        /// </summary>
        /// <param name="word_row">TableRow instance represent a row</param>
        /// <param name="cell_index">Index of the cell we want to change data</param>
        /// <param name="run">A run instance that represent properties of data in that cell</param>
        protected virtual void ChangeCellData(TableRow word_row, int cell_index, Run run)
        {
            // Removes that text of the copied cell and add new text
            TableCell curr_cell = word_row.Descendants<TableCell>().ElementAt(cell_index);

            string fill_color = "abcdef";
            string text = run.Descendants<Text>().FirstOrDefault().Text.ToLower();

            // Text need to apply fill color to table cell ?
            if (StaticValues.table_value_with_color_properties.ContainsKey(text))
            {
                TableCellProperties properties = curr_cell.TableCellProperties;
                fill_color = StaticValues.table_value_with_color_properties[text];
                if (properties == null)
                {
                    properties = new TableCellProperties();
                    Shading shading = new Shading()
                    {
                        Color = "auto",
                        Fill = fill_color,
                        Val = ShadingPatternValues.Clear
                    };
                    properties.Append(shading);
                    curr_cell.Append(properties);
                }
                else
                {
                    Shading shading = properties.Shading;
                    if (shading == null)
                    {
                        shading = new Shading()
                        {
                            Color = "auto",
                            Fill = fill_color,
                            Val = ShadingPatternValues.Clear
                        };
                        properties.Append(shading);
                    }
                    else
                    {
                        shading.Color = "auto";
                        shading.Fill = fill_color;
                        shading.Val = ShadingPatternValues.Clear;
                    }
                }
            }

            curr_cell.RemoveAllChildren<Paragraph>(); 
            curr_cell.Append(new Paragraph(run));
        }

        /// <summary>
        /// Add image to table (if some links is founds)
        /// </summary>
        protected void AddImagesInTable()
        {
            List<SdtElement> image_ccs = content_control.Descendants<SdtElement>().Where(s => s.Descendants<SdtAlias>().FirstOrDefault().Val.ToString().ToLower() == StaticValues.image_cc_name).ToList();
            if (image_ccs.Count == 0 || image_urls.Count == 0)
                return;

            int count = 0;
            foreach (SdtElement image_cc in image_ccs)
            {
                if (count < image_urls.Count)
                {
                    Helper_WordPicture pic_helper = new Helper_WordPicture(image_cc, image_urls[count]);
                    pic_helper.AddPictureFromUri();
                }
            }

        }
    }    
}
