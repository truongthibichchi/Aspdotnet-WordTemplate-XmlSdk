using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WordTemplate.Helpers
{
    // Base class for helper class which working with Word.
    // Other helper classes will inherit from this class
    class Helper_WordBase
    {
        // The content control specify the area where we going to work with
        protected SdtElement content_control;

        // The required query to get the required data from the database
        protected string sql_query;

        /// <summary>
        /// The function used to initialize the required data
        /// </summary>
        /// <param name="content_control">SdtElement instance for the content control</param>
        public virtual bool Init(SdtElement content_control)
        {
            this.content_control = content_control;
            if (this.content_control == null)
            {
                StaticValues.logs += "[Error]Can't get the content control" + Environment.NewLine;
                return false;
            }

            // Get the SQL statement in the nested content control 
            XElement sql_content_control = Helper_WordBase.GetContentControlByTag(Helper_WordBase.GetContentControlXMLBySdtElement(content_control), StaticValues.sql_cc_name);
            if (sql_content_control != null)
                sql_query = sql_content_control.Value.ToString();

            if (sql_query != null && !StaticValues.use_default_tablename_in_sql)
                sql_query = sql_query.Replace(StaticValues.dummy_table_name, StaticValues.project_table_name);

            return true;
        }

        internal static void DeleteAllSQLQueryInWord(WordprocessingDocument document)
        {
            List<SdtElement> sql_content_control_list = document.MainDocumentPart.Document.Body.Descendants<SdtElement>().Where(s => s.Descendants<SdtAlias>().FirstOrDefault().Val.ToString().ToLower() == StaticValues.sql_cc_name || s.Descendants<SdtAlias>().FirstOrDefault().Val.ToString().ToLower() == StaticValues.repeatsql_cc_name).ToList();

            foreach (SdtElement content_control in sql_content_control_list)
            {
                // Recheck to make sure
                if (content_control.InnerText.ToLower().Contains("select"))
                    content_control.Remove();
            }
        }

        #region Content control helper functions
        /// <summary>
        /// Get the XML that present the content control with alias named tag
        /// </summary>
        /// <param name="ancestor">The XML represents the document we want to search</param>
        /// <param name="tag">Title/Alias of the content control we need to retrieve</param>
        /// <returns> XElement instance for the XML represents the content control</returns>
        public static XElement GetContentControlByTag(XContainer ancestor, string tag)
        {
            XElement result;
            try
            {
                result = (from e in ancestor.Descendants(StaticValues.w + "sdt")
                          where e.Elements(StaticValues.w + "sdtPr").Elements(StaticValues.w + "alias").Attributes(StaticValues.w + "val").FirstOrDefault<XAttribute>().Value.ToLower() == tag.ToLower()
                          select e).FirstOrDefault<XElement>();
            }
            catch (Exception ex)
            {
                StaticValues.logs += ex.Message + Environment.NewLine;
                result = null;
            }
            return result;
        }

        /// <summary>
        /// Retrieve the XML for the content control
        /// </summary>
        /// <param name="content_control">The SdtElement instance represent the content control</param>
        /// <returns> The XML represenyt the 1st content control found</returns>
        public static XElement GetContentControlXMLBySdtElement(SdtElement content_control)
        {
            return XElement.Parse(content_control.OuterXml);
        }

        /// <summary>
        /// Get the data in a content control
        /// </summary>
        /// <param name="content_control">The SdtElement instance represent the content control</param>
        /// <returns> The content of the content control in string</returns>
        public static string GetContentControlContents(SdtElement content_control)
        {
            string result = "";
            try
            {
                foreach (Text t in content_control.Descendants<Text>())
                    result += t.Text;
            }
            catch (Exception ex)
            {
                StaticValues.logs += ex.Message + Environment.NewLine;
                result = null;
            }
            return result;
        }

        /// <summary>
        /// Add text to content control
        /// </summary>
        /// <param name="content_control">The content control where we want to input text</param>
        /// <param name="append_text">The text we want to append</param>
        public static void ReplaceContentsInContentControl(SdtElement content_control, string append_text)
        {
            content_control.Descendants<Text>().First().Text = append_text;
            content_control.Descendants<Text>().Skip(1).ToList().ForEach(t => t.Remove());
        }
        #endregion

        /// <summary>
        /// Add WHERE ... to sql query to specify additional condition. 
        /// This method is used to process data in "Repeat" content control
        /// For now, we only support '=' operator
        /// </summary>
        /// <param name="table_row">A row define fields name (column name) and values (cell) for WHERE clause</param>
        public void AddConditionToSQLQuery(DataRow table_row)
        {
            this.sql_query = AddCondtionalToSQLQuery(this.sql_query, table_row);
        }

        public static string AddCondtionalToSQLQuery(string query, DataRow table_row)
        {
            if (table_row == null)
                return query;

            string where_clause = "";

            foreach (DataColumn column in table_row.Table.Columns)
            {
                // Get only the 1st value
                string column_name = column.ColumnName;
                int counter = 0;
                if (!table_row.IsNull(column))
                {
                    int temp;
                    if (query.ToLower().Contains("where") || counter != 0)
                        where_clause += " and ";
                    bool isNumeric = int.TryParse(table_row.Field<string>(column_name), out temp);
                    // TODO: Different SQL condition with different type ???
                    if (isNumeric)
                        where_clause += column_name + " = " + table_row.Field<string>(column_name) + "";
                    else
                        where_clause += column_name + " = N\'" + table_row.Field<string>(column_name) + "\'";
                    counter++;
                }

            }

            if (where_clause == "")
                return query;

            if (query.ToLower().Contains("where"))
                query += where_clause;
            else
                query += " WHERE " + where_clause;

            return query;
        }

        /// <summary>
        /// Get data from database using the sql query 
        /// </summary>
        /// <param name="new_query">Pass in this value if you don't want to use the default query implemented to do the database retrieve</param>
        /// <returns></returns>
        public DataTable GetDataFromDatabaseUsingSQL(string new_query = null)
        {
            if (new_query == null)
                new_query = sql_query;

            if (new_query == null)
                return null;

            SqlConnection connection = new SqlConnection(StaticValues.connection_str);
            SqlCommand command = new SqlCommand(new_query, connection);
            connection.Open();

            SqlDataAdapter data_adapter = new SqlDataAdapter(command);
            DataTable datatable = new DataTable();
            data_adapter.Fill(datatable);
            connection.Close();
            data_adapter.Dispose();

            return datatable;
        }

        /// <summary>
        /// Fix line break (if have)
        /// </summary>
        /// <param name="document">Represent the document</param>
        public static void PostProcessorFixLineBreaks(WordprocessingDocument document)
        {
            string text = "\n";
            foreach (Text current in document.MainDocumentPart.Document.Body.Descendants<Text>())
            {
                if (current.Text.Contains(text))
                {
                    string[] arg_68_0 = current.Text.Split(new string[]
                    {
                            text
                    }, StringSplitOptions.RemoveEmptyEntries);
                    bool flag = true;
                    string[] array = arg_68_0;
                    for (int i = 0; i < array.Length; i++)
                    {
                        string text2 = array[i];
                        if (!flag)
                        {
                            current.InsertBeforeSelf<Break>(new Break());
                        }
                        flag = false;
                        current.InsertBeforeSelf<Text>(new Text
                        {
                            Text = text2
                        });
                    }
                    current.Remove();
                }
            }
            document.MainDocumentPart.Document.Save();
        }
    }
}

