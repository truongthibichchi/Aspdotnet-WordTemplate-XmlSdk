using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WordTemplate.Helpers
{
    class Helper_WordRepeat : Helper_WordBase
    {

        public Helper_WordRepeat(SdtElement content_control)
        {
            Init(content_control);
        }

        /// <summary>
        /// Clone the data of the 1st content control (that specified to save contents to copy) and append them to the end of the current selected content control
        /// </summary>
        public void CloneData()
        {
            // Get the "Content" content control, which is what we want to copy
            SdtElement target_copy_content_control = content_control.Descendants<SdtElement>().Where(s => s.Descendants<SdtAlias>().FirstOrDefault().Val.ToString().ToLower() == StaticValues.repeatcontent_cc_name).FirstOrDefault();
            
            content_control.AppendChild(target_copy_content_control.CloneNode(true));
       }

        public void Execute()
        {
            SdtElement repeat_sql_cc = content_control.Descendants<SdtElement>().Where(s => s.Descendants<SdtAlias>().FirstOrDefault().Val.ToString().ToLower() == StaticValues.repeatsql_cc_name).FirstOrDefault();
            sql_query = GetContentControlContents(repeat_sql_cc);
            if (!StaticValues.use_default_tablename_in_sql)
                sql_query = sql_query.Replace(StaticValues.dummy_table_name, StaticValues.project_table_name);
            DataTable result = GetDataFromDatabaseUsingSQL();
            
            for (int i = 0; i < result.Rows.Count - 1; i++)
                CloneData();
            
            List<string> columnNames = (from dc in result.Columns.Cast<DataColumn>()
                                        select dc.ColumnName.ToLower()).ToList();

            WordTemplateManager sub_manager;
            int row_count = 0;
            foreach (SdtElement repeated_cc in content_control.Descendants<SdtElement>().Where(s => s.Descendants<SdtAlias>().FirstOrDefault().Val.ToString().ToLower() == StaticValues.repeatcontent_cc_name))
            {
                foreach (SdtElement nested_cc in repeated_cc.Descendants<SdtElement>())
                {
                    SdtAlias alias = nested_cc.Descendants<SdtAlias>().FirstOrDefault();
                    if (alias.Val.ToString().ToLower() == StaticValues.repeatval_cc_name)
                    {
                        string sql_content = GetContentControlContents(nested_cc);         

                        string append_data = "";
                        // If the content sql specify the data from the above sql
                        if (columnNames.Contains(sql_content.ToLower()))
                            append_data = result.Rows[row_count][sql_content.ToLower()].ToString();
                        else
                        {
                            // If the content sql specify a query
                            if (!StaticValues.use_default_tablename_in_sql)
                                sql_content = sql_content.Replace(StaticValues.dummy_table_name, StaticValues.project_table_name);

                            DataRow conditional = result.Rows[row_count];
                            sql_content = Helper_WordBase.AddCondtionalToSQLQuery(sql_content, conditional);

                            DataTable selectvalue_result = GetDataFromDatabaseUsingSQL(sql_content);
                            if (selectvalue_result != null)
                            {
                                // TODO: Print out more values if possible
                                // Only print the 1st value found
                                append_data = selectvalue_result.Rows[0][0].ToString();
                            }
                        }
                        ReplaceContentsInContentControl(nested_cc, append_data);                    
                    }
                    else
                    {
                        DataRow conditional = result.Rows[row_count];
                        sub_manager = new WordTemplateManager(nested_cc);
                        sub_manager.Process(conditional);
                    }        
                }
                row_count++;
            }      
        }
    }
}
