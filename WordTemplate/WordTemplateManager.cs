using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using WordTemplate.Helpers;

namespace WordTemplate
{
    public class WordTemplateManager
    {
        public static WordprocessingDocument document;
        protected SdtElement content_control;

        /// <summary>
        /// Construtor for the class to work with the whole document
        /// </summary>
        /// <param name="doc">Represent a document</param>
        public WordTemplateManager(WordprocessingDocument doc)
        {
            document = doc;
        }

        /// <summary>
        /// Some data are required. Call this function and pass in required data
        /// </summary>
        /// <param name="word_template_location">Local location of input word template</param>
        /// <param name="target_table_name">The table in which we going to get data from</param>
        /// <param name="connection_str">Connection string of the database we work on</param>
        public static bool Init(string word_template_location, string target_table_name, string connection_str)
        {
            // Update config file location
            StaticValues.configfile_name = "\\config.txt";

            StaticValues.configfile_name = System.IO.Path.GetDirectoryName(word_template_location) + StaticValues.configfile_name;
            
            // TODO: Check validity of these functions
            StaticValues.ReadFromFile();        // Load config
            // Unload serialized data
            StaticValues.table_value_with_color_properties = JsonConvert.DeserializeObject<Dictionary<string, string>>(StaticValues.table_value_with_color_properties_json);

            // Duplicate template to store result
            int index = word_template_location.LastIndexOf('.');
            string result_word_path = word_template_location.Substring(0, index) + StaticValues.word_result_suffix + "." + word_template_location.Substring(index + 1);
            File.Copy(word_template_location, result_word_path, true);

            // Init values
            StaticValues.word_template_path = result_word_path;
            StaticValues.project_table_name = target_table_name.Split('.')[0];
            StaticValues.connection_str = connection_str;

            return true;
        }

        /// <summary>
        /// This function execute the code for word template fill-in. Call this after calling the Init() function
        /// </summary>
        public static void Run()
        {
            try
            {
                WordprocessingDocument document = WordprocessingDocument.Open(StaticValues.word_template_path, true);
                WordTemplateManager manager = new WordTemplateManager(document);
                manager.Process();
                Helper_WordBase.PostProcessorFixLineBreaks(document);
                Helper_WordBase.DeleteAllSQLQueryInWord(document);
                document.Save();
                document.Close();
            }
            catch (Exception ex)
            {
                //StaticValues.logs += ex.Message + Environment.NewLine;
                StaticValues.logs += ex.ToString() + Environment.NewLine;
                StaticValues.is_success = false;
                
                // Write to debug screen on VS
                System.Diagnostics.Debug.WriteLine(StaticValues.logs);
            }
        }

        /// <summary>
        /// Constructor for the class to work with a part of the document, specified by a content control
        /// </summary>
        /// <param name="content_control">SdtElement instance represents a content control</param>
        public WordTemplateManager(SdtElement content_control)
        {
            this.content_control = content_control;
        }

        /// <summary>
        /// Main function to scan and fill data in content control in word template
        /// </summary>
        /// <param name="condional">A DataRow instance contains required values for conditional (used in repeat content control)</param>
        public void Process(DataRow condional = null)
        {
            List<SdtElement> content_control_list = new List<SdtElement>();

            if (content_control != null)
                content_control_list.Add(content_control);       
            else
                content_control_list = document.MainDocumentPart.Document.Body.Elements<SdtElement>().ToList();

            foreach (SdtElement content_control in content_control_list)
                ProcessBasedOnContentControl(content_control, condional);

        }

        /// <summary>
        /// Determine the type of content control and take neccessary actions
        /// </summary>
        /// <param name="content_control">A SdteElement instance represent a content control</param>
        /// <param name="condional">A DataRow instance contains required values for conditional (used in repeat content control)</param>
        protected void ProcessBasedOnContentControl(SdtElement content_control, DataRow conditional)
        {
            SdtAlias alias = content_control.Descendants<SdtAlias>().FirstOrDefault();

            if (alias.Val.ToString().ToLower() == StaticValues.repeat_cc_name)
            {
                Helper_WordRepeat repeat_helper = new Helper_WordRepeat(content_control);
                repeat_helper.Execute();
            }
            else if (alias.Val.ToString().ToLower() == StaticValues.columntable_cc_name)
            {
                Helper_WordTable_Columns table_helper = new Helper_WordTable_Columns(content_control);
                table_helper.AddConditionToSQLQuery(conditional);
                table_helper.AddDataToTable();
            }
            else if (alias.Val.ToString().ToLower() == StaticValues.rowtable_cc_name)
            {
                Helper_WordTable_Rows table_helper = new Helper_WordTable_Rows(content_control);
                table_helper.AddConditionToSQLQuery(conditional);
                table_helper.AddDataToTable();
            }
            else if (alias.Val.ToString().ToLower() == StaticValues.hybridtable_cc_name)
            {
                Helper_WordTable_Hybrid table_helper = new Helper_WordTable_Hybrid(content_control);
                table_helper.AddConditionToSQLQuery(conditional);
                table_helper.AddDataToTable();
            }
            else if (alias.Val.ToString().ToLower() == StaticValues.barchart_cc_name)
            {
                Helper_WordBarChart chart_helper = new Helper_WordBarChart(content_control);
                chart_helper.AddConditionToSQLQuery(conditional);
                chart_helper.UpdateChartFromSQL();
            }
            else if (alias.Val.ToString().ToLower() == StaticValues.piechart_cc_name)
            {
                Helper_WordPieChart chart_helper = new Helper_WordPieChart(content_control);
                chart_helper.AddConditionToSQLQuery(conditional);
                chart_helper.UpdateChartFromSQL();
            }
            else if (alias.Val.ToString().ToLower() == StaticValues.linechart_cc_name)
            {
                Helper_WordLineChart chart_helper = new Helper_WordLineChart(content_control);
                chart_helper.AddConditionToSQLQuery(conditional);
                chart_helper.UpdateChartFromSQL();
            }
            else if (alias.Val.ToString().ToLower() == StaticValues.val_cc_name)
            {
                Helper_WordBase helper = new Helper_WordBase();
                helper.Init(content_control);
                helper.AddConditionToSQLQuery(conditional);
                string content = helper.GetDataFromDatabaseUsingSQL().Rows[0][0].ToString();

                if (Helper_WordPicture.IsUrl(content) && Helper_WordPicture.IsImageUrl(content))
                {
                    // It is picture
                    SdtElement image_cc = content_control.Descendants<SdtElement>().Where(s => s.Descendants<SdtAlias>().FirstOrDefault().Val.ToString().ToLower() == StaticValues.image_cc_name).FirstOrDefault();
                    Helper_WordPicture pic_helper = new Helper_WordPicture(image_cc, content);
                    pic_helper.AddPictureFromUri();

                }
                else // Just normal text 
                    Helper_WordBase.ReplaceContentsInContentControl(content_control, content);
            }
        }
    }
}
