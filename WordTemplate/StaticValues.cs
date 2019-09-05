using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace WordTemplate
{
    public class StaticValues
    {
        // XML Namespace required for Microsoft Word
        public static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";

        public static XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        public static XNamespace customPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        public static XNamespace customVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

        public static XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public static XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        public static XNamespace n = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
        public static XNamespace v = "urn:schemas-microsoft-com:vml";

        // Error logs
        public static string logs = "";
        public static bool is_success = true;

        //-----------------------------------------------------------------------------------------------------------
        // Word template
        public static string word_template_path = @"F:\sample.docx";
        public static string word_result_suffix = "_result";

        // Word properties
        public static string default_font = "Calibri";
        public static string default_font_size = "18";

        // Chart
        public static bool barchart_rotatedtable = false;
        public static bool piechart_rotatedtable = true;
        public static bool linechart_rotatedtable = false;

        // Content control (note: Write in lowercase only)
        public static string barchart_cc_name = "bar_chart";
        public static string piechart_cc_name = "pie_chart";
        public static string linechart_cc_name = "line_chart";
        public static string hybridtable_cc_name = "raw_table_rows_v2";
        public static string rowtable_cc_name = "raw_table_rows";
        public static string columntable_cc_name = "raw_table_columns";
        public static string content_cc_name = "content";
        public static string sql_cc_name = "selectrows";
        public static string repeat_cc_name = "repeat";
        public static string repeatcontent_cc_name = "rcontent";
        public static string val_cc_name = "selectvalue";
        public static string repeatval_cc_name = "selectvalue";          // Use with repeat content control
        public static string repeatsql_cc_name = "selectrepeatingdata";  // Use with repeat content control
        public static string value_placeholder_name = "dummyvalue";
        public static string image_cc_name = "picture";


        // Table content control
        // Apply color to a cell with a specific value
        // Note: write value and key in lowercase
        public static Dictionary<string, string> table_value_with_color_properties = new Dictionary<string, string>();
        public static string table_value_with_color_properties_json;  

        // Others
        public static string no_data_found_str = "No data found";
        public static string connection_str = @"Data Source=.;Initial Catalog = ServiceManagement; Integrated Security = True";
        public static string dummy_table_name = "queryTable";
        public static string project_table_name = "Project1";
        public static bool use_default_tablename_in_sql = false;

        static StaticValues()
        {
            // Color value can be found by type "Color picker" into Google.com
            // Or here https://www.w3schools.com/colors/colors_picker.asp
            table_value_with_color_properties.Add("critical", "e03838");
            table_value_with_color_properties.Add("high", "ff570a");
            table_value_with_color_properties.Add("medium", "eeff00");
            table_value_with_color_properties.Add("low", "e2ffc9");
            table_value_with_color_properties_json = JsonConvert.SerializeObject(table_value_with_color_properties, Newtonsoft.Json.Formatting.Indented);
        }

        // Variables for the file
        public static string configfile_name = "\\config.txt";
        private static Dictionary<string, Reference> config_items = new Dictionary<string, Reference> {
            { "Font", new Reference(() => default_font, val => { default_font = (string) val; })},
            { "Font_Size", new Reference(() => default_font_size, val => { default_font_size = (string) val; }) },
            { "Word_Result_Suffix", new Reference(() => word_result_suffix, val => { word_result_suffix = (string) val; } )},

            { "BarChart_RotatedTable", new Reference(() => barchart_rotatedtable, val => { barchart_rotatedtable = (bool) val; })},
            { "PieChart_RotatedTable", new Reference(() => piechart_rotatedtable, val => { piechart_rotatedtable = (bool) val; }) },
            { "LineChart_RotatedTable", new Reference(() => linechart_rotatedtable, val => { linechart_rotatedtable = (bool) val; } )},


            { "Barchart_CC_Name", new Reference(() => barchart_cc_name, val => { barchart_cc_name = (string) val; }) },
            { "Piechart_CC_Name", new Reference(() => piechart_cc_name, val => { piechart_cc_name = (string) val; }) },
            { "Linechart_CC_Name", new Reference(() => linechart_cc_name, val => { linechart_cc_name = (string) val; }) },
            { "Hybridtable_CC_Name", new Reference(() => hybridtable_cc_name, val => { hybridtable_cc_name = (string) val; }) },
            { "Table_Row_CC_Name", new Reference(() => rowtable_cc_name, val => { rowtable_cc_name = (string) val; }) },
            { "Table_Column_CC_Name", new Reference(() => columntable_cc_name, val => { columntable_cc_name = (string) val; }) },
            { "Content_CC_Name", new Reference(() => content_cc_name, val => { content_cc_name = (string) val; }) },
            { "SQL_CC_Name", new Reference(() => sql_cc_name, val => { sql_cc_name = (string) val; }) },
            { "RepeatContent_CC_Name", new Reference(() => repeatcontent_cc_name, val => { repeatcontent_cc_name = (string) val; }) },
            { "Repeat_CC_Name", new Reference(() => repeat_cc_name, val => { repeat_cc_name = (string) val; }) },
            { "Repeat_SQL_Name", new Reference(() => repeatsql_cc_name, val => { repeatsql_cc_name = (string) val; }) },
            { "Value_CC_name", new Reference(() => val_cc_name, val => { val_cc_name = (string) val; }) },
            { "Repeat_Value_CC_Name", new Reference(() => repeatval_cc_name, val => { repeatval_cc_name = (string) val; }) },
            { "Place_Holder_Name", new Reference(() => value_placeholder_name, val => { value_placeholder_name = (string) val; }) },
            { "Image_CC_Name", new Reference(() => image_cc_name, val => { image_cc_name = (string) val; }) },
            { "NoDataFound_Message", new Reference(() => no_data_found_str, val => { no_data_found_str = (string) val; }) },
            { "Connection_String", new Reference(() => connection_str, val => { connection_str = (string) val; }) },
            { "Dummy_Table_Name", new Reference(() => dummy_table_name, val => { dummy_table_name = (string) val; }) },
            { "Use_Default_Table_Name", new Reference(() => use_default_tablename_in_sql, val => { use_default_tablename_in_sql = (bool) val; }) },
            { "Cell_color_properties_with_Text", new Reference(() => table_value_with_color_properties_json, val => { table_value_with_color_properties_json = (string) val; })}
        };

        // Call this method to read the config file or create if it doesn't exist.
        public static void ReadFromFile()
        {
            if (File.Exists(configfile_name))
            {
                var config = new StreamReader(configfile_name);
                try
                {
                    while (!config.EndOfStream)
                    {
                        string[] current = config.ReadLine().Split(':');
                        string newVal = current[1].Replace(" ", string.Empty);
                        config_items[current[0]].Set(newVal);
                    }
                }
                catch (Exception e)
                {
                    // If the reading fails for some reason, this will reset the file to default.
                    Console.WriteLine(e.Message);
                    config.Close();
                    ResetToDefault();
                }
                config.Close();
            }
            else
            { ResetToDefault(); }

        }

        // Resets all values and labels in the config file to the defaults.
        public static void ResetToDefault()
        {
            var file = new StreamWriter(configfile_name);
            foreach (KeyValuePair<string, Reference> pair in config_items)
            {
                string line = pair.Key + ": " + pair.Value.Get();
                file.WriteLine(line);
            }
            file.Close();
        }

    }

    // This stores a reference to a variable so that it can be modified as part of a list.
    class Reference
    {
        public Func<object> Get { get; private set; }
        public Action<object> Set { get; private set; }
        public Reference(Func<object> getter, Action<object> setter)
        {
            Get = getter;
            Set = setter;
        }
    }
}
