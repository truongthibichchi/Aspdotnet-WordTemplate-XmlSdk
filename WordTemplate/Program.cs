using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using System;
using WordTemplate.Helpers;


// Install the DocumentFormat.OpenXML Nuget package
namespace WordTemplate
{
    class Program
    {
        static void Main(string[] args)
        {
            WordTemplateManager.Run();
            Console.WriteLine(StaticValues.logs);
        }
    }
}
