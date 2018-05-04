using OfficeOpenXml; // This is the top namespace for EPPlus, if your reference isn't found use the command -> Update-Package -reinstall in the NuGet Console
using System;
using System.IO;
using System.Collections.Generic;
using WaterAnalysisTool.Loader;
using WaterAnalysisTool.Components;
using System.Text.RegularExpressions;
using WaterAnalysisTool.Analyzer;

namespace WaterAnalysisTool
{
    class Program
    {
        #region Constants
        public const double CoD_THRESHOLD = 0.7;
        public const double VERSION_NUMBER = 1.0;
        public const String COMMAND_REGEX = "[^\\s\"']+|\"([^\"]*)\"|'([^']*)'";
        public const String METHOD = "ICP-SS";
        #endregion

        public static void Main(string[] args)
        {
            FileInfo infile, outfile;
            String method;
            Double r2val;
            bool flag;

            String stringArgs = null;
            Regex r = new Regex(COMMAND_REGEX);
            MatchCollection arguments = null;

            // Startup Message
            Console.WriteLine("ICP-AES Text File Parser version " + VERSION_NUMBER + ".\nType \"usage\" for a list of commands.\n");

            #region Command Loop
            do
            {
                try
                {
                    // Reseting
                    flag = false;
                    infile = null;
                    outfile = null;
                    method = null;
                    stringArgs = null;
                    arguments = null;

                    Console.Write("Enter command: ");
                    stringArgs = Console.ReadLine();

                    if (stringArgs.ToLower().Equals("usage"))
                        Console.WriteLine("\tparse <location/name of input> <location/name for output>\n\tanalyze <location/name of input> <r^2 threshold>\n\tType \"exit\" to exit.");

                    else
                    {
                        // Check if stringArgs matches expected command structure
                        arguments = r.Matches(stringArgs);

                        if (arguments.Count > 1)
                        {
                            #region Parse Command
                            if (arguments[0].Value.ToLower().Equals("parse"))
                            {
                                if (arguments.Count > 2)
                                {
                                    #region Input cleaning
                                    // Input file cleaning
                                    String file = arguments[1].Value.Replace("\"", "").Replace("\'", ""); // Get rid of quotes
                                    if (!file.Contains(".")) // If it has no extension, add ".txt"
                                        file = file + ".txt";
                                    infile = new FileInfo(file);

                                    if (!infile.Extension.Equals(".txt"))
                                        throw new IOException("Input file must be a text file with the '.txt' extension.");

                                    // Output file cleaning
                                    file = arguments[2].Value.Replace("\"", "").Replace("\'", ""); // Get rid of quotes
                                    if (!file.Contains(".")) // If it has no extension, add ".xlsx"
                                        file = file + ".xlsx";
                                    outfile = new FileInfo(file);

                                    if (!outfile.Extension.Equals(".xlsx") && !outfile.Extension.Equals(".xls") && !outfile.Extension.Equals(".xlsm") && !outfile.Extension.Equals(".xlm"))
                                        throw new IOException("Output file must be compatible with Excel and have one of the following extensions: '.xlsx', '.xls', '.xlsm', or '.xlm'");

                                    // Method cleaning
                                    if (arguments.Count > 3) // There is an optional method argument (possibly)
                                        method = arguments[3].Value.Replace("\"", "").Replace("\'", ""); // Get rid of quotes

                                    else // There is no optional method argument, use default
                                        method = METHOD;
                                    #endregion

                                    if (infile.Exists)
                                    {
                                        if (outfile.Exists)
                                        {
                                            Console.WriteLine("\tA file of the name " + outfile.Name + " already exists at " + (outfile.ToString().Substring(0, outfile.Name.Length)) + ".");
                                            Console.Write("\tThis operation will overwrite this file. Continue? (y/n): ");

                                            if (Console.ReadLine().ToLower().Equals("n"))
                                            {
                                                Console.WriteLine("\tParse operation cancelled.");
                                            }

                                            else
                                            {
                                                outfile.Delete();

                                                using (ExcelPackage p = new ExcelPackage(outfile))
                                                {
                                                    p.Workbook.Properties.Title = arguments[2].Value.Split('.')[0];

                                                    DataLoader loader = new DataLoader(infile.OpenText(), p, method);
                                                    loader.Load();
                                                }
                                            }
                                        }

                                        else
                                        {
                                            using (ExcelPackage p = new ExcelPackage(outfile))
                                            {
                                                p.Workbook.Properties.Title = arguments[2].Value.Split('.')[0];

                                                DataLoader loader = new DataLoader(infile.OpenText(), p, method);
                                                loader.Load();
                                            }
                                        }
                                    }

                                    else
                                        Console.WriteLine("\tCould not locate " + infile.ToString());
                                }

                                else
                                    Console.WriteLine("\t" + stringArgs + " is an invalid command. For a list of valid commands enter \"usage\".");
                            }
                            #endregion

                            #region Analyze Command
                            else if (arguments[0].Value.ToLower().Equals("analyze"))
                            {
                                String file = arguments[1].Value.Replace("\"", "").Replace("\'", ""); // Get rid of quotes
                                if (!file.Contains(".")) // If it has no extension, add ".xlsx"
                                    file = file + ".xlsx";
                                infile = new FileInfo(file);

                                if (infile.Exists)
                                {
                                    if (arguments.Count > 2)
                                    {
                                        // Optional threshold argument entered
                                        if (Double.TryParse(arguments[2].Value, out r2val))
                                        {
                                            if (r2val >= 0.0 && r2val <= 1)
                                            {
                                                using (ExcelPackage p = new ExcelPackage(infile))
                                                {
                                                    foreach (ExcelWorksheet sheet in p.Workbook.Worksheets)
                                                    {
                                                        // Check if correlation worksheet already exists
                                                        if(sheet.Name.Equals("Correlation"))
                                                        {
                                                            Console.WriteLine("\tA correlation worksheet already exists for this file.");
                                                            Console.Write("\tThis operation will overwrite it. Continue? (y/n): ");

                                                            if (Console.ReadLine().ToLower().Equals("n"))
                                                            {
                                                                Console.WriteLine("\tAnalyze operation cancelled.");
                                                                flag = true;
                                                                break;
                                                            }

                                                            else
                                                            {
                                                                p.Workbook.Worksheets.Delete(sheet);
                                                                break;
                                                            }
                                                        }
                                                    }

                                                    if (!flag)
                                                    {
                                                        AnalyticsLoader analyticsLoader = new AnalyticsLoader(p, r2val);
                                                        analyticsLoader.Load();
                                                    }
                                                }
                                            }

                                            else
                                                Console.WriteLine("\t" + arguments[2] + " is an invalid threshold. Threshold must be a value between 0 and 1 inclusive.");
                                        }

                                        else
                                            Console.WriteLine("\t" + arguments[2] + " is an invalid threshold. Threshold must be numeric and a value between 0 and 1 inclusive.");
                                    }

                                    else
                                    {
                                        using (ExcelPackage p = new ExcelPackage(infile))
                                        {
                                            foreach (ExcelWorksheet sheet in p.Workbook.Worksheets)
                                            {
                                                // Check if correlation worksheet already exists
                                                if (sheet.Name.Equals("Correlation"))
                                                {
                                                    Console.WriteLine("\tA correlation worksheet already exists for this file.");
                                                    Console.Write("\tThis operation will overwrite it. Continue? (y/n): ");

                                                    if (Console.ReadLine().ToLower().Equals("n"))
                                                    {
                                                        Console.WriteLine("\tAnalyze operation cancelled.");
                                                        flag = true;
                                                        break;
                                                    }

                                                    else
                                                    {
                                                        p.Workbook.Worksheets.Delete(sheet);
                                                        break;
                                                    }
                                                }
                                            }

                                            if (!flag)
                                            {
                                                AnalyticsLoader analyticsLoader = new AnalyticsLoader(p, CoD_THRESHOLD);
                                                analyticsLoader.Load();
                                            }
                                        }
                                    }
                                }

                                else
                                    Console.WriteLine("\tCould not locate " + infile.ToString());
                            }
                            #endregion

                            else
                                Console.WriteLine("\t" + stringArgs + " is an invalid command. For a list of valid commands enter \"usage\".");
                        }

                        else
                            Console.WriteLine("\t" + stringArgs + " is an invalid command. For a list of valid commands enter \"usage\".");
                    }
                }

                // Exception catching
                catch(Exception e)
                {
                    Console.WriteLine("\t" + e.Message);
                    // Console.WriteLine("\t" + e.ToString());
                }

                Console.WriteLine(); // Some formatting
            } while (!stringArgs.ToLower().Equals("exit"));
            #endregion

            Console.WriteLine("Exiting...");
        }
    }
}
