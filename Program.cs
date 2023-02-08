using CsvHelper;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;

namespace ueloc
{
    class Program
    {
        static void Main(string[] args)
        {
            //args 0 : source xlsx
            //args 1 : csv source strings output path
            //args 2 : po output dir
            //args 3 : po file name
            //args 4 : replace regex

            if (args.Length < 4){
                Console.WriteLine("Invalid arguments");
                return;
            }

            string filePath = args[0];
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()))
                {
                    Dictionary<string, List<LocalizedString>> culturalStrings = new Dictionary<string, List<LocalizedString>>();
                    List<string> cultures = new List<string>();

                    var result = reader.AsDataSet();
                    var table = result.Tables[0];

                    SaveCultureData(culturalStrings, cultures, table);
                    ParseCulturalStrings(culturalStrings, cultures, table);

                    foreach (var culturalString in culturalStrings)
                    {
                        var strings = culturalString.Value;

                        //Regex regex = new Regex("{{\\w.*}}");
                        Regex regex = new Regex(args[4]);
                        for (int i = 0; i < strings.Count; i++)
                        {
                            LocalizedString s = strings[i];
                            //var match = regex.Match(s.target);
                            var matches = regex.Matches(s.target);
                            foreach (Match match in matches)
                            {
                                if (match.Success)
                                {
                                    ReplaceIdWithString(strings, ref s, regex, match);
                                }
                            }
                        }
                    }

                    MakeSourceCSV(args, culturalStrings, cultures);

                    foreach(string culture in cultures)
                    {
                        if (culture.Contains("Empty_"))
                        {
                            continue;
                        }
                        //Make cultural directories
                        string path = args[2] + "\\" + culture;
                        Directory.CreateDirectory(path);

                        List<LocalizedString> localizedStrings = culturalStrings[culture];

                        //Make cultural po
                        using (var writer = new StreamWriter(path + "\\" + args[3]))
                        {
                            writer.WriteLine();
                            writer.WriteLine();

                            foreach(var locStr in localizedStrings)
                            {
                                writer.WriteLine("msgctxt \"," + locStr.key + "\"");
                                writer.WriteLine("msgid \"" + locStr.source + "\"");
                                writer.WriteLine("msgstr \"" + locStr.target + "\"");
                                writer.WriteLine();
                                writer.WriteLine();
                            }
                        }
                    }
                }
            }
        }

        private static void ReplaceIdWithString(List<LocalizedString> strings, ref LocalizedString s, Regex regex, Match match)
        {
            string searchId = match.Value;
            searchId = searchId.Replace("{", "");
            searchId = searchId.Replace("}", "");
            
            if(searchId == s.key)
            {
                Console.WriteLine("Key error : " + searchId);
                Console.ReadKey();
                throw new InvalidDataException("Key error : " + searchId);
            }

            bool wasIdFound = false;
            string ReplaceTargetString = "";
            string ReplaceSourceString = "";

            for (int i = 0; i < strings.Count; i++)
            {
                LocalizedString localizedString = strings[i];
                if (localizedString.key == searchId)
                {
                    var nestedMatches = regex.Matches(localizedString.target);
                    foreach (Match nestedMatch in nestedMatches)
                    {
                        if (nestedMatch.Success)
                        {
                            ReplaceIdWithString(strings, ref localizedString, regex, nestedMatch);
                        }
                    }

                    ReplaceSourceString = localizedString.source;
                    ReplaceTargetString = localizedString.target;
                    wasIdFound = true;
                    break;
                }
            }

            if (wasIdFound)
            {
                s.source = s.source.Replace(match.Value, ReplaceSourceString);
                s.target = s.target.Replace(match.Value, ReplaceTargetString);
            }
        }

        private static void MakeSourceCSV(string[] args, Dictionary<string, List<LocalizedString>> culturalStrings, List<string> cultures)
        {
            using (var writer = new StreamWriter(args[1]))
            {
                using (CsvWriter csv = new CsvWriter(writer, System.Globalization.CultureInfo.InvariantCulture))
                {
                    List<SourceValue> csvStrings = new List<SourceValue>();
                    var sourceStrings = culturalStrings[cultures[0]];
                    foreach (var sourceStr in sourceStrings)
                    {
                        csvStrings.Add(new SourceValue() { Key = sourceStr.key, SourceString = sourceStr.source });
                    }

                    csv.WriteRecords(csvStrings);
                }
            }
        }

        private static void ParseCulturalStrings(Dictionary<string, List<LocalizedString>> culturalStrings, List<string> cultures, System.Data.DataTable table)
        {
            for (int i = 1; i < table.Rows.Count; i++)
            {
                var values = table.Rows[i].ItemArray;

                string keyStr = values[0].ToString();
                string sourceStr = values[1].ToString();

                if(keyStr.Length == 0)
                {
                    continue;
                }

                for (int j = 1; j < values.Length; j++)
                {
                    string culture = cultures[j - 1];

                    //reader.

                    LocalizedString localizedString = new LocalizedString();
                    localizedString.key = keyStr;
                    localizedString.source = sourceStr;
                    localizedString.target = values[j].ToString();

                    List<LocalizedString> strings = new List<LocalizedString>();
                    culturalStrings.TryGetValue(culture, out strings);
                    strings.Add(localizedString);
                }
            }
        }

        private static void SaveCultureData(Dictionary<string, List<LocalizedString>> culturalStrings, List<string> cultures, System.Data.DataTable table)
        {
            var firstRow = table.Rows[0];
            culturalStrings.Clear();
            cultures.Clear();
            for (int i = 1; i < table.Columns.Count; i++)
            {
                string culture = firstRow.ItemArray[i].ToString();
                if(culture.Length == 0)
                {
                    culture = "Empty_"+i.ToString();
                }
                culturalStrings.Add(culture, new List<LocalizedString>());
                cultures.Add(culture);
            }
        }

        class LocalizedString 
        {
            public string key { get; set; }
            public string source { get; set; }
            public string target { get; set; }
        }

        struct SourceValue
        {
            public string Key { get; set; }
            public string SourceString { get; set; }
        }
    }
}
