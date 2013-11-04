
using ClosedXML.Excel;
using HtmlAgilityPack;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace StateSimulation
{


    class Program
    {
        // initialize mongo
        static MongoClient client = new MongoClient("mongodb://localhost:27017/swim");
        static MongoServer server = client.GetServer();
        static MongoDatabase database = server.GetDatabase("mrobinson-resort-guide");
        static MongoCollection<Event> events = database.GetCollection<Event>("events");

        static void Main(string[] args)
        {
            //InitializeDBFromAIA();

            //InitializeBoysExcelFilesFromDB();
            //InitializeGirlsExcelFilesFromDB();
            ScoreBoysFromExcelFiles();
            ScoreGirlsFromExcelFiles();

            Console.ReadLine();
        }

        public static void ScoreBoysFromExcelFiles()
        {
            List<Entry> boysResults = new List<Entry>();

            var book = new XLWorkbook("boys2.xlsx",XLEventTracking.Disabled);
            var boysSheet = book.Worksheet("Boys");

            var boysResults2 = new List<Entry>();

            foreach (var ev in events.FindAll().Where(e => e.Gender == "M"))
            {
                int column = (ev.Order / 2) + 2;
                for (int row = 2; row <= boysSheet.LastRowUsed().RowNumber(); row++)
                {
                    if (!boysSheet.Cell(row, column).IsEmpty())
                    {
                        ev.Entries.Add(new Entry
                        {
                            Event = boysSheet.Cell(1,column).GetValue<string>(),
                            Name = boysSheet.Cell(row, 1).GetValue<string>(),
                            School = boysSheet.Cell(row, 2).GetValue<string>(),
                            Time = TimeSpan.Parse(boysSheet.Cell(row, column).GetValue<DateTime>().TimeOfDay.ToString())//TimeSpan.Parse(rangeBoys.Cell(i, column).GetValue<string>())
                        });
                    }
                }

                boysResults2.AddRange(ev.Results);

                //var rangeBoys = boysSheet.RangeUsed();
                //int column = (ev.Order / 2) + 2;
                //var sorted = rangeBoys.Sort(column);
                //for (var i = 1; i < 17; i++)
                //{
                //    boysResults.Add(new Entry
                //    {
                //        Name = sorted.Cell(i, 1).GetValue<string>(),
                //        School = sorted.Cell(i, 2).GetValue<string>(),
                //        Rank = i,
                //        Time = TimeSpan.Parse(sorted.Cell(i,column).GetValue<DateTime>().TimeOfDay.ToString())//TimeSpan.Parse(rangeBoys.Cell(i, column).GetValue<string>())
                //    });
                //}
            }

            
            var scores = from entry in boysResults2
                         group entry by entry.School into school
                         select new
                         {
                             School = school.First().School,
                             Score = school.Sum(s => s.Points)
                         };

            Console.WriteLine("BOYS SCORES\n==============================");
            foreach (var score in scores.OrderByDescending(s => s.Score))
                Console.WriteLine(score.School + " " + score.Score.ToString());

        }

        public static void ScoreGirlsFromExcelFiles()
        {
            List<Entry> girlsResults = new List<Entry>();

            var book = new XLWorkbook("girls.xlsx", XLEventTracking.Disabled);
            var girlsSheet = book.Worksheet("Girls");

            var girlsResults2 = new List<Entry>();

            foreach (var ev in events.FindAll().Where(e => e.Gender == "F"))
            {
                int column = ev.Order + 2;
                for (int row = 2; row <= girlsSheet.LastRowUsed().RowNumber(); row++)
                {
                    if (!girlsSheet.Cell(row, column).IsEmpty())
                    {
                        ev.Entries.Add(new Entry
                        {
                            Event = girlsSheet.Cell(1, column).GetValue<string>(),
                            Name = girlsSheet.Cell(row, 1).GetValue<string>(),
                            School = girlsSheet.Cell(row, 2).GetValue<string>(),
                            Time = TimeSpan.Parse(girlsSheet.Cell(row, column).GetValue<DateTime>().TimeOfDay.ToString())//TimeSpan.Parse(rangeBoys.Cell(i, column).GetValue<string>())
                        });
                    }
                }

                girlsResults2.AddRange(ev.Results);

                //var rangeBoys = boysSheet.RangeUsed();
                //int column = (ev.Order / 2) + 2;
                //var sorted = rangeBoys.Sort(column);
                //for (var i = 1; i < 17; i++)
                //{
                //    boysResults.Add(new Entry
                //    {
                //        Name = sorted.Cell(i, 1).GetValue<string>(),
                //        School = sorted.Cell(i, 2).GetValue<string>(),
                //        Rank = i,
                //        Time = TimeSpan.Parse(sorted.Cell(i,column).GetValue<DateTime>().TimeOfDay.ToString())//TimeSpan.Parse(rangeBoys.Cell(i, column).GetValue<string>())
                //    });
                //}
            }


            var scores = from entry in girlsResults2
                         group entry by entry.School into school
                         select new
                         {
                             School = school.First().School,
                             Score = school.Sum(s => s.Points)
                         };

            Console.WriteLine("GIRLS SCORES\n==============================");
            foreach (var score in scores.OrderByDescending(s => s.Score))
                Console.WriteLine(score.School + " " + score.Score.ToString());

        }


        public static void InitializeBoysExcelFilesFromDB()
        {
            var book = new XLWorkbook();
            var boysSheet = book.Worksheets.Add("Boys");
            var girlsSheet = book.Worksheets.Add("Girls");

            Dictionary<string, Dictionary<string, string>> boys = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, Dictionary<string, string>> girls = new Dictionary<string, Dictionary<string, string>>();

            // reset from DB
            foreach (var ev in events.FindAll())
            {
                foreach (var q in ev.Qualifiers)
                {
                    var name = Regex.Replace(q.Name, @"\s+", " ") + "|" + q.School;
                    if (ev.Gender == "M")
                    {
                        if (!boys.ContainsKey(name)) boys[name] = new Dictionary<string, string>();
                        boys[name][ev.ToString()] = q.Time.ToString();
                    }
                    else
                    {
                        if (!girls.ContainsKey(name)) girls[name] = new Dictionary<string, string>();
                        girls[name][ev.ToString()] = q.Time.ToString();
                    }
                }
            }

            boysSheet.Cell(1, 1).Value = "Name";
            boysSheet.Cell(1, 2).Value = "School";
            var col = 2;
            foreach (var ev in events.FindAll().Where(e => e.Gender == "M").OrderBy(e => e.Order))
            {
                col = col + 1;
                boysSheet.Cell(1, col).Value = ev.ToString();
            }
            var row = 1;
            foreach (var boy in boys)
            {
                row = row + 1;
                boysSheet.Cell(row, 1).Value = boy.Key.Split(new char[] { '|' })[0];
                boysSheet.Cell(row, 2).Value = boy.Key.Split(new char[] { '|' })[1];
                boysSheet.Cell(row, 3).Value = boy.Value.ContainsKey("200 medley relay") ? boy.Value["200 medley relay"] : "";
                boysSheet.Cell(row, 4).Value = boy.Value.ContainsKey("200 free") ? boy.Value["200 free"] : "";
                boysSheet.Cell(row, 5).Value = boy.Value.ContainsKey("200 medley") ? boy.Value["200 medley"] : "";
                boysSheet.Cell(row, 6).Value = boy.Value.ContainsKey("50 free") ? boy.Value["50 free"] : "";
                boysSheet.Cell(row, 7).Value = boy.Value.ContainsKey("100 fly") ? boy.Value["100 fly"] : "";
                boysSheet.Cell(row, 8).Value = boy.Value.ContainsKey("100 free") ? boy.Value["100 free"] : "";
                boysSheet.Cell(row, 9).Value = boy.Value.ContainsKey("500 free") ? boy.Value["500 free"] : "";
                boysSheet.Cell(row, 10).Value = boy.Value.ContainsKey("200 free relay") ? boy.Value["200 free relay"] : "";
                boysSheet.Cell(row, 11).Value = boy.Value.ContainsKey("100 back") ? boy.Value["100 back"] : "";
                boysSheet.Cell(row, 12).Value = boy.Value.ContainsKey("100 breast") ? boy.Value["100 breast"] : "";
                boysSheet.Cell(row, 13).Value = boy.Value.ContainsKey("400 free relay") ? boy.Value["400 free relay"] : "";
                boysSheet.Cell(row, 14).FormulaR1C1 = "=COUNTA(RC[-11]:RC[-1])";
            }
            book.SaveAs("boys.xlsx");
        }

        public static void InitializeGirlsExcelFilesFromDB()
        {
            var book = new XLWorkbook();
            //var boysSheet = book.Worksheets.Add("Boys");
            var girlsSheet = book.Worksheets.Add("Girls");

            Dictionary<string, Dictionary<string, string>> boys = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, Dictionary<string, Entry>> girls = new Dictionary<string, Dictionary<string, Entry>>();

            // reset from DB
            foreach (var ev in events.FindAll())
            {
                foreach (var q in ev.Qualifiers)
                {
                    var name = Regex.Replace(q.Name, @"\s+", " ") + "|" + q.School;
                    if (ev.Gender == "M")
                    {
                        if (!boys.ContainsKey(name)) boys[name] = new Dictionary<string, string>();
                        boys[name][ev.ToString()] = q.Time.ToString() + "| (" + q.Rank.ToString() + ")";
                    }
                    else
                    {
                        if (!girls.ContainsKey(name)) girls[name] = new Dictionary<string, Entry>();
                        girls[name][ev.ToString()] = q;// q.Time.ToString() + "| (" + q.Rank.ToString() + ")";
                    }
                }
            }

            girlsSheet.Cell(1, 1).Value = "Name";
            girlsSheet.Cell(1, 2).Value = "School";
            var col = 1;
            foreach (var ev in events.FindAll().Where(e => e.Gender == "F").OrderBy(e => e.Order))
            {
                col = col + 2;
                girlsSheet.Cell(1, col).Value = ev.ToString();
            }
            var row = 1;
            foreach (var boy in girls)
            {
                row = row + 1;
                girlsSheet.Cell(row, 1).Value = boy.Key.Split(new char[] { '|' })[0];
                girlsSheet.Cell(row, 2).Value = boy.Key.Split(new char[] { '|' })[1];
                girlsSheet.Cell(row, 3).Value = boy.Value.ContainsKey("200 medley relay") ? boy.Value["200 medley relay"].Time.ToString() : "";
                girlsSheet.Cell(row, 5).Value = boy.Value.ContainsKey("200 free") ? boy.Value["200 free"].Time.ToString() : "";
                girlsSheet.Cell(row, 7).Value = boy.Value.ContainsKey("200 medley") ? boy.Value["200 medley"].Time.ToString() : "";
                girlsSheet.Cell(row, 9).Value = boy.Value.ContainsKey("50 free") ? boy.Value["50 free"].Time.ToString() : "";
                girlsSheet.Cell(row, 11).Value = boy.Value.ContainsKey("100 fly") ? boy.Value["100 fly"].Time.ToString() : "";
                girlsSheet.Cell(row, 13).Value = boy.Value.ContainsKey("100 free") ? boy.Value["100 free"].Time.ToString() : "";
                girlsSheet.Cell(row, 15).Value = boy.Value.ContainsKey("500 free") ? boy.Value["500 free"].Time.ToString() : "";
                girlsSheet.Cell(row, 17).Value = boy.Value.ContainsKey("200 free relay") ? boy.Value["200 free relay"].Time.ToString() : "";
                girlsSheet.Cell(row, 19).Value = boy.Value.ContainsKey("100 back") ? boy.Value["100 back"].Time.ToString() : "";
                girlsSheet.Cell(row, 21).Value = boy.Value.ContainsKey("100 breast") ? boy.Value["100 breast"].Time.ToString() : "";
                girlsSheet.Cell(row, 23).Value = boy.Value.ContainsKey("400 free relay") ? boy.Value["400 free relay"].Time.ToString() : "";

                girlsSheet.Cell(row, 4).Value = boy.Value.ContainsKey("200 medley relay") ? boy.Value["200 medley relay"].Rank.ToString() : "";
                girlsSheet.Cell(row, 6).Value = boy.Value.ContainsKey("200 free") ? boy.Value["200 free"].Rank.ToString() : "";
                girlsSheet.Cell(row, 8).Value = boy.Value.ContainsKey("200 medley") ? boy.Value["200 medley"].Rank.ToString() : "";
                girlsSheet.Cell(row, 10).Value = boy.Value.ContainsKey("50 free") ? boy.Value["50 free"].Rank.ToString() : "";
                girlsSheet.Cell(row, 12).Value = boy.Value.ContainsKey("100 fly") ? boy.Value["100 fly"].Rank.ToString() : "";
                girlsSheet.Cell(row, 14).Value = boy.Value.ContainsKey("100 free") ? boy.Value["100 free"].Rank.ToString() : "";
                girlsSheet.Cell(row, 16).Value = boy.Value.ContainsKey("500 free") ? boy.Value["500 free"].Rank.ToString() : "";
                girlsSheet.Cell(row, 18).Value = boy.Value.ContainsKey("200 free relay") ? boy.Value["200 free relay"].Rank.ToString() : "";
                girlsSheet.Cell(row, 20).Value = boy.Value.ContainsKey("100 back") ? boy.Value["100 back"].Rank.ToString() : "";
                girlsSheet.Cell(row, 22).Value = boy.Value.ContainsKey("100 breast") ? boy.Value["100 breast"].Rank.ToString() : "";
                girlsSheet.Cell(row, 24).Value = boy.Value.ContainsKey("400 free relay") ? boy.Value["400 free relay"].Rank.ToString() : ""; 
                girlsSheet.Cell(row, 25).FormulaR1C1 = "=COUNTA(RC[-11]:RC[-1])";
            }
            book.SaveAs("girls.xlsx");
        }

        public static void InitializeDBFromAIA()
        {
            // initialize events
            List<Event> meet = new List<Event>
            {
                new Event { Order = 1, Gender = "F", Distance = 200, Stroke = "medley", IsRelay = true },
                new Event { Order = 3, Gender = "F", Distance = 200, Stroke = "free", IsRelay = false },
                new Event { Order = 5, Gender = "F", Distance = 200, Stroke = "medley", IsRelay = false },
                new Event { Order = 7, Gender = "F", Distance = 50, Stroke = "free", IsRelay = false },
                new Event { Order = 9, Gender = "F", Distance = 100, Stroke = "fly", IsRelay = false },
                new Event { Order = 11, Gender = "F", Distance = 100, Stroke = "free", IsRelay = false },
                new Event { Order = 13, Gender = "F", Distance = 500, Stroke = "free", IsRelay = false },
                new Event { Order = 15, Gender = "F", Distance = 200, Stroke = "free", IsRelay = true },
                new Event { Order = 17, Gender = "F", Distance = 100, Stroke = "back", IsRelay = false },
                new Event { Order = 19, Gender = "F", Distance = 100, Stroke = "breast", IsRelay = false },
                new Event { Order = 21, Gender = "F", Distance = 400, Stroke = "free", IsRelay = true },

                new Event { Order = 2, Gender = "M", Distance = 200, Stroke = "medley", IsRelay = true },
                new Event { Order = 4, Gender = "M", Distance = 200, Stroke = "free", IsRelay = false },
                new Event { Order = 6, Gender = "M", Distance = 200, Stroke = "medley", IsRelay = false },
                new Event { Order = 8, Gender = "M", Distance = 50, Stroke = "free", IsRelay = false },
                new Event { Order = 10, Gender = "M", Distance = 100, Stroke = "fly", IsRelay = false },
                new Event { Order = 12, Gender = "M", Distance = 100, Stroke = "free", IsRelay = false },
                new Event { Order = 14, Gender = "M", Distance = 500, Stroke = "free", IsRelay = false },
                new Event { Order = 16, Gender = "M", Distance = 200, Stroke = "free", IsRelay = true },
                new Event { Order = 18, Gender = "M", Distance = 100, Stroke = "back", IsRelay = false },
                new Event { Order = 20, Gender = "M", Distance = 100, Stroke = "breast", IsRelay = false },
                new Event { Order = 22, Gender = "M", Distance = 400, Stroke = "free", IsRelay = true }
            };

            // clear mongo collection
            events.RemoveAll();

            HtmlWeb web = new HtmlWeb();
            List<Entry> fEntries = new List<Entry>();
            List<Entry> mEntries = new List<Entry>();

            foreach (var swimEvent in meet)
            {
                Console.Write("Loading " + swimEvent.ToString() + "... ");
                HtmlDocument doc = web.Load(swimEvent.DataURL);
                HtmlNodeCollection rows = doc.DocumentNode.SelectNodes(".//tbody/tr");
                Console.Write("processing... ");
                foreach (var row in rows)
                {
                    var data = row.SelectNodes(".//td");
                    var name = swimEvent.IsRelay ? data[2].InnerText.Trim() : data[1].InnerText.Trim();
                    
                    swimEvent.Qualifiers.Add(new Entry
                    {
                        Event = swimEvent.ToString(),
                        Name = swimEvent.IsRelay ? data[2].InnerText.Trim() : data[1].InnerText.Trim(),
                        School = data[2].InnerText.Trim(),
                        Year = swimEvent.IsRelay ? "" : data[3].InnerText.Trim(),
                        Rank = int.Parse(data[0].InnerText.Trim()),
                        Time = TimeSpan.Parse("00:" + (swimEvent.IsRelay ? data[3].InnerText.Trim() : data[4].InnerText.Trim()))
                    });
                }
                events.Save(swimEvent);
                Console.WriteLine("done");
            }
        }
    }
}
