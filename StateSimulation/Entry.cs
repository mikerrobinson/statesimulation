using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Bson.Serialization.IdGenerators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StateSimulation
{
    public class Entry
    {
        public static List<int> points = new List<int> { 20, 17, 16, 15, 14, 13, 12, 11, 9, 7, 6, 5, 4, 3, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

        public Entry() { }
        
        public Entry(Entry e, int order)
        {
            this.Time = e.Time;
            this.Name = e.Name;
            this.School = e.School;
            this.Year = e.Year;
            this.Event = e.Event;
            this.Rank = order;
        }

        [BsonId(IdGenerator = typeof(StringObjectIdGenerator))]
        public string ID { get; set; }

        public string Name { get; set; }
        public string School { get; set; }
        public string Year { get; set; }
        public TimeSpan Time { get; set; }
        public string Event { get; set; }
        public int Rank { get; set; }
        public int Points
        {
            get
            {
                return (Name.Equals(School)) ? points[Rank] * 2 : points[Rank];
            }
        }
    }
}
