using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Bson.Serialization.IdGenerators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StateSimulation
{
    public class Event
    {
        public Event()
        {
            this.Qualifiers = new List<Entry>();
            this.Entries = new List<Entry>();
        }

        [BsonId(IdGenerator = typeof(StringObjectIdGenerator))]
        public string ID { get; set; }
        public List<Entry> Qualifiers { get; set; }
        public List<Entry> Entries { get; set; }
        public IEnumerable<Entry> Results 
        { 
            get
            {
                return Entries.OrderBy(e => e.Time).Select((entry, index) => new Entry(entry, index));
            } 
        }
        public bool IsRelay { get; set; }
        public int Distance { get; set; }
        public string Stroke { get; set; }
        public string Gender { get; set; }
        public int Order { get; set; }
        public string DataURL 
        {
            get
            {
                return String.Format("http://www.aia365.com/leaderboards/swimming-{0}/swimming/{1}{2}{3}/d1/",
                    Gender == "F" ? "girls" : "boys",
                    Stroke,
                    IsRelay ? "relay" : "individual",
                    Distance.ToString());
            }
        }

        public override string ToString()
        {
            return String.Format("{0} {1}{2}",
                Distance.ToString(),
                Stroke,
                IsRelay ? " relay" : "");
        }
    }
}
