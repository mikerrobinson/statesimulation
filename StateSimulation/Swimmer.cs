using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Bson.Serialization.IdGenerators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StateSimulation
{
    public class Swimmer
    {
        [BsonId(IdGenerator = typeof(StringObjectIdGenerator))]
        public string ID { get; set; }
        public string Name { get; set; }
        public string School { get; set; }
        public string Year { get; set; }
        public List<Entry> Entries { get; set; }
    }
}
