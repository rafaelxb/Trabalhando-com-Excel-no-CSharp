using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HealthScore
{
    public class Events
    {
        public int Id { get; set;}
        public string Name { get; set; }

        public int HealthScoreDiscount { get; set; }

        public Events(int id, string name, int healthScoreDiscount)
        {
            Id = id;
            Name = name;
            HealthScoreDiscount = healthScoreDiscount;
        }

    }
}
