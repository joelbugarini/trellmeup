using System;
using System.Collections;
using System.Linq;
using System.Collections.Generic;


namespace trellmeup
{
    public class Sprint
    {
       public int Id { get; set; }
       public string Name { get; set; }
       public int Points { get; set; } 
       public string DeadLine { get; set; }

       public List<Ticket> Tickets { get; set; }
    }
}
