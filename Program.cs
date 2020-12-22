using ExcelDataReader;
using System;
using System.IO;
using System.Collections;
using System.Linq;
using System.Collections.Generic;

namespace trellmeup
{
    class Program
    {
        static void Main(string[] args)
        {
           Console.WriteLine("trello analysis generator");
           System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

           var tickets = GetTickets();
           var sprints = GetSprints();
           var sprintTickets = LoadTickets(sprints, tickets);

           ReportFactory factory = new ReportFactory();
           factory.Build(sprintTickets);
        }

        private static List<Ticket> GetTickets()
        {
            var filePath = "Observaciones.xlsx";
            var ignoredList = new List<string>(File.ReadAllLines("ignored.txt"));
            int ct = 0; 

            var tickets = new List<Ticket>();
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {                    
                    do
                    {
                        while (reader.Read())
                        {
                            if (!ignoredList.Contains(reader.GetString(1))) 
                            {
                                var ticket = new Ticket();
                                ticket.Populate(reader, ct++);
                                tickets.Add(ticket);
                            }
                        }
                    } while (reader.NextResult());
                }
            }
            
            var orderedTickets = tickets
                   .Where(x => !x.Title.Contains("[archived]"))
                   .Where(x => x.List == "Backlog" || 
                               x.List == "Sprint Backlog" ||
                               x.List == "In Progress" ||
                               x.List == "Complete" ||
                               x.List == "Code Review")
                   .OrderBy(x => x.List == "Complete"? 1 :
                                 x.List == "Code Review"? 2 :
                                 x.List == "In Progress"? 3 :
                                 x.List == "Sprint Backlog"? 4 :
                                 x.List == "Backlog"? 5 : 6)
                   .ThenBy(x => x.Id).ToList();

                   
           int accum = 0;
           foreach(Ticket ticket in orderedTickets) 
           {
               accum = ticket.Sum(accum);
               Console.WriteLine(ticket.Id + " " + ticket.Points + " " + 
                       ticket.Accum + " " + ticket.List + " " +
                       ticket.CardNo + "  " + ticket.Title);
           }

           return orderedTickets;
 
        }


        private static List<Sprint> GetSprints()
        {
            var sprints = new List<Sprint>();

            sprints.Add(new Sprint(){ Id = 1, Name = "Primera", Points = 100, DeadLine = "8 de enero." });
            sprints.Add(new Sprint(){ Id = 2, Name = "Segunda", Points = 200, DeadLine = "22 de enero." });
            sprints.Add(new Sprint(){ Id = 3, Name = "Tercera", Points = 200, DeadLine = "5 de febrero." });
            sprints.Add(new Sprint(){ Id = 4, Name = "Cuarta",  Points = 200, DeadLine = "19 de febrero." });
            sprints.Add(new Sprint(){ Id = 5, Name = "Quinta",  Points = 200, DeadLine = "5 de marzo." });
            sprints.Add(new Sprint(){ Id = 6, Name = "Sexta",   Points = 200, DeadLine = "19 de marzo." });
            sprints.Add(new Sprint() { Id = 7, Name = "Septima", Points = 200, DeadLine = "2 de abril" });
            sprints.Add(new Sprint() { Id = 8, Name = "Octava", Points = 200, DeadLine = "16 de abril" });
            sprints.Add(new Sprint() { Id = 9, Name = "Novena", Points = 200, DeadLine = "30 de abril" });
            sprints.Add(new Sprint() { Id = 10, Name = "Decima", Points = 200, DeadLine = "14 de mayo" });

            return sprints;
        }

        private static List<Sprint> LoadTickets(List<Sprint> sprints, List<Ticket> tickets)
        {
            int currentSprint = 1;
            var sprint = sprints.First(x => x.Id == currentSprint);
            int accum = sprint.Points;
            foreach (var ticket in tickets)
            {
                if(sprint.Tickets == null)
                    sprints.First(x => x.Id == currentSprint).Tickets = new List<Ticket>();

                if (ticket.Accum <= accum)
                    sprint.Tickets.Add(ticket);
                else
                {
                    sprint.Tickets.Add(ticket);
                    currentSprint++;
                    sprint = sprints.First(x => x.Id == currentSprint);
                    accum += sprint.Points;
                }                    
            }
            
            return sprints;
        }
    }
}
