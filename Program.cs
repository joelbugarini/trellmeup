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

            var filePath = "Observaciones.xlsx";
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
                           try{
                                var ticket = new Ticket();
                                ticket.Populate(reader, ct++);
                                tickets.Add(ticket);
                           } catch{}
                                                        
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
                   .ThenBy(x => x.Id);
                   
           int accum = 0;
           foreach(Ticket ticket in orderedTickets) 
           {

               accum = ticket.Sum(accum);
               Console.WriteLine(ticket.Id + " " + ticket.Points + " " + ticket.Accum + " " + ticket.List + " " +ticket.CardNo + "  " + ticket.Title);
           }
        }
    }
}
