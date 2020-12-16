using System;
using ExcelDataReader;

namespace trellmeup
{
    public class Ticket
    {
       public int Id { get; set; }
       public string List { get; set; } 
       public string Title { get; set; }
       public string Description { get; set; }
       public int Points { get; set; }
       public string Due { get; set; }
       public string Members { get; set; }
       public string Labels { get; set; }
       public int CardNo { get; set; }
       public string CardURL { get; set; }
       public int Accum { get; set; }

       public void Populate(IExcelDataReader reader, int _Id)
       {
           Id = _Id;
       
           List = reader.GetString(0);
           Title = reader.GetString(1);
           Description = reader.GetString(2);
           Points = Integerize(reader, 3);
           Due = reader.GetString(4);
           Members = reader.GetString(5);
           Labels = reader.GetString(6);
           CardNo = Integerize(reader, 7);
           CardURL = reader.GetString(8);
           Accum = 0;

       }

       public int Sum(int prev){
           Accum = prev + Points;
           return Accum;
       }

        private int Integerize(IExcelDataReader reader, int position)
        {
            int result = 0;
            var fileType = reader.GetFieldType(position);
            if(fileType.Equals(typeof(System.Int32)))
            {
                return reader.GetInt32(position);
            }
            if(fileType.Equals(typeof(System.Double)))
                return Convert.ToInt32(reader.GetDouble(position));
            if(fileType.Equals(typeof(System.String)))
            {
                int.TryParse(reader.GetString(position), out result);
                return result;
            }

            return 0;

        }
    }
}
