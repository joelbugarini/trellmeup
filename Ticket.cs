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
       public string Area { get; set; }
       public string Extract { get; set; }
       public Sprint Sprint { get; set; }

       public void Populate(IExcelDataReader reader, int _Id)
       {
           Id = _Id;
            List = reader.GetString(0);
           Title = reader.GetString(1);
           Description = GetDescription(reader.GetString(2));
           Extract = GetExtract(reader.GetString(2));
           Points = Integerize(reader, 3);
           Due = reader.GetString(4);
           Members = reader.GetString(5);
           Labels = reader.GetString(6);
           Area = GetArea(reader.GetString(6));
           CardNo = Integerize(reader, 7);
           CardURL = reader.GetString(8);
           Accum = 0;
       }

        private string GetExtract(string text)
        {
            if (text == null)
                return "";

            int start = text.IndexOf("[[");
            int end = text.IndexOf("]]");

            if (start >= 0)
                return text.Substring(start + 2, end - start - 2);
            else
                return "";
        }


       private string GetArea(string text)
       {
            if(text == null)
               return "";

           if(text.Contains("Estado")) return "Estado";
           if(text.Contains("Municipios")) return "Municipios";
           if(text.Contains("Calidad")) return "Calidad";
           if(text.Contains("Mayor")) return "Mayor";
           if(text.Contains("Obras")) return "Obras";
           if(text.Contains("Sistema")) return "Sistema";

           return "";
       }

        private string GetDescription(string text)
        {
            if (text == null)
                return "";

            int start = text.IndexOf("[[");
            int end = text.IndexOf("]]");
            if (start >= 0)
                return text.Remove(start, end - start - 1);
            else
                return text;
        }

       public int Sum(int prev){
           Accum = prev + Points;
           return Accum;
       }

        private int Integerize(IExcelDataReader reader, int position)
        {
            int result = 0;
            var fileType = reader.GetFieldType(position);

            if (fileType == null)
                return 0;

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
