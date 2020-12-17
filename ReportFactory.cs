using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace trellmeup
{
    public class ReportFactory
    {
        public void Build(){
            var date = DateTime.Now.ToString("yyyyMMddHHmmss");
             using(WordprocessingDocument doc = 
                WordprocessingDocument.Create("ReleasePlan_"+date+".docx",WordprocessingDocumentType.Document))
             {
                 MainDocumentPart mainPart = doc.AddMainDocumentPart();
                 new Document(new Body()).Save(mainPart);

                 Body body = mainPart.Document.Body;
                 body.Append(new Paragraph(new Run(new Text("Plan de Liberaciones"))));
                 body.Append(new Paragraph(new Run(new Text("La Dirección General de Tecnologías de la Información les comparte el plan de liberaciones de desarrollo del proyecto: “Sistema de Gestión de Auditorias y Seguimiento”, donde intentamos estimar cuando las funcionalidades nuevas, cambios al sistema o mejoras podrían ser entregados por el equipo de Desarrollo."))));
                 body.Append(new Paragraph(new Run(new Text("No obstante, las fallas en el sistema no entrarán en esta planificación porque serán tratados con la urgencia que estos ameriten y se corregirán cuanto antes sea posible."))));
                 body.Append(new Paragraph(new Run(new Text("Todo esto, claro, en función a la capacidad de desarrollo del equipo y respetando los intereses del Instituto Superior de Auditoria y Fiscalización, que serán resguardados por el Director General de Tecnologías de la Información."))));
                 body.Append(new Paragraph(new Run(new Text("La idea es tener una guía que refleje las expectativas acerca de lo que se llevará a cabo y cuando se liberara, naturalmente de manera especulativa y por la misma razón, se espera que cambie constantemente."))));
                 body.Append(new Paragraph(new Run(new Text("Hello World!"))));
                 mainPart.Document.Save();
             }
        }
    }
}
