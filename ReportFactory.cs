using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections;
using System.Linq;
using System.Collections.Generic;

namespace trellmeup
{
    public class ReportFactory
    {
        public void Build(){
            var date = DateTime.Now.ToString("yyyyMMddHHmmss");
            var Sprints = new List<Sprint>();
             using(WordprocessingDocument doc = 
                WordprocessingDocument.Create("ReleasePlan_"+date+".docx",WordprocessingDocumentType.Document))
             {
                 MainDocumentPart mainPart = doc.AddMainDocumentPart();
                 new Document(new Body()).Save(mainPart);

                 NumberingPart(mainPart);

                 Body body = mainPart.Document.Body;
                 body.Append(new Paragraph(new Run(new Text("Plan de Liberaciones"))));
                 body.Append(new Paragraph(new Run(new Text("La Dirección General de Tecnologías de la Información les comparte el plan de liberaciones de desarrollo del proyecto: “Sistema de Gestión de Auditorias y Seguimiento”, donde intentamos estimar cuando las funcionalidades nuevas, cambios al sistema o mejoras podrían ser entregados por el equipo de Desarrollo."))));
                 body.Append(new Paragraph(new Run(new Text("No obstante, las fallas en el sistema no entrarán en esta planificación porque serán tratados con la urgencia que estos ameriten y se corregirán cuanto antes sea posible."))));
                 body.Append(new Paragraph(new Run(new Text("Todo esto, claro, en función a la capacidad de desarrollo del equipo y respetando los intereses del Instituto Superior de Auditoria y Fiscalización, que serán resguardados por el Director General de Tecnologías de la Información."))));
                 body.Append(new Paragraph(new Run(new Text("La idea es tener una guía que refleje las expectativas acerca de lo que se llevará a cabo y cuando se liberara, naturalmente de manera especulativa y por la misma razón, se espera que cambie constantemente."))));
                 body.Append(new Paragraph(new Run(new Text("Metas de liberación"))));
                 body.Append(new Paragraph(new Run(new Text("A continuación, presentamos las metas generales que el equipo de desarrollo pretende alcanzar a mediano plazo referente al “Sistema de Gestión de Auditorias y Seguimiento” y sistemas secundarios del Instituto"))));

                body.Append(new Paragraph(paragraphPropertyNumbering(), new Run(new Text("Metas de liberacion"))));
                mainPart.Document.Save();
             }
        }

        private ParagraphProperties paragraphPropertyNumbering ()
        {
            return new ParagraphProperties(
                        new NumberingProperties(
                            new NumberingLevelReference() { Val = 0 },
                            new NumberingId() { Val = 1 }
                        )
                    );
        }
        private void NumberingPart (MainDocumentPart mainDocumentPart)
        {
            NumberingDefinitionsPart numberingPart =
              mainDocumentPart.AddNewPart<NumberingDefinitionsPart>("defaultNumberingDefinition");

            Numbering element =
              new Numbering(
                new AbstractNum(
                  new Level(
                    new NumberingFormat() { Val = NumberFormatValues.Bullet },
                    new LevelText() { Val = "🎈" }
                  )
                  { LevelIndex = 0 }
                )
                { AbstractNumberId = 1 },
                new NumberingInstance(
                  new AbstractNumId() { Val = 1 }
                )
                { NumberID = 1 });

                element.Save(numberingPart);
        }
    }
}
