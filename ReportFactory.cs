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
        public void Build(List<Sprint> Sprints){
            var date = DateTime.Now.ToString("yyyyMMddHHmmss");

             using(WordprocessingDocument doc = 
                WordprocessingDocument.Create("ReleasePlan_"+date+".docx",WordprocessingDocumentType.Document))
             {
                 MainDocumentPart mainPart = doc.AddMainDocumentPart();
                Language(mainPart);
                NumberingPart(mainPart);

                 new Document(new Body()).Save(mainPart);
                

                 Body body = mainPart.Document.Body;
                 body.Append(new Paragraph(new Run(new Text("Plan de Liberaciones"))));
                body.Append(new Paragraph(new Run(new Text(""))));
                body.Append(new Paragraph(new Run(new Text("La Dirección General de Tecnologías de la "+
                 "Información les comparte el plan de liberaciones de desarrollo del proyecto: “Sistema de "+
                 "Gestión de Auditorias y Seguimiento”, donde intentamos estimar cuando las funcionalidades "+
                 "nuevas, cambios al sistema o mejoras podrían ser entregados por el equipo de Desarrollo."))));
    
                body.Append(new Paragraph(new Run(new Text("No obstante, las fallas en el sistema no entrarán"+
                 " en esta planificación porque serán tratados con la urgencia que estos ameriten y se corregirán "+
                 "cuanto antes sea posible."))));
        
                body.Append(new Paragraph(new Run(new Text("Todo esto, claro, en función a la capacidad de "+
                 "desarrollo del equipo y respetando los intereses del Instituto Superior de Auditoria y "+
                 "Fiscalización, que serán resguardados por el Director General de Tecnologías de la " + 
                 "Información."))));
          
                body.Append(new Paragraph(new Run(new Text("La idea es tener una guía que refleje las expectativas"+
                 " acerca de lo que se llevará a cabo y cuando se liberara, naturalmente de manera especulativa y "+
                 "por la misma razón, se espera que cambie constantemente."))));
                body.Append(new Paragraph(new Run(new Text(""))));


                body.Append(new Paragraph(new Run(new Text("A continuación, presentamos las metas generales que el"+
                 " equipo de desarrollo pretende alcanzar a mediano plazo referente al “Sistema de Gestión de "+
                 "Auditorias y Seguimiento” y sistemas secundarios del Instituto."))));

                body.Append(new Paragraph(new Run(new Text(""))));

                foreach (var sprint in Sprints)
                {
                    if (sprint.Tickets == null)
                        break;

                    body.Append(new Paragraph(new Run(new Text(sprint.Name + " liberación " ))));
                    body.Append(new Paragraph(new Run(new Text("Estimación: " + sprint.DeadLine))));
                    body.Append(new Paragraph(new Run(new Text("Para esta liberación, incluiremos en el "+ 
                        "sistema funcionalidades como " + string.Join(", ", sprint.Tickets
                        .OrderBy(x => x.Extract).Select(x => x.Extract).Distinct().ToArray())))));

                    Table table = new Table();
                    table.AppendChild<TableProperties>(tableProps());


                    var th = new TableRow();

                    var th_area = new TableCell();
                    th_area.Append(new Paragraph(new Run(new Text("Área"))));

                    var th_card_no = new TableCell();
                    th_card_no.Append(new Paragraph(new Run(new Text("No."))));

                    var th_title = new TableCell();
                    th_title.Append(new Paragraph(new Run(new Text("Historia de Usuario"))));

                    th.Append(th_area);
                    th.Append(th_card_no);
                    th.Append(th_title);

                    table.Append(th);

                    foreach (var ticket in sprint.Tickets)
                    {
                        var tr = new TableRow();

                        var tc_area = new TableCell();
                        tc_area.Append(new Paragraph(new Run(new Text(ticket.Area))));

                        var tc_card_no = new TableCell();
                        tc_card_no.Append(new Paragraph(new Run(new Text(ticket.CardNo.ToString()))));

                        var tc_title = new TableCell();
                        tc_title.Append(new Paragraph(new Run(new Text(ticket.Title))));

                        tr.Append(tc_area);
                        tr.Append(tc_card_no);
                        tr.Append(tc_title);

                        table.Append(tr);
                    }

                    body.Append(table);
                    body.Append(new Paragraph(new Run(new Text(""))));
                    body.Append(new Paragraph(new Run(new Text(""))));
                }

                mainPart.Document.Save();
             }
        }

        public static TableProperties tableProps()
        {
            TableProperties props = new TableProperties(
                new TableStyle { Val = "LightShading-Accent5" },
                new TableBorders(
                new TopBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new BottomBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new LeftBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new RightBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new InsideHorizontalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new InsideVerticalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    
                    Size = 12
                }));

            return props;
        }

        private void Language(MainDocumentPart mainPart)
        {
            DocumentSettingsPart objDocumentSettingPart =
                               mainPart.AddNewPart<DocumentSettingsPart>();
            objDocumentSettingPart.Settings = new Settings();
            Compatibility objCompatibility = new Compatibility();
            CompatibilitySetting objCompatibilitySetting =
                new CompatibilitySetting()
                {
                    Name = CompatSettingNameValues.CompatibilityMode,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "14"
                };
            objCompatibility.Append(objCompatibilitySetting);
            ActiveWritingStyle aws = new ActiveWritingStyle()
            {
                Language = "es-MX"
            };
            objDocumentSettingPart.Settings.Append(objCompatibility);
            objDocumentSettingPart.Settings.Append(aws);

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
