using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace LetterGeneratorApp.Services
{
    public class TemplateInitializer
    {
        public static void EnsureTemplateExists(string templatePath)
        {
            string directory = Path.GetDirectoryName(templatePath);
            
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Если шаблон уже существует и валиден, не создаём новый
            if (File.Exists(templatePath))
            {
                try
                {
                    using (var doc = WordprocessingDocument.Open(templatePath, false))
                    {
                        // Если документ откроется без ошибок, шаблон валиден
                        return;
                    }
                }
                catch
                {
                    // Если файл повреждён, удаляем его
                    File.Delete(templatePath);
                }
            }

            // Создать новый валидный шаблон
            CreateNewTemplate(templatePath);
        }

        private static void CreateNewTemplate(string templatePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Заголовок письма
                var heading = body.AppendChild(new Paragraph());
                var headingRun = heading.AppendChild(new Run());
                headingRun.AppendChild(new Text { Text = "{LETTER_SUBJECT}" });
                var headingProps = heading.PrependChild(new ParagraphProperties());
                headingProps.AppendChild(new ParagraphStyleId { Val = "Heading1" });

                // Пустая строка
                body.AppendChild(new Paragraph());

                // Дата
                var dateParagraph = body.AppendChild(new Paragraph());
                var dateRun = dateParagraph.AppendChild(new Run());
                dateRun.AppendChild(new Text { Text = "{LETTER_DATE}" });

                // Пустая строка
                body.AppendChild(new Paragraph());

                // Приветствие
                var greetingParagraph = body.AppendChild(new Paragraph());
                var greetingRun = greetingParagraph.AppendChild(new Run());
                greetingRun.AppendChild(new Text { Text = "Уважаемый(ая) {RECIPIENT_NAME}!" });

                // Пустая строка
                body.AppendChild(new Paragraph());

                // Основной текст
                var bodyParagraph = body.AppendChild(new Paragraph());
                var bodyRun = bodyParagraph.AppendChild(new Run());
                bodyRun.AppendChild(new Text { Text = "{LETTER_BODY}" });

                // Пустая строка
                body.AppendChild(new Paragraph());

                // Подпись
                var signatureParagraph = body.AppendChild(new Paragraph());
                var signatureRun = signatureParagraph.AppendChild(new Run());
                signatureRun.AppendChild(new Text { Text = "С уважением," });

                // Пустая строка
                body.AppendChild(new Paragraph());

                // ФИО подписывающего
                var senderNameParagraph = body.AppendChild(new Paragraph());
                var senderNameRun = senderNameParagraph.AppendChild(new Run());
                senderNameRun.AppendChild(new Text { Text = "{SENDER_NAME}" });

                // Должность подписывающего
                var senderPostParagraph = body.AppendChild(new Paragraph());
                var senderPostRun = senderPostParagraph.AppendChild(new Run());
                senderPostRun.AppendChild(new Text { Text = "{SENDER_POST}" });

                mainPart.Document.Save();
            }
        }
    }
}
