using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using ProjectAPI.Data;
using ProjectAPI.Dtos;
using System.Drawing;
using System.Drawing.Imaging;
using System.Security.Cryptography;
using System.Text;
using Image = System.Drawing.Image;
using Question = ProjectAPI.Dtos.Question;

namespace ProjectAPI.Services
{
    public interface IQuestionService
    {
        public Task AddQuestionsInBulk(IFormFile file);
        Task<byte[]> GeneratePdf(List<int> questionIds);
    }

    public class QuestionService : IQuestionService
    {
        private readonly QuestionBankContext _context;
        private static Dictionary<string, string> _imageHashCache = new Dictionary<string, string>();

        public QuestionService(QuestionBankContext context)
        {
            _context = context;
        }

        public Task AddQuestionsInBulk(IFormFile file)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var questions = new List<Data.Question>();
            var imagePath = "C:\\Users\\6147954\\source\\repos\\ProjectAPI\\Images\\";

            if (!Directory.Exists(imagePath))
                Directory.CreateDirectory(imagePath);

            using (var package = new ExcelPackage(file.OpenReadStream()))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    var question = new Data.Question
                    {
                        Subject = ProcessOption(worksheet.Cells[row, 1], imagePath),
                        Topic = ProcessOption(worksheet.Cells[row, 2], imagePath),
                        DifficultyLevel = ProcessOption(worksheet.Cells[row, 3], imagePath),
                        QuestionText = ProcessOption(worksheet.Cells[row, 4], imagePath),
                        OptionA = ProcessOption(worksheet.Cells[row, 5], imagePath),
                        OptionB = ProcessOption(worksheet.Cells[row, 6], imagePath),
                        OptionC = ProcessOption(worksheet.Cells[row, 7], imagePath),
                        OptionD = ProcessOption(worksheet.Cells[row, 8], imagePath),
                        CorrectAnswer = ProcessOption(worksheet.Cells[row, 9], imagePath),
                        CreatedBy = 2,
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now,
                    };
                    _context.Add(question);
                    _context.SaveChanges();
                }
            }

            return Task.CompletedTask;
        }

        private string ProcessOption(ExcelRange cell, string imagePath)
        {
            if (cell.Value is string)
            {
                return cell.Text;
            }
            else
            {
                // Get the row and column of the cell
                int row = cell.Start.Row;
                int column = cell.Start.Column;

                // Check if the worksheet contains any drawings (images)
                var worksheet = cell.Worksheet;
                if (worksheet.Drawings.Count > 0)
                {
                    foreach (var drawing in worksheet.Drawings)
                    {
                        if (drawing is ExcelPicture picture && picture.From.Row == row - 1 && picture.From.Column == column - 1)
                        {
                            // Generate a unique hash for the image
                            string imageHash = GetImageHash(picture.Image);

                            // Check if the image hash already exists in the cache
                            if (_imageHashCache.TryGetValue(imageHash, out string existingImagePath))
                            {
                                return existingImagePath;  // Return existing image path
                            }

                            // Assuming picture.Image is a System.Drawing.Image object
                            ImageFormat imageFormat = ImageFormat.Png; // Default to PNG

                            // Get the extension based on the image format
                            string imageExtension = ".png"; // Default to PNG extension
                            if (picture.Image.RawFormat.Equals(ImageFormat.Jpeg))
                            {
                                imageFormat = ImageFormat.Jpeg;
                                imageExtension = ".jpg";
                            }
                            else if (picture.Image.RawFormat.Equals(ImageFormat.Gif))
                            {
                                imageFormat = ImageFormat.Gif;
                                imageExtension = ".gif";
                            }
                            else if (picture.Image.RawFormat.Equals(ImageFormat.Bmp))
                            {
                                imageFormat = ImageFormat.Bmp;
                                imageExtension = ".bmp";
                            }

                            // Generate a unique file name for the image
                            var imageName = Guid.NewGuid().ToString() + imageExtension;
                            var imageFullPath = Path.Combine(imagePath, imageName);

                            // Save the image using the System.Drawing.Image's Save method
                            picture.Image.Save(imageFullPath, imageFormat);

                            // Cache the new image path
                            var relativeImagePath = Path.Combine("Images", imageName);
                            _imageHashCache[imageHash] = relativeImagePath;

                            return relativeImagePath;
                        }
                    }
                }

                // If no image is found in the specified cell, return the cell's text value
                return cell.Text;
            }
        }


        private string GetImageHash(Image image)
        {
            using (var ms = new MemoryStream())
            {
                image.Save(ms, ImageFormat.Png); // Save image to memory stream
                var imageBytes = ms.ToArray(); // Convert to byte array

                using (var sha256 = SHA256.Create())
                {
                    var hashBytes = sha256.ComputeHash(imageBytes); // Generate hash
                    return Convert.ToBase64String(hashBytes); // Return hash as string
                }
            }
        }

        public async Task<byte[]> GeneratePdf(List<int> questionIds)
 {
     // Retrieve questions from database
    
     try
     {
         var questions = await _context.Questions
                                  .Where(q => questionIds.Contains(q.QuestionId))
                                  .ToListAsync();

         if (questions == null || questions.Count == 0)
         {
             throw new ArgumentException("No questions found for the provided IDs.");
         }

         using (var ms = new MemoryStream())
         {
             PdfWriter writer = new PdfWriter(ms);
             PdfDocument pdf = new PdfDocument(writer);
             Document document = new Document(pdf);

             // Add questions to the PDF
             foreach (var question in questions)
             {
                 AddContentToPdf(document, "Question", question.QuestionText);
                 AddContentToPdf(document, "A", question.OptionA);
                 AddContentToPdf(document, "B", question.OptionB);
                 AddContentToPdf(document, "C", question.OptionC);
                 AddContentToPdf(document, "D", question.OptionD);
                 document.Add(new Paragraph(" ")); // Empty line
             }

             document.Close();

             return ms.ToArray();
         }

     }
     catch (Exception pdfEx)
     {

         throw new Exception(pdfEx.Message);
     }
 }

 private void AddContentToPdf(Document document, string optionLabel, string content)
 {
     if (string.IsNullOrEmpty(content))
     {
         document.Add(new Paragraph($"{optionLabel}: N/A"));
     }
     else if (IsImagePath(content))
     {
         // Assuming the content is a relative image path
         var imagePath = Path.Combine("C:\\Users\\6147954\\source\\repos\\ProjectAPI\\", content);

         if (File.Exists(imagePath))
         {
             ImageData imageData = ImageDataFactory.Create(imagePath);
             iText.Layout.Element.Image image = new iText.Layout.Element.Image(imageData);
             float imageHeightInPoints = 10 * 12f; // Approx. 12 points per line, so 15 lines = 15 * 12 = 180 points
             float imageAspectRatio = image.GetImageWidth() / image.GetImageHeight(); // Width/Height
             float imageWidthInPoints = imageHeightInPoints * imageAspectRatio;

             image.ScaleToFit(imageWidthInPoints, imageHeightInPoints);
             document.Add(new Paragraph($"{optionLabel}:"));
             document.Add(image);
         }
         else
         {
             document.Add(new Paragraph($"{optionLabel}: [Image not found]"));
         }
     }
     else
     {
         // Assuming the content is plain text
         document.Add(new Paragraph($"{optionLabel}: {content}"));
     }
 }
 private bool IsImagePath(string content)
 {
     // Check if the content is a relative path (e.g., "Images\file.png" or "Images/file.jpg")
     return content.StartsWith("Images\\") || content.StartsWith("Images/");
 }
    }
}




