using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ProjectAPI.Dtos;
using ProjectAPI.Services;

namespace ProjectAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class QuestionController : ControllerBase
    {

        private readonly IQuestionService questionService;

        public QuestionController(IQuestionService questionService)
        {
            this.questionService = questionService;
        }

       
        [HttpPut]
        public async Task<IActionResult> AddQuestionsInBulk(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                    return BadRequest("No file uploaded");
                await questionService.AddQuestionsInBulk(file);
                return Ok();
            }
            catch (Exception ex)
            {
                var message = new { Message = ex.Message };
                return NotFound(message);
            }
        }

        [HttpPost("download-pdf")]
        public async Task<IActionResult> DownloadPdf([FromBody] List<int> questionsIds)
        {
            try
            {
                var pdfBytes = await questionService.GeneratePdf(questionsIds);
                return File(pdfBytes, "application/pdf", "SelectedQuestions.pdf");
            }
            //catch(iText.Kernel.PdfException pdfEx)
            //{
            //    return BadRequest(new {message = pdfEx.Message });
            //}
            catch(Exception ex)
            {
                return BadRequest(new {message = ex.Message});
            }
        }
    }
}
