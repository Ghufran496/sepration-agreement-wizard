using System;
using System.IO;
using System.Threading.Tasks;
using DocumentGenerationApi.Models;
using DocumentGenerationApi.Services;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DocumentGenerationApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentController : ControllerBase
    {
        private readonly DocumentGenerationService _documentService;
        
        public DocumentController(DocumentGenerationService documentService)
        {
            _documentService = documentService;
        }
        
        [HttpPost("generate")]
        [EnableCors("AllowAngularApp")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public IActionResult GenerateDocument([FromBody] DocumentRequest request)
        {
            try
            {
                if (request == null || request.PartyInfo == null)
                {
                    return BadRequest("Invalid request data");
                }
                
                byte[] documentData = _documentService.GenerateSeparationAgreement(request);
                
                string fileName = $"SeparationAgreement_{DateTime.Now:yyyyMMdd}.docx";
                
                return File(documentData, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, $"Error generating document: {ex.Message}");
            }
        }
    }
} 