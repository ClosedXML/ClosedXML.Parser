using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;

namespace ClosedXML.Parser.Function
{
    public static class ParseFormula
    {
        [FunctionName("parse-formula")]
        public static Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req)
        {
            var refStyle = req.Query["style"] == "R1C1" ? ReferenceStyle.R1C1 : ReferenceStyle.A1;
            string formulaText = req.Query["formula"];

            var serializerSetting = new JsonSerializerSettings
            {
                Formatting = Formatting.Indented,
                Converters = { new AstNodeConverter(refStyle) }
            };

            try
            {
                var nodes = refStyle == ReferenceStyle.A1
                    ? FormulaParser<ScalarValue, AstNode>.CellFormulaA1(formulaText, new F())
                    : FormulaParser<ScalarValue, AstNode>.CellFormulaR1C1(formulaText, new F());
                return Task.FromResult<IActionResult>(new JsonResult(new
                {
                    formula = formulaText,
                    style = refStyle.ToString(),
                    ast = nodes
                }, serializerSetting));
            }
            catch (ParsingException e)
            {
                return Task.FromResult<IActionResult>(new JsonResult(new
                {
                    formula = formulaText,
                    style = refStyle.ToString(),
                    error = e.Message
                }, serializerSetting)
                {
                    StatusCode = StatusCodes.Status422UnprocessableEntity
                });
            }
        }
    }
}
