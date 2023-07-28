using System.Text.Json;
using ClosedXML.Parser.Web.Models;
using Microsoft.AspNetCore.Mvc;

namespace ClosedXML.Parser.Web.Controllers;

public class HomeController : Controller
{
    public IActionResult Index()
    {
        return View();
    }

    [HttpGet]
    public IActionResult Parse([FromQuery] string formula, [FromQuery] ReferenceStyle style)
    {
        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            MaxDepth = 128, // Some trees can be rather deep
            Converters = { new AstNodeConverter(style) },
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };
        
        try
        {
            var nodes = style == ReferenceStyle.A1
                ? FormulaParser<ScalarValue, AstNode>.FormulaA1(formula, new F())
                : FormulaParser<ScalarValue, AstNode>.FormulaR1C1(formula, new F());
            return new JsonResult(new FormulaModel
            {
                Formula = formula,
                Style = style,
                Ast = nodes
            }, options);
        }
        catch (ParsingException e)
        {
            return UnprocessableEntity(new FormulaModel
            {
                Formula = formula,
                Style = style,
                Error = e.Message
            });
        }
    }
}