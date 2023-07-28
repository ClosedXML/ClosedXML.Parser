using ClosedXML.Parser.Web.Models;
using Microsoft.AspNetCore.Mvc;

namespace ClosedXML.Parser.Web.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;

    public HomeController(ILogger<HomeController> logger)
    {
        _logger = logger;
    }

    public IActionResult Index()
    {
        return View();
    }

    [HttpGet]
    public IActionResult Parse([FromQuery] string formula, [FromQuery] string mode)
    {
        var isA1 = mode is "A1" or not "R1C1";
        var sanitizedMode = isA1 ? "A1" : "R1C1";
        try
        {
            var nodes = isA1 
                ? FormulaParser<ScalarValue, AstNode>.FormulaA1(formula, new F())
                : FormulaParser<ScalarValue, AstNode>.FormulaR1C1(formula, new F());
            return Json(new FormulaModel
            {
                Formula = formula,
                Mode = sanitizedMode,
                Ast = nodes
            });
        }
        catch (ParsingException e)
        {
            return UnprocessableEntity(new FormulaModel
            {
                Formula = formula,
                Mode = sanitizedMode,
                Error = e.Message
            });
        }
    }
}