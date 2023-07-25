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
    public IActionResult Parse([FromQuery] string formula)
    {
        const string referenceStyle = "A1";
        try
        {
            var nodes = FormulaParser<ScalarValue, AstNode>.FormulaA1(formula, new F());
            return Json(new FormulaModel
            {
                Formula = formula,
                ReferenceStyle = referenceStyle,
                Ast = nodes
            });
        }
        catch (ParsingException e)
        {
            return UnprocessableEntity(new FormulaModel
            {
                Formula = formula,
                ReferenceStyle = referenceStyle,
                Error = e.Message
            });
        }
    }
}