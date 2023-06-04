using Antlr4.Runtime;

using var file = File.OpenRead(@"c:\Temp\formulas.txt");
using var reader = new StreamReader(file);

var goodCount = 0;
var badCount = 0;
var total = 0;
var sw = System.Diagnostics.Stopwatch.StartNew();
for (var line = reader.ReadLine(); line != null; line = reader.ReadLine())
{
    total++;
    line = line.Substring(0, line.Length - 1);
    line = line.Substring(1);
    line = line.Replace("\"\"", "\"");
    var inputStream = new AntlrInputStream(line.ToString());
    var speakLexer = new FormulaLexer(inputStream);
    speakLexer.RemoveErrorListeners();
    var commonTokenStream = new CommonTokenStream(speakLexer);
    var speakParser = new FormulaParser(commonTokenStream, TextWriter.Null, TextWriter.Null)
    {
        Interpreter =
        {
            PredictionMode = Antlr4.Runtime.Atn.PredictionMode.SLL
        }
    };
    var res = speakParser.formula();
    if (res.exception is not null)
    {
        badCount++;
        Console.WriteLine("ERROR                {0}   {1}", line, res.exception.Message);
    }
    else
    {
        goodCount++;
    }
}

sw.Stop();
Console.WriteLine($"Total: {total}\nGoodCount: {goodCount}\nBadCount: {badCount}\n\nElapsed {sw.ElapsedMilliseconds} ms");
Console.WriteLine("\nPress any key...\n");
Console.ReadKey();

