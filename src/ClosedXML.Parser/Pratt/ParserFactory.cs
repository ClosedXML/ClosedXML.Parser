using ClosedXML.Parser.Pratt.Parselets;

namespace ClosedXML.Parser.Pratt;

internal static class ParserFactory
{
    public static Parser<TNode, TContext> Create<TScalar, TNode, TContext>(
        IAstFactory<TScalar, TNode, TContext> factory)
    {
        var parser = new Parser<TNode, TContext>();

        // Register prefix parselets
        parser.Register(TokenType.Number, new NumberParselet<TScalar, TNode, TContext>(factory, parser));
        parser.Register(TokenType.LeftParen, new GroupParselet<TNode, TContext>(parser));
        parser.Register(TokenType.Ident, new IdentParselet<TScalar,TNode,TContext>(factory, parser));

        // Register operation parselets
        parser.Register(TokenType.Plus, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Addition, BindingPower.Addition));
        parser.Register(TokenType.Minus, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Subtraction, BindingPower.Subtraction));
        parser.Register(TokenType.Mul, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Multiplication, BindingPower.Multiplication));
        parser.Register(TokenType.Div, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Division, BindingPower.Division));
        parser.Register(TokenType.Pow, new BinaryOpParselet<TScalar, TNode, TContext>(factory, parser, BinaryOperation.Power, BindingPower.Exponentiation));

        return parser;
    }
}
