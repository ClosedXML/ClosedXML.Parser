using static ClosedXML.Parser.ReferenceAxisType;

namespace ClosedXML.Parser.Tests.Rules;

public class ArgumentListRuleTests
{
    [Theory]
    [MemberData(nameof(TestCases))]
    public void Argument_list_can_have_zero_or_more_arguments(string formula, AstNode expectedNode)
    {
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Theory]
    [InlineData("FUN(1,(A1,two),3)")]
    [InlineData("FUN(1,((A1,two)),3)")]
    [InlineData("FUN( 1 , ( ( A1 , two ) ) , 3 ) ")]
    public void Argument_list_interpreters_comma_as_argument_separator_but_nested_expression_interprets_comma_as_range_union_operator(string formula)
    {
        var expectedNode = new FunctionNode("FUN")
        {
            Children = new AstNode[]
            {
                new ValueNode(1),
                new BinaryNode(
                    BinaryOperation.Union,
                    new ReferenceNode(new ReferenceArea(Relative, 1, Relative, 1)),
                    new NameNode("two")),
                new ValueNode(3)
            }
        };
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    public static IEnumerable<object[]> TestCases
    {
        get
        {
            // argument_list : CLOSE_BRACE
            yield return new object[]
            {
                "FUN()",
                new FunctionNode("FUN") { Children = Array.Empty<AstNode>() }
            };

            // argument_list : arg_expression CLOSE_BRACE
            yield return new object[]
            {
                "FUN(TRUE)",
                new FunctionNode("FUN") { Children = new AstNode[] { new ValueNode("Logical", true)} }
            };
            yield return new object[]
            {
                "FUN(1.5)",
                new FunctionNode("FUN") { Children = new AstNode[] { new ValueNode("Number", 1.5)} }
            };

            // argument_list : arg_expression CLOSE_BRACE
            // arg_expression is not a value directly, but another node
            yield return new object[]
            {
                "FUN(100%)",
                new FunctionNode("FUN")
                {
                    Children = new AstNode[]
                    {
                        new UnaryNode(UnaryOperation.Percent)
                        {
                            Children = new AstNode[]
                            {
                                new ValueNode("Number", 100.0)
                            }
                        }
                    }
                }
            };

            // argument_list : COMMA CLOSE_BRACE
            yield return new object[]
            {
                "FUN(,)",
                new FunctionNode("FUN") { Children = new AstNode[] { new ValueNode("Blank", string.Empty), new ValueNode("Blank", string.Empty) } }
            };

            // argument_list : COMMA COMMA CLOSE_BRACE
            yield return new object[]
            {
                "FUN(  ,  , )",
                new FunctionNode("FUN") { Children = new AstNode[] { new ValueNode("Blank", string.Empty), new ValueNode("Blank", string.Empty), new ValueNode("Blank", string.Empty) } }
            };

            // argument_list : arg_expression COMMA COMMA arg_expression CLOSE_BRACE
            yield return new object[]
            {
                "FUN(  TRUE ,  , 1.0 )",
                new FunctionNode("FUN")
                {
                    Children = new AstNode[]
                    {
                        new ValueNode("Logical", true),
                        new ValueNode("Blank", string.Empty),
                        new ValueNode("Number", 1.0)
                    }
                }
            };

            // argument_list : COMMA arg_expression COMMA CLOSE_BRACE
            yield return new object[]
            {
                "FUN(   , TRUE , )",
                new FunctionNode("FUN")
                {
                    Children = new AstNode[]
                    {
                        new ValueNode("Blank", string.Empty),
                        new ValueNode("Logical", true),
                        new ValueNode("Blank", string.Empty),
                    }
                }
            };

            // argument_list : arg_expression COMMA arg_expression COMMA arg_expression CLOSE_BRACE
            yield return new object[]
            {
                "FUN( FALSE  , TRUE , 1 )",
                new FunctionNode("FUN")
                {
                    Children = new AstNode[]
                    {
                        new ValueNode("Logical", false),
                        new ValueNode("Logical", true),
                        new ValueNode("Number", 1.0),
                    }
                }
            };
        }
    }
}