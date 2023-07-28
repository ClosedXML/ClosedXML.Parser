﻿namespace ClosedXML.Parser;

public record UnaryNode(UnaryOperation Operation) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return Operation switch
        {
            UnaryOperation.Percent => "%",
            UnaryOperation.Minus => "-",
            UnaryOperation.Plus => "+",
            _ => throw new NotSupportedException()
        };
    }
};