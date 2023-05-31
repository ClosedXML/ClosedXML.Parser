//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     ANTLR Version: 4.9.2
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

// Generated from c:\\Users\\havli\\source\\repos\\ClosedXML.Parser\\src\\ClosedXML.Parser\\Formula.g4 by ANTLR 4.9.2

// Unreachable code detected
#pragma warning disable 0162
// The variable '...' is assigned but its value is never used
#pragma warning disable 0219
// Missing XML comment for publicly visible type or member '...'
#pragma warning disable 1591
// Ambiguous reference in cref attribute
#pragma warning disable 419

using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using IToken = Antlr4.Runtime.IToken;

/// <summary>
/// This interface defines a complete generic visitor for a parse tree produced
/// by <see cref="FormulaParser"/>.
/// </summary>
/// <typeparam name="Result">The return type of the visit operation.</typeparam>
[System.CodeDom.Compiler.GeneratedCode("ANTLR", "4.9.2")]
[System.CLSCompliant(false)]
public interface IFormulaVisitor<Result> : IParseTreeVisitor<Result> {
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.formula"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitFormula([NotNull] FormulaParser.FormulaContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitExpression([NotNull] FormulaParser.ExpressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.concat_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitConcat_expression([NotNull] FormulaParser.Concat_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.additive_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitAdditive_expression([NotNull] FormulaParser.Additive_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.multiplying_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitMultiplying_expression([NotNull] FormulaParser.Multiplying_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.pow_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitPow_expression([NotNull] FormulaParser.Pow_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.percent_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitPercent_expression([NotNull] FormulaParser.Percent_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.prefix_atom_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitPrefix_atom_expression([NotNull] FormulaParser.Prefix_atom_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.atom_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitAtom_expression([NotNull] FormulaParser.Atom_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.ref_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitRef_expression([NotNull] FormulaParser.Ref_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.ref_intersection_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitRef_intersection_expression([NotNull] FormulaParser.Ref_intersection_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.ref_range_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitRef_range_expression([NotNull] FormulaParser.Ref_range_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.ref_atom_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitRef_atom_expression([NotNull] FormulaParser.Ref_atom_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.constant"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitConstant([NotNull] FormulaParser.ConstantContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.constant_list_rows"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitConstant_list_rows([NotNull] FormulaParser.Constant_list_rowsContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.constant_list_row"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitConstant_list_row([NotNull] FormulaParser.Constant_list_rowContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.postfix_operator"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitPostfix_operator([NotNull] FormulaParser.Postfix_operatorContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.prefix_operator"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitPrefix_operator([NotNull] FormulaParser.Prefix_operatorContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.cell_reference"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitCell_reference([NotNull] FormulaParser.Cell_referenceContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.local_cell_reference"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitLocal_cell_reference([NotNull] FormulaParser.Local_cell_referenceContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.external_cell_reference"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitExternal_cell_reference([NotNull] FormulaParser.External_cell_referenceContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.user_defined_function_call"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitUser_defined_function_call([NotNull] FormulaParser.User_defined_function_callContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.user_defined_function_name"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitUser_defined_function_name([NotNull] FormulaParser.User_defined_function_nameContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.argument_list"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitArgument_list([NotNull] FormulaParser.Argument_listContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.argument"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitArgument([NotNull] FormulaParser.ArgumentContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.arg_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitArg_expression([NotNull] FormulaParser.Arg_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.arg_concat_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitArg_concat_expression([NotNull] FormulaParser.Arg_concat_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.arg_additive_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitArg_additive_expression([NotNull] FormulaParser.Arg_additive_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.arg_multiplying_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitArg_multiplying_expression([NotNull] FormulaParser.Arg_multiplying_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.arg_pow_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitArg_pow_expression([NotNull] FormulaParser.Arg_pow_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.arg_percent_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitArg_percent_expression([NotNull] FormulaParser.Arg_percent_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.arg_prefix_atom_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitArg_prefix_atom_expression([NotNull] FormulaParser.Arg_prefix_atom_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.arg_atom_expression"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitArg_atom_expression([NotNull] FormulaParser.Arg_atom_expressionContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.name_reference"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitName_reference([NotNull] FormulaParser.Name_referenceContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.structure_reference"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitStructure_reference([NotNull] FormulaParser.Structure_referenceContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.table_identifier"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitTable_identifier([NotNull] FormulaParser.Table_identifierContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.table_name"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitTable_name([NotNull] FormulaParser.Table_nameContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.ref_function_call"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitRef_function_call([NotNull] FormulaParser.Ref_function_callContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="FormulaParser.function_call"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitFunction_call([NotNull] FormulaParser.Function_callContext context);
}
