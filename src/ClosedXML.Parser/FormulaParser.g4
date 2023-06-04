parser grammar FormulaParser;

options { tokenVocab = FormulaLexer; }

/* ----------------------------------------------------------------------------
 *                               PARSER RULES
 * -------------------------------------------------------------------------- */

formula
        : expression EOF
        ;

/* ---------------------------- Expression ---------------------------- */

expression
        : concat_expression ((GREATER_OR_EQUAL_THAN | LESS_OR_EQUAL_THAN | LESS_THAN | GREATER_THAN | NOT_EQUAL | EQUAL) concat_expression)*
        ;

concat_expression
        : additive_expression (CONCAT additive_expression)*
        ;

additive_expression
        : multiplying_expression ((PLUS | MINUS) multiplying_expression)*
        ;

multiplying_expression
        : pow_expression ((MULT | DIV) pow_expression)*
        ;

pow_expression
        : percent_expression (POW percent_expression)*
        ;

percent_expression
        : prefix_atom_expression PERCENT?
        ;

prefix_atom_expression
        : (PLUS | MINUS) prefix_atom_expression
        | atom_expression
        ;

/*
 * Atom is not LL(1), because there are two possible paths for a OPEN_BRACE:
 * * `OPEN_BRACE expression CLOSE_BRACE`
 * * `OPEN_BRACE ref_expression CLOSE_BRACE` through ref_intersection_expression
 * * Of course, it can be nested, so this is the key pain point.
 * It causes issues for expressions like `SUM((A1:A5) (A2:G4))` that should go
 * through ref expressions, but go through normal expression instead
 * and expression doesn't associate reference operations.
 */
atom_expression
        : constant
        | OPEN_BRACE expression CLOSE_BRACE
        | function_call
        | ref_expression
        ;

/* ------------------------- Reference expression -------------------------- */

ref_expression
        : ref_intersection_expression (COMMA ref_intersection_expression)*
        ;

ref_intersection_expression
        : ref_range_expression (SPACE ref_range_expression)*
        ;

ref_range_expression
        : ref_atom_expression (COLON ref_atom_expression)*
        ;

/*
 * name_reference and structure_reference are not very unambiguous, e.g.
 * BOOK_PREFIX NAME are a prefix for both
 */
ref_atom_expression
        : REF_CONSTANT
        | OPEN_BRACE ref_expression CLOSE_BRACE
        | cell_reference
        | ref_function_call
        | name_reference
        | structure_reference
        ;

/* ------------------------------- Constants ------------------------------- */

/*
 * Constant doesn't contain #REF!, because that is in the ref_expressions and
 * by removing it from here, there isn't an ambiguity whether `constant` or
 * `ref_atom_expression` is the correct path from atom. This is pretty edge case
 * anyway and even Excel has bugs there (e.g. `=(A1 ,#REF!)` can't be parsed despite
 * its validity).
 */
constant
        : NONREF_ERRORS
        | LOGICAL_CONSTANT
        | NUMERICAL_CONSTANT
        | STRING_CONSTANT
        | OPEN_CURLY constant_list_rows CLOSE_CURLY
        ;

constant_list_rows
        : constant_list_row (SEMICOLON constant_list_row)*
        ;

constant_list_row
        : (constant | REF_CONSTANT) (COMMA (constant | REF_CONSTANT))*
        ;

/* ---------------------------- Cell references ---------------------------- */

cell_reference
        : A1_REFERENCE                      // local_cell_reference
        | BANG_REFERENCE                    // external_cell_reference
        | SHEET_RANGE_PREFIX A1_REFERENCE   // external_cell_reference
        | SINGLE_SHEET_REFERENCE            // external_cell_reference
        ;

/* ------------------------------- Arguments ------------------------------- */

// TODO: Argument count should be 0-253, but ANTLR doesn't accept specificic number of repeats
argument_list
        : arg_expression? (COMMA arg_expression?)* CLOSE_BRACE
        ;

/*
 * argExpression is basically a copy of Expression with two important differences:
 * 1. There is no union operator ','. That is because individual arguments are separated by ','
 *    and there would be a confusion whether comma is a union or an arg separator.
 * 2. Once there are braces '('/')', the inner rule is expression or refExpression
 *    because there is no longer any confusion, ',' is a union (i.e. `AREAS((A1,B5))` will
 *    return number 2).
 */
arg_expression
        : arg_concat_expression ((GREATER_OR_EQUAL_THAN | LESS_OR_EQUAL_THAN | LESS_THAN | GREATER_THAN | NOT_EQUAL | EQUAL) arg_concat_expression)*
        ;

arg_concat_expression
        : arg_additive_expression (CONCAT arg_additive_expression)*
        ;

arg_additive_expression
        : arg_multiplying_expression ((PLUS | MINUS) arg_multiplying_expression)*
        ;

arg_multiplying_expression
        : arg_pow_expression ((MULT | DIV) arg_pow_expression)*
        ;

arg_pow_expression
        : arg_percent_expression (POW arg_percent_expression)*
        ;

arg_percent_expression
        : arg_prefix_atom_expression PERCENT?
        ;

arg_prefix_atom_expression
        : (PLUS | MINUS) arg_prefix_atom_expression
        | arg_atom_expression
        ;

/*
 * Official grammar duplicates all ref expression elements, but there is no need
 * to have duplicate rules. It's enough that an argument expression in the first
 * level (=no braces) doesn't have ability to call ',' union operation.
 * Union is at the very top of ref expression hierarchy, so it is enough to
 * use `ref_intersection_expression` in the argument atom here. The nodes below
 * are identical to arg ref nodes to the expression ref nodex.
 */
arg_atom_expression
        : constant
        | OPEN_BRACE expression CLOSE_BRACE
        | function_call
        | ref_intersection_expression
        ;

/*
 * Name reference, where local and external name have been inlined to reduce nesting.
 *
 * external_name originally also had a `bang-name` alternative, but that rule
 * has been omitted, because [MS-XLSX] 2.2.2.1, cell formulas can't use it.
 * Some other usage of formula might, but that is for later.
 */
name_reference
        : NAME                                                                 // local name
        | (SINGLE_SHEET_PREFIX | BOOK_PREFIX) NAME                             // external name
        ;

/* -------------------------- Structure reference -------------------------- */

structure_reference
        : table_identifier? INTRA_TABLE_REFERENCE
        ;

table_identifier
        : BOOK_PREFIX? table_name
        ;

/* table-name is the name of the table the structure reference refers to. If it
 * is missing, the formula containing the structure reference MUST be entered
 * into a cell which belongs to a table and that table's name is used as the
 * table-name. table-name MUST be the value of the displayName attribute of
 * some table element ([ISO/IEC29500-1:2016] section 18.5.1.2). It MUST NOT be
 * any other user-defined name.
 */
table_name
        : NAME
        ;

/* ---------------------------- Function calls ----------------------------- */

ref_function_call
        : REF_FUNCTION_LIST argument_list
        ;

function_call
        : (/*FUNCTION_LIST | FUTURE_FUNCTION_LIST |*/ CELL_FUNCTION_LIST | USER_DEFINED_FUNCTION_NAME) argument_list
        ;