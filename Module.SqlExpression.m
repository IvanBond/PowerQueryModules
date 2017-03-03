// Copyright (c) Microsoft Corporation.  All rights reserved.

section SqlExpression;

shared SqlExpression.SchemaFrom = (schema) =>
let
    NullNativeExpressions = Table.TransformColumns(schema, {{"NativeDefaultExpression", each null}, {"NativeExpression", each null}}),
    RemoveNativeType = Table.RemoveColumns(NullNativeExpressions, {"NativeTypeName"}),
    AddNativeType = Table.AddColumn(RemoveNativeType, "NativeTypeName", each
        if ([Kind] = "text") then
            if ([IsVariableLength] = false) then "nchar"
            else "nvarchar"
        else if ([Kind] = "number") then
            if ([TypeName] = "Decimal.Type") then "decimal"
            else if ([TypeName] = "Currency.Type") then "money"
            else if ([TypeName] = "Int64.Type") then "bigint"
            else if ([TypeName] = "Int32.Type") then "int"
            else if ([TypeName] = "Int16.Type") then "smallint"
            else if ([TypeName] = "Int8.Type") then "tinyint"
            else "double"
        else if ([Kind] = "date") then "date"
        else if ([Kind] = "datetime") then "datetime2"
        else if ([Kind] = "datetimezone") then "datetimeoffset"
        else if ([Kind] = "time") then "time"
        else if ([Kind] = "logical") then "bit"
        else if ([Kind] = "binary") then
            if ([IsVariableLength] = false) then "binary"
            else "varbinary"
        else null),
    Reorder = Table.ReorderColumns(AddNativeType, Table.ColumnNames(schema))
in
    Reorder;

    
