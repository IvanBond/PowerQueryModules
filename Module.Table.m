// Copyright (c) Microsoft Corporation.  All rights reserved.

section Table;

//
// Parameter helpers
//

ColumnsSelector.CreateColumnsList = (value) =>
    if (value is text) then {value}
    else value;

ColumnsSelector.CreateListKeySelector = (value) =>
    if (value is text) then each Record.Field(_, value)
    else each Record.SelectFields(_, value);

TableColumnTransformOperations.CreateTransformOperationsList = (value) =>
    if (value is list and List.Count(value) = 2 and value{0} is text and value{1} is function) then {value}
    else value;

TableEquationCriteria.CreateListEquationCriteria = (value) =>
    if (value is function) then value
    else if (TableEquationCriterion.IsTableEquationCriterion(value)) then TableEquationCriterion.CreateListEquationCriterion(value)
    else if (value is list and not List.IsEmpty(value) and value{0} is list) then List.Transform(value, TableEquationCriterion.CreateListEquationCriterion)
    else ColumnsSelector.CreateListKeySelector(value);

TableEquationCriterion.CreateListEquationCriterion = (value) =>
    {ColumnsSelector.CreateListKeySelector(value{0}), value{1}};

TableEquationCriterion.IsTableEquationCriterion = (value) =>
    value is list and List.Count(value) = 2
        and value{0} is text
        and value{1} is function;    

TableEquationCriterion.Default = (row) => 
    each try Record.SelectFields(_, Record.FieldNames(row)) otherwise _;

TableRowReplacementOperation.CreateListReplacementOperation = (value) =>
    value;

TableRowReplacementOperation.IsTableRowReplacementOperation = (value) =>
    value is list and List.Count(value) = 2
        and value{0} is record
        and value{1} is record;

TableRowReplacementOperations.CreateListReplacementOperations = (value) =>
    if (TableRowReplacementOperation.IsTableRowReplacementOperation(value)) then {TableRowReplacementOperation.CreateListReplacementOperation(value)}
    else List.Transform(value, each TableRowReplacementOperation.CreateListReplacementOperation(_));

//
// Information
//

shared Table.ColumnCount = (table as table) as number =>
    List.Count(Table.ColumnNames(table));

//
// Row selection and access
//

shared Table.AlternateRows = (table as table, offset as number, skip as number, take as number) as table =>
    Table.FromRecords(List.Alternate(Table.ToRecords(table), skip, take, offset), Value.Type(table));

shared Table.InsertRows = (table as table, offset as number, rows as list) as table =>
    Table.ReplaceRows(table, offset, 0, rows);

shared Table.LastN = (table as table, countOrCondition) as table =>
    Table.FromRecords(List.LastN(Table.ToRecords(table), countOrCondition), Value.Type(table));

shared Table.Last = (table as table, optional default) =>
    List.Last(Table.ToRecords(table), default);

shared Table.MatchesAllRows = (table as table, condition as function) as logical =>
    List.MatchesAll(Table.ToRecords(table), condition);

shared Table.MatchesAnyRows = (table as table, condition as function) as logical =>
    List.MatchesAny(Table.ToRecords(table), condition);

shared Table.Partition = (table as table, column as text, groups as number, hash as function) as list => 
    List.Generate(
        () => 0,
        (i) => i < groups,
        (i) => i + 1, 
        (i) => Table.SelectRows(table, (row) => Number.Mod(hash(Record.Field(row, column)), groups) = i));

shared Table.Range = (table as table, offset as number, optional count as nullable number) as table =>
let
    skippedInput = Table.Skip(table, offset)
in
    if (count = null) then skippedInput
    else Table.FirstN(skippedInput, count);

shared Table.RemoveRows = (table as table, offset as number, optional count as nullable number) as table =>
    Table.ReplaceRows(table, offset, if (count = null) then 1 else count, {});

shared Table.Repeat = (table as table, count as number) as table =>
    Table.Combine(List.Repeat({table}, count));

shared Table.ReplaceRows = (table as table, offset as number, count as number, rows as list) as table =>
    Table.FromRecords(List.ReplaceRange(Table.ToRecords(table), offset, count, rows), Type.ReplaceTableKeys(Value.Type(table), {}));

shared Table.ReverseRows = (table as table) as table =>
    Table.FromRecords(List.Reverse(Table.ToRecords(table)), Value.Type(table));

//
// Column selection and access
//

shared Table.HasColumns = (table as table, columns) as logical =>
    List.ContainsAll(Table.ColumnNames(table), ColumnsSelector.CreateColumnsList(columns));

shared Table.PrefixColumns = (table as table, prefix as text) as table =>
let
    names = Table.ColumnNames(table),
    renames = List.Transform(names, each {_, prefix & "." & _})
in
    Table.RenameColumns(table, renames);

shared Table.ColumnsOfType = (table as table, listOfTypes as list) as list =>
let 
    columnsInfo = Record.ToTable(Type.RecordFields(Type.TableRow(Value.Type(table))))
in
    Table.SelectRows(columnsInfo, each
        List.MatchesAny(listOfTypes,
            (typeToCompare) => Type.Is([Value][Type], typeToCompare)))[Name];

//
// Transformation
//

shared Table.AddColumn = (
    table as table,
    newColumnName as text,
    columnGenerator as function,
    optional columnType as nullable type
) as table =>
let
    newColumnType = if (columnType <> null) then columnType else Type.FunctionReturn(Value.Type(columnGenerator))
in
    TableModule!Table.AddColumns(table, {newColumnName}, each {columnGenerator(_)}, {newColumnType});

shared Table.DuplicateColumn = (
    table as table,
    columnName as text,
    newColumnName as text,
    optional columnType as nullable type
) as table =>
let
    newColumnType = if (columnType <> null) then columnType else Type.TableColumn(Value.Type(table), columnName)
in
    TableModule!Table.AddColumns(table, {newColumnName}, each {Record.Field(_, columnName)}, {newColumnType});

shared Table.FillUp = (table as table, columns as list) as table =>
    Table.ReverseRows(Table.FillDown(Table.ReverseRows(table), columns));

shared Table.RemoveLastN = (table as table, optional countOrCondition) as table => 
    Table.ReverseRows(Table.Skip(Table.ReverseRows(table), countOrCondition));

shared Table.RemoveFirstN = Table.Skip;

shared Table.ExpandListColumn = (table as table, column as text) as table =>
    TableModule!Table.ExpandListColumn(table, column, /* singleOrDefault: */ false);

shared Table.ExpandTableColumn = (
    table as table,
    column as text,
    columnNames as list,
    optional newColumnNames as nullable list
) as table =>
let
    newColumnNamesToUse =
        if (newColumnNames <> null and List.Count(newColumnNames) <> List.Count(columnNames)) then
            error [Reason = "Expression.Error", Message = LibraryModule!UICulture.GetString("TableExpandTableColumn_ColumnAndNewColumnNamesMustHaveSameCount")]
        else newColumnNames
in
    Table.ExpandRecordColumn(Table.ExpandListColumn(table, column), column, columnNames, newColumnNamesToUse);

Table.UniqueName = (
    table as table
) =>
    "_" & Text.Combine(Table.ColumnNames(table)); // TODO: More efficient unique name

shared Table.TransformRows = (table as table, transform as function) as list =>
    List.Transform(Table.ToRecords(table), transform);

shared Table.Transpose = (table as table, optional columns) as table =>
    Table.FromColumns(Table.ToRows(table), columns);

shared Table.DemoteHeaders = (table as table) as table =>
    Table.FromRows(List.Combine({{Table.ColumnNames(table)}, Table.ToRows(table)}));

shared Table.ToRows = (table as table) as list =>
    List.Transform(Table.ToRecords(table), (record) => Record.FieldValues(record));

shared Table.ToColumns = (table as table) as list =>
    List.Transform(Table.ColumnNames(table), (column) => Table.Column(table, column));

//
// Membership
//

GetListEquationCriteria = (equationCriteria) => 
    if (equationCriteria <> null) then TableEquationCriteria.CreateListEquationCriteria(equationCriteria)
    else null;

GetTableContainsEquationCriteria = (row, equationCriteria) => 
    if (equationCriteria <> null) then TableEquationCriteria.CreateListEquationCriteria(equationCriteria)
    else TableEquationCriterion.Default(row);

shared Table.Contains = (
    table as table,
    row as record,
    optional equationCriteria
) as logical =>
    List.Contains(Table.ToRecords(table), row, GetTableContainsEquationCriteria(row, equationCriteria));

shared Table.ContainsAll = (
    table as table,
    rows as list,
    optional equationCriteria
) as logical =>
    List.AllTrue(List.Transform(rows, each Table.Contains(table, _, equationCriteria)));

shared Table.ContainsAny = (
    table as table,
    rows as list,
    optional equationCriteria
) as logical =>
    List.AnyTrue(List.Transform(rows, each Table.Contains(table, _, equationCriteria)));

shared Table.IsDistinct = (table as table, optional comparisonCriteria) as logical =>
    List.IsDistinct(Table.ToRecords(table), GetListEquationCriteria(comparisonCriteria));

shared Table.PositionOf = (
    table as table,
    row as record,
    optional occurrence,
    optional equationCriteria
) =>
    List.PositionOf(Table.ToRecords(table), row, occurrence, GetListEquationCriteria(equationCriteria));

shared Table.PositionOfAny = (
    table as table,
    rows as list,
    optional occurrence as number,
    optional equationCriteria
) =>
    List.PositionOfAny(Table.ToRecords(table), rows, occurrence, GetListEquationCriteria(equationCriteria));

shared Table.RemoveMatchingRows = (
    table as table,
    rows as list,
    optional equationCriteria
) as table =>
    Table.FromRecords(List.RemoveMatchingItems(Table.ToRecords(table), rows, GetListEquationCriteria(equationCriteria)), Value.Type(table));

shared Table.ReplaceMatchingRows = (
    table as table,
    replacements as list,
    optional equationCriteria
) as table =>
    Table.FromRecords(List.ReplaceMatchingItems(
        Table.ToRecords(table),
        TableRowReplacementOperations.CreateListReplacementOperations(replacements),
        GetListEquationCriteria(equationCriteria)), Value.Type(table));

//
// Comparison
//

shared Table.Max = (
    table as table,
    comparisonCriteria,
    optional default
) =>
    Table.First(TableModule!Table.SortDescending(table, comparisonCriteria), default);

shared Table.MaxN = (
    table as table,
    comparisonCriteria,
    countOrCondition
) as table =>
    Table.FirstN(TableModule!Table.SortDescending(table, comparisonCriteria), countOrCondition);

shared Table.Min = (
    table as table,
    comparisonCriteria,
    optional default
) =>
    Table.First(Table.Sort(table, comparisonCriteria), default);

shared Table.MinN = (
    table as table,
    comparisonCriteria,
    countOrCondition
) as table =>
    Table.FirstN(Table.Sort(table, comparisonCriteria), countOrCondition);

//
// Other
//

shared Table.Buffer = (table as table) as table =>
    Table.FromRecords(List.Buffer(Table.ToRecords(table)), Value.Type(table));

shared Table.FindText = (table as table, text as text) as table =>
    Table.FromRecords(List.FindText(Table.ToRecords(table), text), Value.Type(table));

shared Replacer.ReplaceValue = (value, old, new) =>
    if (value = old) then new
    else value;

shared Replacer.ReplaceText =
    Text.Replace;

shared Table.ReplaceValue = (
    table as table,
    oldValue,
    newValue,
    replacer as function,
    columnsToSearch as list
) as table =>
    if (oldValue is function or newValue is function) then
        Table.ReplaceValueSlow(table, oldValue, newValue, replacer, columnsToSearch)
    else
        if replacer = Replacer.ReplaceValue then
            Table.TransformColumns(table,
            List.Transform(columnsToSearch, (column) => { column,
                Value.ReplaceType((value) => replacer(value, oldValue, newValue),
                    Type.ForFunction([ReturnType = Type.Union({Value.Type(newValue),
                        Table.SelectRows(Record.ToTable(Type.RecordFields(Type.TableRow(Value.Type(table)))), each [Name] = column){0}[Value][Type]}),
                        Parameters = [value = type any]], 1))}))
        else
            Table.TransformColumns(table,
                List.Transform(columnsToSearch, (column) => { column,
                    Value.ReplaceType((value) => replacer(value, oldValue, newValue),
                        Type.ForFunction([ReturnType = Type.FunctionReturn(Value.Type(replacer)), Parameters = [value = type any]], 1))}));

Table.ReplaceValueSlow = (
    table as table,
    oldValue,
    newValue,
    replacer as function,
    columnsToSearch as list
) as table =>
    let
        getValueAsFunction = (searchValue) =>
            if searchValue is function then
                (row, value) => searchValue(row)
            else
                (row, value) => searchValue,
        oldValue = getValueAsFunction(oldValue),
        newValue = getValueAsFunction(newValue),
        f = (row, value) => try replacer(value, oldValue(row, value), newValue(row, value)) otherwise value,
        pairs = List.Transform(columnsToSearch, each {_, f})
    in
        Table.TransformColumnsWithRowContext(table, pairs);

//
// Temporary helpers to support Table.ReplaceValue
//

Table.TransformColumnsWithRowContext = (
    table as table,
    transformOperations,
    optional missingField as number
) as table =>
let
    transformOperationsList = TableColumnTransformOperations.CreateTransformOperationsList(transformOperations)
in
    Table.FromRecords(Table.TransformRows(table, each Record.TransformFieldsWithRowContext(_, transformOperationsList, missingField)), Table.ColumnNames(table));

Record.TransformFieldsWithRowContext = (row, transformOpList, missingField) =>
    let
        recordAsList = List.Transform(transformOpList, each [Name=_{0},Value=_{1}(row, Record.Field(row, _{0}))])
    in
        row & Record.FromList(recordAsList);

Record.FromList = (list) => Record.FromTable(Table.FromRecords(list));

//
// Temporary helpers to access optimized paths
//
// TODO: Replace use of these helpers with use of the general functions + optimization rules
//

shared Table.IsEmpty = (
    table as table
) as logical
    => List.IsEmpty(Table.ToRecords(table));

shared Table.SplitColumn = (
    table as table,
    sourceColumn as text,
    splitter as function,
    optional columnNamesOrNumber,
    optional default,
    optional extraColumns
) as table =>
    let
        columnNames = Table.ColumnNames(table),
        position = List.PositionOf(columnNames, sourceColumn),
        extra = if (extraColumns = null) then ExtraValues.Ignore else extraColumns
    in
        if (position = -1) then
            error [Reason = "Expression.Error", Message = LibraryModule!UICulture.GetString("ValueException_MissingField", {sourceColumn})]
        else
            let
                uniqueName = Table.UniqueName(table),
                columns =
                    if columnNamesOrNumber = null then
                        let
                            firstRow = table{0}?,
                            firstRowSplitsNumber = if firstRow = null then 0 else List.Count(splitter(Record.Field(firstRow, sourceColumn)))
                        in List.Transform({1..firstRowSplitsNumber}, (x) => sourceColumn & "." & Number.ToText(x))
                    else if columnNamesOrNumber is list then columnNamesOrNumber
                    else if (columnNamesOrNumber is number) and (columnNamesOrNumber >= 0) then List.Transform({1..columnNamesOrNumber}, (x) => sourceColumn & "." & Number.ToText(x))
                    else error [Reason = "Expression.Error", Message = LibraryModule!UICulture.GetString("TableSplitColumnArgumentTypeError")],
                renames = List.Transform(columns, (x) => { uniqueName & x, x }),
                uniqueNames = List.Transform(renames, (x) => x{0}),
                columnType = let tag = Value.Metadata(splitter)[Text.ToText]? in if tag = true then (type nullable text) meta [Serialized.Text=true] else type any,
                columnTypes = List.Transform(columns, each columnType)
            in
                Table.RenameColumns(
                    Table.SelectColumns(
                        TableModule!Table.AddColumns(
                            table,
                            uniqueNames,
                            (x) => LibraryModule!List.Normalize(splitter(Record.Field(x, sourceColumn)), List.Count(columns), default, extra),
                            columnTypes),
                        List.ReplaceRange(columnNames, position, 1, uniqueNames)),
                    renames);

shared Table.CombineColumns = (
    table as table,
    sourceColumns as list,
    combiner as function,
    column as text
) as table =>
    if (List.IsEmpty(sourceColumns)) then
        Table.AddColumn(table, column, (x)=>combiner({}))
    else
        let
            uniqueName = Table.UniqueName(table)
        in
            Table.RenameColumns(
                Table.RemoveColumns(
                    Table.ReorderColumns(
                        Table.AddColumn(
                            table,
                            uniqueName,
                            (x) => combiner(Record.FieldValues(Record.SelectFields(x, sourceColumns))),
                            Type.FunctionReturn(Value.Type(combiner))),
                        {uniqueName, List.Last(sourceColumns)}),
                    sourceColumns),
                {uniqueName, column});

shared Table.FirstValue = (
    table as table,
    optional default as any
) as any =>
    let
        firstRow = Table.First(table),
        firstCells = Record.FieldValues(firstRow)
    in
        if firstRow = null or List.IsEmpty(firstCells)
            then default
        else
            firstCells{0};

//
// Folding
//

IsTrueSelector = (selector as function) as logical =>
let
    expression = RowExpression.From(selector)
in
    expression[Kind] = "Constant" and expression[Value] = true;

GetHandler = (viewHandlers as record, name as text, minArgs as number) as nullable function =>
let
    function = Record.FieldOrDefault(viewHandlers, name),
    invokable = function is function and Type.FunctionRequiredParameters(Value.Type(function)) >= minArgs
in
    if invokable then function else null;

// TODO (Delta): Remove this when Delta.Since is live
Delta.SinceInternal =
let
    realDeltaSince = #shared[Delta.Since]?
in
    if realDeltaSince <> null then realDeltaSince
    else (table, tag) => error "Delta module not loaded.";

//TODO (Action): Remove these when actions are available
TableAction.InsertRowsInternal =
let
    realInsertRows = #shared[TableAction.InsertRows]?
in
    if realInsertRows <> null then realInsertRows
    else (table, updates) => error "Action module not loaded.";

TableAction.UpdateRowsInternal =
let
    realUpdateRows = #shared[TableAction.UpdateRows]?
in
    if realUpdateRows <> null then realUpdateRows
    else (table, rowsToInsert) => error "Action module not loaded.";

TableAction.DeleteRowsInternal =
let
    realDeleteRows = #shared[TableAction.DeleteRows]?
in
    if realDeleteRows <> null then realDeleteRows
    else (table) => error "Action module not loaded.";

ValueAction.NativeStatementInternal =
let
    realNativeStatement = #shared[ValueAction.NativeStatement]?
in
    if realNativeStatement <> null then realNativeStatement
    else (value, statement, optional parameters, optional options) => error "Action module not loaded.";

Handlers.FromTable = (table as nullable table) as record =>
[
    // TODO (DirectQuery): Enable when Value.Expression is live
/*
    GetExpression = () => Value.Expression(table),
*/

    GetRowCount = () => Table.RowCount(table),
    GetRows = () => table,
    GetType = () => Value.Type(table),

    OnAddColumns = (constructors) => List.Accumulate(
        constructors,
        table,
        (state, item) => Table.AddColumn(state, item[Name], item[Function], item[Type])),

    OnDistinct = (columns) => Table.Distinct(table, columns),

    OnGroup = (keys, aggregates) => Table.Group(table, keys, List.Transform(aggregates, each {[Name], [Function], [Type]})),

    OnSelectColumns = (columns) => Table.SelectColumns(table, columns),

    OnSelectRows = (condition) => Table.SelectRows(table, condition),

    OnSkip = (count) => Table.Skip(table, count),

    OnSort = (order) => Table.Sort(table, List.Transform(order, each {[Name], [Order]})),

    OnTake = (count) => Table.FirstN(table, count),

    OnDeltaSince = (tag) => Delta.SinceInternal(table, tag),

    OnNativeQuery = (query, optional parameters, optional options) =>
        Value.NativeQuery(table, query, parameters, options),

    OnInsertRows = (rowsToInsert) => TableAction.InsertRowsInternal(table, rowsToInsert),

    OnUpdateRows = (updates, optional selector) =>
    let
        table = if (selector <> null) then Table.SelectRows(table, selector) else table
    in
        TableAction.UpdateRowsInternal(table, List.Transform(updates, each {[Name], [Function]})),

    OnDeleteRows = (optional selector) =>
    let
        table = if (selector <> null) then Table.SelectRows(table, selector) else table
    in
        TableAction.DeleteRowsInternal(table),

    OnNativeStatement = (statement, optional parameters, optional options) =>
        ValueAction.NativeStatementInternal(table, statement, parameters, options)
];

Handlers.AddDefaults = (handlers as record, tableDefault as function, actionDefault as function) => handlers &
[
    OnAddColumns = (constructors) => tableDefault(handlers[OnAddColumns](constructors)),
    OnDistinct = (columns) => tableDefault(handlers[OnDistinct](columns)),
    OnGroup = (keys, aggregates) => tableDefault(handlers[OnGroup](keys, aggregates)),
    OnSelectColumns = (columns) => tableDefault(handlers[OnSelectColumns](columns)),
    OnSelectRows = (selector) => tableDefault(handlers[OnSelectRows](selector)),
    OnSkip = (count) => tableDefault(handlers[OnSkip](count)),
    OnSort = (order) => tableDefault(handlers[OnSort](order)),
    OnTake = (count) => tableDefault(handlers[OnTake](count)),

    OnInsertRows = (rowsToInsert) => actionDefault(handlers[OnInsertRows](rowsToInsert)),
    OnUpdateRows = (updates, optional selector) => actionDefault(handlers[OnUpdateRows](updates, selector)),
    OnDeleteRows = (optional selector) => actionDefault(handlers[OnDeleteRows](selector))
];

shared Table.View = (
    table as nullable table,
    handlers as record
) as table =>
    let
        tableDefault = if (handlers[TableDefault]? <> null) then handlers[TableDefault]? else (table) => table,
        actionDefault = if (handlers[ActionDefault]? <> null) then handlers[ActionDefault]? else (action) => action,
        defaultHandlers = if (table <> null) then Handlers.FromTable(table) else null,
        defaultHandlersWithKind = if (defaultHandlers <> null) then Handlers.AddDefaults(defaultHandlers, tableDefault, actionDefault) else [],
        defaultHandlersWithoutExpression = defaultHandlersWithKind & [GetExpression = () => null],
        handlersWithoutTableAndKind = Record.RemoveFields(handlers, {"TableDefault", "ActionDefault"}, MissingField.Ignore),
        viewHandlers = defaultHandlersWithKind & handlersWithoutTableAndKind,
        view = TableModule!Table.FromHandlers(viewHandlers),
        accumulatingSelectRowsHandler =
        [
            OnSelectRows = (selector) =>
            let
                createTable = (accumulatedSelector) => let
                    this = Table.View(null,
                    [
                        GetRows = () => Table.SelectRows(view, accumulatedSelector),
                        GetRowCount = () => Table.RowCount(GetRows()),
                        GetType = () => Value.Type(GetRows()),

                        OnAddColumns = (constructors) => List.Accumulate(
                            constructors,
                            GetRows(),
                            (state, item) => Table.AddColumn(state, item[Name], item[Function], item[Type])),
                        OnDistinct = (columns) => Table.Distinct(GetRows(), columns),
                        OnGroup = (keys, aggregates) => Table.Group(GetRows(), keys, List.Transform(aggregates, each {[Name], [Function], [Type]})),
                        OnSelectColumns = (columns) => Table.SelectColumns(GetRows(), columns),
                        OnSkip = (count) => Table.Skip(GetRows(), count),
                        OnSort = (order) => Table.Sort(GetRows(), List.Transform(order, each {[Name], [Order]})),
                        OnTake = (count) => Table.FirstN(GetRows(), count),

                        OnDeltaSince = (tag) => Delta.SinceInternal(GetRows(), tag),
                        OnNativeQuery = (query, optional parameters, optional options) =>
                            Value.NativeQuery(GetRows(), query, parameters, options),

                        OnInsertRows = (rowsToInsert) => TableAction.InsertRowsInternal(view, rowsToInsert),
                        OnUpdateRows = (updates, optional selector) =>
                            if (selector = null or IsTrueSelector(selector)) then
                                let
                                    function = GetHandler(viewHandlers, "OnUpdateRows", 2)
                                in
                                    if function <> null then function(updates, accumulatedSelector)
                                    else TableAction.UpdateRowsInternal(GetRows(), updates)
                            else TableAction.UpdateRowsInternal(Table.SelectRows(@this, selector), updates),
                        OnDeleteRows = (optional selector) =>
                            if (selector = null or IsTrueSelector(selector)) then
                                let
                                    function = GetHandler(viewHandlers, "OnDeleteRows", 1)
                                in
                                    if function <> null then function(accumulatedSelector)
                                    else TableAction.DeleteRowsInternal(GetRows())
                            else TableAction.DeleteRowsInternal(Table.SelectRows(@this, selector)),
                        OnSelectRows = (selector) =>
                            if (selector = null or IsTrueSelector(selector)) then @this
                            else @createTable(each accumulatedSelector(_) and selector(_)),
                        OnNativeStatement = (statement, optional parameters, optional options) =>
                            ValueAction.NativeStatementInternal(GetRows(), statement, parameters, options)
                    ])
            in
                this
            in
                createTable(selector)
        ],
        accumulatingSelectRowsView =
            if (handlers[OnSelectRows]? = null)
                then TableModule!Table.FromHandlers(Handlers.FromTable(view) & accumulatingSelectRowsHandler)
                else view
    in
        accumulatingSelectRowsView;

//
// Information and statistics
//

shared Table.Schema = (table as table) as table =>
    Type.TableSchema(Value.Type(Table.FirstN(table, 0)));

shared Type.TableSchema = (tableType as type) as table =>
    let
        GetTypeName = (t) =>
            if (Type.IsNullable(t) and not Type.Is(t, Null.Type)) then @GetTypeName(Type.NonNullable(t))
            else LibraryModule!Type.Name(t),
        GetTypeKind = (t) =>
            if (Type.IsNullable(t) and not Type.Is(t, Null.Type)) then @GetTypeKind(Type.NonNullable(t))
            else LibraryModule!Type.Kind(t),

        rowType = Type.TableRow(tableType),
        columns = Record.ToTable(Type.RecordFields(rowType)),
        columns1 = Table.AddIndexColumn(columns, "Position"),
        columnTypes1 = Table.ExpandRecordColumn(columns1, "Value", {"Type"}, {"TypeValue"}),
        columnTypes2 = Table.AddColumn(columnTypes1, "TypeName", each GetTypeName([TypeValue]), type text),
        columnTypes3 = Table.AddColumn(columnTypes2, "Kind", each GetTypeKind([TypeValue]), type text),
        columnTypes4 = Table.AddColumn(columnTypes3, "IsNullable", each Type.IsNullable([TypeValue]), type logical),
        reordered = Table.ReorderColumns(columnTypes4, {"Name", "Position", "TypeName", "Kind", "IsNullable", "TypeValue"}),
        extendedInfo1 = Table.AddColumn(reordered, "Facets", each Type.Facets([TypeValue]), Type.FunctionReturn(Value.Type(Type.Facets))),
        extendedInfo2 = Table.AddColumn(extendedInfo1, "Documentation", each Value.Metadata([TypeValue])),
        extendedInfo3 = Table.RemoveColumns(extendedInfo2, {"TypeValue"}),
        extendedInfo4 = Table.ExpandRecordColumn(
            extendedInfo3,
            "Facets",
            {
                "NumericPrecisionBase",
                "NumericPrecision",
                "NumericScale",
                "DateTimePrecision",
                "MaxLength",
                "IsVariableLength",
                "NativeTypeName",
                "NativeDefaultExpression",
                "NativeExpression"
            }),
        extendedInfo5 = Table.ExpandRecordColumn(
            extendedInfo4,
            "Documentation",
            {
                "Documentation.FieldDescription",
                "Documentation.IsWritable"
            },
            {
                "Description",
                "IsWritable"
            }),
        typedInfo = Table.TransformColumnTypes(
            extendedInfo5,
            {
                {"Description", type nullable text},
                {"IsWritable", type nullable logical}
            })
    in
        typedInfo;

shared Table.Profile = (table as table) as table =>
    let
        Type.IsAny = (t, types) => not List.IsEmpty(List.Select(types,  each Type.Is(t, _))),

        scalar = {type null, type logical, type number, type text, type date, type datetime, type datetimezone, type time, type duration},
        subtractable = {type null, type number, type date, type datetime, type datetimezone, type duration},
        multiplicable = {type null, type number},

        isScalar = (t) => Type.IsAny(t, scalar) or Type.IsNullable(t) and @isScalar(Type.NonNullable(t)),
        canSubtract = (t) => Type.IsAny(t, subtractable) or Type.IsNullable(t) and @canSubtract(Type.NonNullable(t)),
        canMultiply = (t) => Type.IsAny(t, multiplicable) or Type.IsNullable(t) and @canMultiply(Type.NonNullable(t)),

        // TODO: Needed to pass through conversion to QueryExpression; review.
        equalsNull = each _ = null,
        aggregateFunctions =
        {
            [Name = "Min", CanApply = (t) => isScalar(t), Function = List.Min],
            [Name = "Max", CanApply = (t) => isScalar(t), Function = List.Max],
            [Name = "Average", CanApply = (t) => isScalar(t) and canSubtract(t), Function = List.Average],
            [Name = "StandardDeviation", CanApply = (t) => isScalar(t) and canSubtract(t) and canMultiply(t), Function = List.StandardDeviation],
            [Name = "Count", CanApply = (t) => true, Function = List.Count],
            [Name = "NullCount", CanApply = (t) => true, Function = (values) => List.Count(List.Select(values, equalsNull))],
            [Name = "DistinctCount", CanApply = (t) => isScalar(t), Function = (values) => List.Count(List.Distinct(values))]
        },
        columnNames = Table.ColumnNames(table),

        createView = (columnNames, aggregateFunctions) =>
            let
                aggregates = List.Transform(columnNames,
                    (columnName) => List.Transform(
                        List.Select(aggregateFunctions, each [CanApply](Type.TableColumn(Value.Type(table), columnName))),
                        (entry) =>
                        {
                            entry[Name] & "/" & columnName,
                            (rows) => entry[Function](Table.Column(rows, columnName))
                        })),
                aggregates2 = List.Combine(aggregates),

                groupProfile = Table.Buffer(Table.Group(table, {}, aggregates2)),
                groupProfile2 = Record.ToTable(groupProfile{0}),

                // TODO bug 6019085: Workaround for Table.Group returning 0 rows over empty tables
                rawProfile = Table.FromRecords(List.Transform(aggregates2, each [Name = _{0}, Value = _{1}(table)])),
                profile = if (not Table.IsEmpty(groupProfile)) then groupProfile2 else rawProfile,

                profile2 = Table.SplitColumn(profile, "Name", Splitter.SplitTextByEachDelimiter({"/"}, QuoteStyle.Csv, false), {"Function", "Column"}),
                profileTable = Table.Pivot(profile2, List.Distinct(profile2[Function]), "Function", "Value"),
                profileTable2 = Table.SelectColumns(profileTable, {"Column"} & List.Transform(aggregateFunctions, each [Name]), MissingField.UseNull),
                profileView = Table.View(null,
                [
                    GetType = () =>
                        let
                            fields = List.Transform(
                                aggregateFunctions,
                                each {[Name], [Type = Type.FunctionReturn(Value.Type([Function])), Optional = false]}),
                            fieldsTable = #table({"Name", "Value"}, {{"Column", [Type = Text.Type, Optional = false]}} & fields),
                            rowType = Type.ForRecord(Record.FromTable(fieldsTable), false)
                        in
                            TableModule!Type.ForTable(rowType),

                    GetRows = () => profileTable2,

                    OnSelectColumns = (columns) =>
                        if (List.Contains(columns, "Column")) then @createView(columnNames, List.Select(aggregateFunctions, each List.Contains(columns, [Name])))
                        else error "The 'Column' column must appear in the selection.",

                    OnSelectRows = (selector) =>
                        let
                            columns = #table({"Column"}, List.Transform(columnNames, each {_})),
                            filteredColumns = Table.SelectRows(columns, selector),
                            filteredColumnNames = filteredColumns[Column]
                        in
                            @createView(filteredColumnNames, aggregateFunctions)
                ])
            in
                profileView
    in
        createView(columnNames, aggregateFunctions);

    
