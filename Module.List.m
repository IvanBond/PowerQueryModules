// Copyright (c) Microsoft Corporation.  All rights reserved.
    
section List;

shared List.NonNullCount = (list as list) as number => List.Count(List.RemoveNulls(list));

shared List.MatchesAll = (list as list, condition as function) as logical => not List.MatchesAny(list, (item) => not condition(item));

shared List.MatchesAny = (list as list, condition as function) as logical => not List.IsEmpty(List.Select(list, condition));

shared List.Range = (list as list, offset as number, optional count as number) as list =>
let
    skippedInput = List.Skip(list, offset)
in
    if (count = null) then skippedInput
    else List.FirstN(skippedInput, count);

shared List.RemoveItems = (list1 as list, list2 as list) as list =>
    List.Select(list1, (x) => not List.Contains(list2, x));

shared List.ReplaceValue = (
    list as list,
    oldValue,
    newValue,
    replacer as function
) as list =>
    let
        f = (val) => try replacer(val, oldValue, newValue) otherwise val
    in
        List.Transform(list, f);

shared List.FindText = (list as list, text as text) as list =>
    List.Select(list, each ContainsTextWithin(_, {text}));

shared List.RemoveLastN = (list as list, optional countOrCondition) as list => 
    List.Reverse(List.Skip(List.Reverse(list), countOrCondition));

shared List.RemoveFirstN = List.Skip;

ContainsTextWithin = (value as any, strings as list) as logical =>
    if (value is text) then
    (
        not List.IsEmpty(List.Select(strings, each Text.Contains(value, _)))
    )
    else if (value is list) then
    (
        not List.IsEmpty(List.Select(value, each @ContainsTextWithin(_, strings)))
    )
    else if (value is record) then
    (
        not List.IsEmpty(List.Select(Table.ToRecords(Record.ToTable(value)), each @ContainsTextWithin([Value], strings)))
    )
    else
    (
        false
    );
