// Copyright (c) Microsoft Corporation.  All rights reserved.
    
section Text;

shared Text.Format = (formatString as text, arguments, optional culture as text) as text =>
    let
        Placeholders = Text.FormatGetPlaceholders(formatString),
        Result =
            if Value.Is(arguments, type list) or Value.Is(arguments, type record) then
                List.Accumulate(Placeholders, [ Text = formatString, Offset = 0 ], MakeReplacement)[Text]
            else
                error Error.Record("Expression.Error", LibraryModule!UICulture.GetString("TextFormat_InvalidArguments"), Value.Type(arguments)),
        MakeReplacement = (state as record, placeholder as record) as record =>
            let
                Reference = placeholder[Reference],
                Argument =
                    try
                        if Value.Is(Reference, type number) then arguments{Reference}
                        else Record.Field(arguments, Reference)
                    otherwise
                        error Error.Record("Expression.Error", LibraryModule!UICulture.GetString("TextFormat_InvalidReference"), placeholder),
                ReplacementText = Text.FormatToText(Argument, culture),
                NewText = Text.ReplaceRange(state[Text], placeholder[Offset] + state[Offset], placeholder[Length], ReplacementText),
                NewOffset = state[Offset] + (Text.Length(ReplacementText) - placeholder[Length])
            in
                [ Text = NewText, Offset = NewOffset ]
    in
        Result;

Text.FormatGetPlaceholders = (formatString as text) as list =>
    let
        ListPlaceholders = GetPlaceholders(0, "#{", "}", true),
        RecordPlaceholders = GetPlaceholders(0, "#[", "]", false),
        GetPlaceholders = (offset as number, openCode as text, closeCode as text, numberReference as logical) as list =>
            let
                TagLength = Text.Length(openCode) + Text.Length(closeCode),
                RelativeOpen = Text.PositionOf(Text.Range(formatString, offset), openCode, Occurrence.First),
                NextOpen = if RelativeOpen = -1 then -1 else RelativeOpen + offset,
                Length = Text.PositionOf(Text.Range(formatString, NextOpen + Text.Length(openCode)), closeCode, Occurrence.First),
                ReferenceText = Text.Range(formatString, NextOpen + Text.Length(openCode), Length),
                NumberReference = try Number.FromText(ReferenceText),
                FoundPlaceholder = 
                    if NextOpen <> -1 and Length <> -1 then
                        {[
                            Offset = NextOpen,
                            Reference =
                                if numberReference then
                                    try Number.FromText(ReferenceText) otherwise -1
                                else
                                    ReferenceText,
                            Length = Text.Length(ReferenceText) + TagLength
                        ]}
                    else
                        {},
                NewOffset = 
                    if NextOpen <> -1 and Length <> -1 then
                        NextOpen + Length + TagLength
                    else
                        Text.Length(formatString),
                Result = 
                    if offset = Text.Length(formatString) then
                        {}
                    else if NextOpen <> -1 and Length = -1 then
                        error Error.Record("DataFormat.Error", LibraryModule!UICulture.GetString("TextFormat_OpenWithoutClose"), NextOpen)
                    else
                        List.Combine({ FoundPlaceholder, @GetPlaceholders(NewOffset, openCode, closeCode, numberReference) })
            in
                Result,
        Placeholders = List.Combine({ RecordPlaceholders, ListPlaceholders })
    in
        Placeholders;
            
Text.FormatToText = (value, optional culture as text) as text =>
    if value = null then
        "null"
    else if Value.Is(value, type list) then
        "list"
    else if Value.Is(value, type record) then
        "record"
    else if Value.Is(value, type function) then
        "function"
    else if Value.Is(value, type table) then
        "table"
    else if Value.Is(value, type type) then
        "type"
    else if Value.Is(value, type action) then
        "action"
    else
        Text.From(value, culture);
    
