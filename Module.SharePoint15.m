// Copyright (c) Microsoft Corporation.  All rights reserved.
    
section SharePoint15;

shared SharePoint15.Tables = (url as text) as table =>
    let
        lists = OData.Feed(url & "/lists"),
        listsTable = Table.SelectColumns(lists, {"Id", "Title", "Items"}), 
        result = ToNavigationTable(listsTable, {"Id"}, "Title", "Items", false) 
    in
        result;

 shared SharePoint15.Contents = (url as text) as table =>
    let
        webUrl = url & "/Folders",
        source = OData.Feed(webUrl),
        folders = CreateContent(source, url, true) 
    in
        Value.ReplaceType(folders, ContentSchema(DateTimeZone.Type, null, Any.Type meta [Preview.Delay="Table"]));

shared SharePoint15.Files = (url as text) as table =>
    let
        webUrl = url & "/lists/?$filter=" & Uri.EscapeDataString("Hidden eq false and BaseType eq 1 and ItemCount gt 0 and IsCatalog eq false and IsApplicationList eq false"),
        source = OData.Feed(webUrl),
        files = CreateFiles(source, url) 
    in
        files;

//
// Transformations
//

ContentSchemaWithZone = ContentSchema(DateTimeZone.Type);

CreateContent = (table as table, siteUrl as text, isTable as logical) as table =>
    let
        withSiteUrl = Table.TransformRows(table, (r) => Record.AddField(r, "SiteUrl", siteUrl)),
        transformation = if isTable then TransformToFolder else TransformToFile,
        result = Table.FromRecords(List.Transform(withSiteUrl, transformation), ContentSchemaWithZone) 
    in
        result;

CreateFiles = (table as table, apiUrl as text) as table =>
    let
        tablesList = List.Transform(Table.ToRecords(table), each TransformToFiles(_, apiUrl)),
        files = Table.Combine(tablesList),
        result = if Table.ColumnCount(files) = 0 then files else Value.ReplaceType(files, ContentSchema(DateTime.Type, {"Name", "Folder Path"}, Binary.Type meta [Preview.Delay="Binary"])) 
    in
        result;

ContentSchema = (dateType as type, optional keyColumns as list, optional contentType as type) =>
    let
        keys = if keyColumns = null then {"Name"} else keyColumns,
        content = if contentType <> null then contentType else Any.Type,
        schema = type table 
            [
                Content=content, 
                Name=text, 
                Extension=text, 
                Date accessed=dateType, 
                Date modified=dateType, 
                Date created=dateType, 
                Attributes=record, 
                Folder Path=text
            ] 
    in
        Type.AddTableKey(schema, keys, true);

CreateFolderContent = (row as record) as table =>
    let
        subFolders = Table.View(null,
        [
            GetRows = () => CreateContent(row[Folders], row[SiteUrl], true),
            GetType = () => ContentSchemaWithZone
        ]),
        files = Table.View(null,
        [
            GetRows = () => CreateContent(row[Files], row[SiteUrl], false),
            GetType = () => ContentSchemaWithZone
        ]),
        combinedTable = Table.AddKey(Table.Combine({subFolders, files}), {"Name"}, true) 
    in
        combinedTable;

CreateFilesContent = (files as list, tableType as type, authority as text) as table =>
    let
        withSiteUrl = List.Transform(files, (r) => Record.AddField(r, "SiteUrl", authority)),
        results = Table.FromRecords(List.Transform(withSiteUrl, TransfromFromFileList), tableType) 
    in
        results;
     
TransformToFolder = (row as record) as record =>
    let
        ParentFolder = GetParentFolder(row[ServerRelativeUrl]),
        FolderPath = NormalizeFolderPath(GetUrlAuthority(row[SiteUrl]) & ParentFolder),
        Result = [
            Content = CreateFolderContent(row),
            Name = row[Name],
            Extension = "",
            Date accessed = null,
            Date modified = DateTimeZone.From(row[TimeLastModified]),
            Date created = DateTimeZone.From(row[TimeCreated]),
            Attributes = [Size = null, Content Type = null, Kind = "Folder"],
            Folder Path = FolderPath] 
    in
        Result;

TransformToFile = (row as record) as record =>
    let
        FilePath = row[SiteUrl] & "/getfilebyserverrelativeurl('" & row[ServerRelativeUrl] & "')/$value",
        FolderPath = GetFileDirectory(GetUrlAuthority(row[SiteUrl]), row[ServerRelativeUrl]),
        Content = Web.Contents(FilePath),
        temp = Record.AddField(
            [Size = row[Length]], 
            "Content Type", 
            () => Value.Metadata(Content)[Content.Type], 
            true),
        Attributes = Record.AddField(
            temp, 
            "Kind", 
            () => GetFileKind(Value.Metadata(Content)[Content.Type]), 
            true),
        Result = [
            Content = Content,
            Name = row[Name],
            Extension = GetFileExtension(row[Name]),
            Date accessed = null,
            Date modified = DateTimeZone.From(row[TimeLastModified]),
            Date created = DateTimeZone.From(row[TimeCreated]),
            Attributes = Attributes,
            Folder Path = FolderPath]
    in
        Result;

TransfromFromFileList = (row as record) as record =>
    let
        FilePath = row[SiteUrl] & row[FileRef],
        FolderPath = GetFileDirectory(row[SiteUrl], row[FileRef]),
        Content = Web.Contents(FilePath),
        temp = Record.AddField(
            [Size = row[File_x0020_Size]], 
            "Content Type", 
            () => Value.Metadata(Content)[Content.Type], 
            true),
        Attributes = Record.AddField(
            temp, 
            "Kind", 
            () => GetFileKind(Value.Metadata(Content)[Content.Type]), 
            true),
        extension = NormalizeExtension(row[FileLeafRef.Suffix]),
        Result = [
            Content = Content,
            Name = row[FileLeafRef],
            Extension = extension,
            Date accessed = null,
            Date modified = row[Modified],
            Date created = row[Created],
            Attributes = Attributes,
            Folder Path = FolderPath] 
    in
        Result;

TransformToFiles = (row as record, apiUrl as text) as table =>
    let
        id = row[Id],
        contents = Web.Contents(
            apiUrl & "/lists/getbyid('" & id & "')/RenderListDataAsStream", 
            [Headers = [#"Content-Type" = "application/json;odata=verbose"],
            Content = Text.ToBinary("{'parameters': {'__metadata': { 'type': 'SP.RenderListDataParameters' }, 'ViewXml': '<View Scope=""RecursiveAll""><Query><Where><Eq><FieldRef Name=""FSObjType"" /><Value Type=""Integer"">0</Value></Eq></Where></Query><ViewFields><FieldRef Name=""FSObjType"" /><FieldRef Name=""LinkFilename"" /><FieldRef Name=""Modified"" /><FieldRef Name=""BaseName"" /><FieldRef Name=""FileSizeDisplay"" /><FieldRef Name=""Created"" /><FieldRef Name=""FileLeafRef.Suffix"" /><FieldRef Name=""FileRef"" /></ViewFields></View>', 'RenderOptions': 2}}")]),
        jsonDocument = Json.Document(contents),
        tableType = ContentSchema(DateTime.Type,{"Name", "Folder Path"}, Binary.Type meta [Preview.Delay="Binary"]),
        Result = Table.View(null,
        [
            GetRows = () => CreateFilesContent(jsonDocument[Row], tableType, GetUrlAuthority(apiUrl)),
            GetType = () => tableType
        ])
    in
        Result;

GetFileExtension = (fileName as text) =>
    let
        parts = Text.Split(fileName, "."),
        extension = List.Last(parts, null) 
    in
        NormalizeExtension(extension);

NormalizeExtension = (extension as text) =>
    let
        result = if Text.StartsWith(extension, ".") then extension else "." & extension
    in
        result;

GetFileKind = (contentType as text) =>
    let
        kind = if Value.Equals(contentType, "application/msaccess")
            then "Access File"
            else if Value.Equals(contentType, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            then "Excel File"
            else if Value.Equals(contentType, "application/vnd.ms-excel")
            then "Excel File"
            else if Value.Equals(contentType, "application/javascript")
            then "Javascript File"
            else if Value.Equals(contentType, "text/javascript")
            then "Javascript File"
            else if Value.Equals(contentType, "application/json")
            then "JSON File"
            else if Value.Equals(contentType, "text/x-json")
            then "JSON File"
            else if Value.Equals(contentType, "application/xhtml+xml")
            then "HTML File"
            else if Value.Equals(contentType, "text/html")
            then "HTML File"
            else if Value.Equals(contentType, "text/csv")
            then "CSV File"
            else if Value.Equals(contentType, "text/xml")
            then "XML File"
            else if Text.StartsWith(contentType, "text/")
            then "Text File"
            else "File" 
    in
        kind;

GetUrlAuthority = (url as text) as text =>
    let
        parts = Uri.Parts(url),
        port = if (parts[Scheme] = "https" and parts[Port] = 443) or (parts[Scheme] = "http" and parts[Port] = 80) then "" else ":" & Text.From(parts[Port]) 
    in
        parts[Scheme] & "://" & parts[Host] & port;

GetFileDirectory = (baseUrl as text, filePath as text) as text =>
    let
        path = GetParentFolder(filePath),
        directory = baseUrl & path
    in
        NormalizeFolderPath(directory);

NormalizeFolderPath = (folderPath as text) as text =>
    let
        result = if Text.EndsWith(folderPath, "/") then folderPath else folderPath & "/"
    in
        result;

GetParentFolder = (relativePath as text) as text =>
    let
        trimmedPath = Text.TrimEnd(relativePath, "/"),
        index = Text.PositionOf(trimmedPath, "/", Occurrence.Last),
        count = Text.Length(trimmedPath) - index,
        path = if index = -1 then "" else Text.RemoveRange(trimmedPath, index, count) 
    in
        path;

ToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    isLink as logical
) as table =>
    let
        tableType = Value.Type(table),
        rowType = Type.TableRow(tableType),
        rowFields = Type.RecordFields(rowType),
        dataColumnField = Record.Field(rowFields, dataColumn),
        dataColumnType = dataColumnField[Type],

        newDataColumnType = if isLink 
            then dataColumnType meta [NavigationTable.ItemKind = "Table", Preview.Delay = "Table"] 
            else dataColumnType meta [NavigationTable.ItemKind = "Table", Preview.Delay = "Table", NavigationTable.IsLeaf = true],
        newDataColumnField = dataColumnField & [Type = newDataColumnType],
        newRowFields = rowFields & Record.FromList({newDataColumnField}, {dataColumn}),
        newRowType = Type.ForRecord(newRowFields, Type.IsOpenRecord(rowType)),
        newTableType = Type.AddTableKey((type table newRowType), keyColumns, true) meta [NavigationTable.NameColumn = nameColumn, NavigationTable.DataColumn = dataColumn],
        navigationTable = Value.ReplaceType(table, newTableType) 
    in
        navigationTable;
    