section DelegatedScopeGraphConnector;

[DataSource.Kind="DelegatedScopeGraphConnector", Publish="DelegatedScopeGraphConnector.Publish"]
shared DelegatedScopeGraphConnector.Contents = Value.ReplaceType(DelegatedScopeGraphConnectorImpl, DelegatedScopeGraphConnectorType);

DelegatedScopeGraphConnectorType = type function(
    targetCloud as (type text meta [
        Documentation.FieldCaption = "Environment",
        Documentation.FieldDescription = "Target cloud environment",
        Documentation.AllowedValues = {"Commercial", "GCC High"}
    ]),
    tenant as (type text meta [
        Documentation.FieldCaption = "Tenant",
        Documentation.FieldDescription = "Azure AD tenant Id"
    ]),
    clientId as (type text meta [
        Documentation.FieldCaption = "Client Id",
        Documentation.FieldDescription = "App Registration client Id"
    ]),
    clientSecret as (type text meta [
        Documentation.FieldCaption = "Client Secret",
        Documentation.FieldDescription = "App Registration client secret"
    ]),
    optional relativePath as (type text meta [
        Documentation.FieldCaption = Extension.LoadString("RelativePathLabel"),
        Documentation.FieldDescription = Extension.LoadString("RelativePathDescription"),
        Documentation.SampleValues = {"v1.0/users", "beta/security/alerts"}
    ]),
    optional pageSize as (type number meta [
        Documentation.FieldCaption = Extension.LoadString("PageSizeLabel"),
        Documentation.FieldDescription = Extension.LoadString("PageSizeDescription"),
        Documentation.DefaultValue = 100
    ])
) as table meta [
        Documentation.Name = Extension.LoadString("ConnectorName"),
        Documentation.LongDescription = Extension.LoadString("ConnectorDescription")
    ];

DelegatedScopeGraphConnectorImpl = (targetCloud as text, tenant as text, clientId as text, clientSecret as text, optional relativePath as text, optional pageSize as number) as table =>
    let
        result = GetAllPagesByNextLink(GetEnvironmentValue(targetCloud, "resource") & relativePath, pageSize)
    in
        result;

GetAllPagesByNextLink = (url as text, optional pageSize as number) as table =>
     Table.GenerateByPage((previous) => 
            let
                nextLink = if (previous = null) then url else Value.Metadata(previous)[NextLink]?,
                page = if (nextLink <> null) then GetPage(nextLink, pageSize) else null
            in
                page
        );

GetPage = (url as text, optional pageSize as number) as table =>
    let
        headers = if(pageSize <> null) then [Prefer = "odata.maxpagesize=" & Number.ToText(pageSize)] else [],
        response = Web.Contents(url, [Headers = headers]),        
        document = Json.Document(response),
        valuesList = if Record.HasFields(document, {"value"}) then document[value] else {document},
        valuesTable = Table.FromList(valuesList, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
        result = try Table.ExpandRecordColumn(valuesTable, "Column1", Record.FieldNames(Table.Column(valuesTable, "Column1"){0})) otherwise #table({},{}),
        nextLink = Record.FieldOrDefault(document, "@odata.nextLink")
    in
        result meta [NextLink = nextLink];

DelegatedScopeGraphConnector = [
    TestConnection = (dataSourcePath) => { "DelegatedScopeGraphConnector.Contents", dataSourcePath },
    Authentication = [
        OAuth = [
            StartLogin=StartLogin,
            FinishLogin=FinishLogin,
            Refresh=Refresh,
            Logout=Logout,
            Label="Delegated Scope Graph Connector"
            ]
    ],
    Label = Extension.LoadString("DataSourceLabel")
];

DelegatedScopeGraphConnector.Publish = [
    Beta = true,
    Category = "Other",
    ButtonText = { Extension.LoadString("ConnectorName"), Extension.LoadString("ConnectorDescription") },
    LearnMoreUrl = "https://powerbi.microsoft.com/",
    SourceImage = DelegatedScopeGraphConnector.Icons,
    SourceTypeImage = DelegatedScopeGraphConnector.Icons
];

DelegatedScopeGraphConnector.Icons = [
    Icon16 = { Extension.Contents("delegated-graph16.png"), Extension.Contents("delegated-graph20.png"), Extension.Contents("delegated-graph24.png"), Extension.Contents("delegated-graph32.png") },
    Icon32 = { Extension.Contents("delegated-graph32.png"), Extension.Contents("delegated-graph40.png"), Extension.Contents("delegated-graph48.png"), Extension.Contents("delegated-graph64.png") }
];

Refresh = (clientApplication, dataSourcePath, oldCredential) => 
    let 
        result = TokenMethod(dataSourcePath, "refresh_token", "refresh_token", oldCredential[refresh_token])
    in
        result;

Logout = (clientApplication, dataSourcePath, token) => GetEnvironmentValue(Json.Document(dataSourcePath)[targetCloud], "logout_uri");

StartLogin = (clientApplication, dataSourcePath, state, display) =>
    let
        redirect_uri = GetEnvironmentValue(Json.Document(dataSourcePath)[targetCloud], "redirect_uri"),
        authorizeUrl = GetEnvironmentValue(Json.Document(dataSourcePath)[targetCloud], "auth_root") & Json.Document(dataSourcePath)[tenant] & "/oauth2/authorize?" & Uri.BuildQueryString([
            response_type = "code",
            client_id = Json.Document(dataSourcePath)[clientId],  
            redirect_uri = redirect_uri,
            state = state,
            scope = "openid"
        ])
    in
        [
            LoginUri = authorizeUrl,
            CallbackUri = redirect_uri,
            WindowHeight = 720,
            WindowWidth = 1024,
            Context = null
        ];

FinishLogin = (clientApplication, dataSourcePath, context, callbackUri, state) =>
    let
        // parse the full callbackUri, and extract the Query string
        parts = Uri.Parts(callbackUri)[Query],
        // if the query string contains an "error" field, raise an error
        // otherwise call TokenMethod to exchange our code for an access_token
        result = if (Record.HasFields(parts, {"error", "error_description"})) then 
                    error Error.Record(parts[error], parts[error_description], parts)
                 else
                    TokenMethod(dataSourcePath, "authorization_code", "code", parts[code])
    in
        result;

TokenMethod = (dataSourcePath, grantType, tokenField, code) =>
    let
        queryString = [
            grant_type = grantType,
            redirect_uri = GetEnvironmentValue(Json.Document(dataSourcePath)[targetCloud], "redirect_uri"),
            client_id = Json.Document(dataSourcePath)[clientId],
            client_secret = Json.Document(dataSourcePath)[clientSecret],
            resource=GetEnvironmentValue(Json.Document(dataSourcePath)[targetCloud], "resource")
        ],
        queryWithCode = Record.AddField(queryString, tokenField, code),

        tokenResponse = Web.Contents(GetEnvironmentValue(Json.Document(dataSourcePath)[targetCloud], "token_uri"), [
            Content = Text.ToBinary(Uri.BuildQueryString(queryWithCode)),
            Headers = [
                #"Content-type" = "application/x-www-form-urlencoded",
                #"Accept" = "application/json"
            ],
            ManualStatusHandling = {400} 
        ]),
        body = Json.Document(tokenResponse),
        result = if (Record.HasFields(body, {"error", "error_description"})) then 
                    error Error.Record(body[error], body[error_description], body)
                 else
                    body
    in
        result;

GetEnvironmentValue = (targetCloud as text, key as text) =>
    let
        valuesCommercial = [
            resource = "https://graph.microsoft.com/",
            auth_root = "https://login.microsoftonline.com/",
            token_uri = "https://login.microsoftonline.com/common/oauth2/token",
            redirect_uri = "https://oauth.powerbi.com/views/oauthredirect.html",
            logout_uri = "https://login.microsoftonline.com/logout.srf"],
        valuesGCCHigh = [
            resource = "https://graph.microsoft.us/",
            auth_root = "https://login.microsoftonline.us/",
            token_uri = "https://login.microsoftonline.us/common/oauth2/token",
            redirect_uri = "https://oauth.powerbigov.us/views/oauthredirect.html",
            logout_uri = "https://login.microsoftonline.us/logout.srf"],

        value = if targetCloud = "GCC High" then Record.Field(valuesGCCHigh, key) else Record.Field(valuesCommercial, key)
    in
        value;


    //
// Common functions from: https://learn.microsoft.com/en-us/power-query/helper-functions
//
// The getNextPage function takes a single argument and is expected to return a nullable table
Table.GenerateByPage = (getNextPage as function) as table =>
    let        
        listOfPages = List.Generate(
            () => getNextPage(null),            // get the first page of data
            (lastPage) => lastPage <> null,     // stop when the function returns null
            (lastPage) => getNextPage(lastPage) // pass the previous page to the next function call
        ),
        // concatenate the pages together
        tableOfPages = Table.FromList(listOfPages, Splitter.SplitByNothing(), {"Column1"}),
        firstRow = tableOfPages{0}?
    in
        // if we didn't get back any pages of data, return an empty table
        // otherwise set the table type based on the columns of the first page
        if (firstRow = null) then
            Table.FromRows({})
        else        
            Value.ReplaceType(
                Table.ExpandTableColumn(tableOfPages, "Column1", Table.ColumnNames(firstRow[Column1])),
                Value.Type(firstRow[Column1])
            );
