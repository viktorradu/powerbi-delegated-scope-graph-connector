// This file contains your Data Connector logic
section DelegatedScopeGraphConnector;

//OAuth2 values
redirect_uri = "https://oauth.powerbi.com/views/oauthredirect.html";
token_uri = "https://login.windows.net/common/oauth2/token";
logout_uri = "https://login.microsoftonline.com/logout.srf";

[DataSource.Kind="DelegatedScopeGraphConnector", Publish="DelegatedScopeGraphConnector.Publish"]
shared DelegatedScopeGraphConnector.Contents = Value.ReplaceType(DelegatedScopeGraphConnectorImpl, DelegatedScopeGraphConnectorType);

DelegatedScopeGraphConnectorType = type function(
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
    ])
) as table meta [
        Documentation.Name = Extension.LoadString("ConnectorName"),
        Documentation.LongDescription = Extension.LoadString("ConnectorDescription")
    ];

DelegatedScopeGraphConnectorImpl = (tenant as text, clientId as text, clientSecret as text, optional relativePath as text) =>
    let
        result = OData.Feed("https://graph.microsoft.com/" & relativePath, null, [Implementation="2.0"])
    in
        result;

// Data Source Kind description
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

// Data Source UI publishing description
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

Logout = (token) => logout_uri;

StartLogin = (clientApplication, dataSourcePath, state, display) =>
    let
        authorizeUrl = "https://login.windows.net/" & Json.Document(dataSourcePath)[tenant] & "/oauth2/authorize?" & Uri.BuildQueryString([
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
            redirect_uri = redirect_uri,
            client_id = Json.Document(dataSourcePath)[clientId],
            client_secret = Json.Document(dataSourcePath)[clientSecret],
            resource="https://graph.microsoft.com/"
        ],
        queryWithCode = Record.AddField(queryString, tokenField, code),

        tokenResponse = Web.Contents(token_uri, [
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

