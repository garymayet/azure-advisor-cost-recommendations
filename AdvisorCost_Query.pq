let
    // Parámetro: ruta al CSV generado por el script PowerShell
    RutaCSV = CsvPath,
    Origen = Csv.Document(File.Contents(RutaCSV),[Delimiter=",", Columns=16, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promovido = Table.PromoteHeaders(Origen, [PromoteAllScalars=true]),
    Tipos = Table.TransformColumnTypes(Promovido,{
        {"Subscription", type text},
        {"SubscriptionId", type text},
        {"Category", type text},
        {"Impact", type text},
        {"Problem", type text},
        {"Solution", type text},
        {"ResourceType", type text},
        {"ResourceGroup", type text},
        {"ResourceId", type text},
        {"RecommendationType", type text},
        {"AnnualSavingsUSD", type number},
        {"Currency", type text},
        {"Term", type text},
        {"TargetSku", type text},
        {"Region", type text},
        {"LastUpdated", type datetime}
    }),
    // Nombre legible por tipo de recomendación
    ConNombreTipo = Table.AddColumn(Tipos, "RecommendationName", each
        if [RecommendationType] = "84b1a508-fc21-49da-979e-96894f1665df" then "VM Reserved Instance"
        else if [RecommendationType] = "89515250-1243-43d1-b4e7-f9437cedffd8" then "Savings Plan"
        else if [RecommendationType] = "e10b1381-5f0a-47ff-8c7b-37bd13d7c974" then "Right-size / Shutdown VM"
        else if [RecommendationType] = "407b6ad6-8e0b-40e7-9384-643520cae0ed" then "Azure Files Reserved Instance"
        else if [RecommendationType] = "5b8ddf04-be28-44ec-ab2c-a63a34d1de13" then "App Service Reserved Instance"
        else if [RecommendationType] = "39a8510f-5bbf-4304-9bcd-4106c996473b" then "Unused/Empty App Service Plan"
        else "Otro", type text),
    // Extrae nombre del recurso al final del ResourceId
    ConRecurso = Table.AddColumn(ConNombreTipo, "ResourceName", each
        let parts = Text.Split([ResourceId], "/") in List.Last(parts), type text),
    // Bandera para segmentar
    ConTerm = Table.AddColumn(ConRecurso, "TermClean", each
        if [Term] = "P1Y" then "1 Año"
        else if [Term] = "P3Y" then "3 Años"
        else "N/A", type text)
in
    ConTerm
