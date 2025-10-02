class result
{
    static json()
    {
        return {names: ["Document", "Default", "Restricted", "Protected",],};
    }

}

class fontresult
{
    static json()
    {
        return defaultClassificationFont;
    }
}

async function fetch(params)
{
    if(params.endsWith("/api/v1/classification-labels"))
    {
        return result;
    }
    if(params.endsWith("/api/v1/xml-fonts"))
    {
        return fontresult;
    }
}

module.exports = fetch;
