class result
{
    static json()
    {
        return {names: ["Document", "Default", "Restricted", "Protected", "NoLabel",]};
    }

}

async function fetch(params) 
{
    return result;
}

module.exports.fetch = fetch;