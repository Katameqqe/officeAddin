class result
{
    static json()
    {
        return {names: ["Document", "Default", "Restricted", "Protected",],};
    }

}

async function fetch(params)
{
    return result;
}

module.exports = fetch;
