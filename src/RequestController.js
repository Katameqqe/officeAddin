class RequestController
{
    constructor()
    {
        this.address = "https://192.168.128.4:443";
        this.classificationLabelEndpoint = "/api/v1/classification-labels";
        this.xmlFontEndpoint = "/api/v1/xml-fonts";
    }
    async getClassifcationLabels()
    {
        const List = await fetch(this.address + this.classificationLabelEndpoint)
            .then(res => res.json())
            .then(resJson => resJson.names)
            .catch(
                err =>
                {
                    console.error("Error fetching classification labels list:", err);
                    return ["Document", "Default", "Restricted", "Protected",];
                });

        console.log(JSON.stringify(List,null,2));
        return List;
    }

    async getClassificationFonts()
    {
        const List = await fetch(this.address + this.xmlFontEndpoint)
            .then(res => res.json())
            .then(resJson => resJson)
            .catch(
                err =>
                {
                    // TODO: function name get classification labels, but in log "suffix".
                    // What do we get or fetch? suffixes?
                    console.error("Error fetching classification fonts:", err);
                    return defaultClassificationFont;
                });

        console.log(JSON.stringify(List,null,2));
        return List;
    }

}

module.exports = RequestController;