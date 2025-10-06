const defaultClassificationFont = 
{
    "hdr":
        [
            {
                "fontName": "Arial",
                "fontColor": "000000",
                "fontSize": "14",
                "text": "Sample Watermark",
            },
        ],
    "ftr":
        [
            {
                "fontName": "Verdana",
                "fontColor": "FF0000",
                "fontSize": "12",
                "text": "Second Line",
            },
        ],
    "wm":
    {
        "fontName": "Arial",
        "fontColor": "000000",
        "fontSize": "36",
        "rotation": "315",
        "transparency": "0.5",
        "text": "Confidential",
    },
};

const defaultClassificationLabel =
[
    "Document",
    "Default",
    "Restricted",
    "Protected"
];

module.exports.defaultClassificationFont = defaultClassificationFont;
module.exports.defaultClassificationLabel = defaultClassificationLabel;