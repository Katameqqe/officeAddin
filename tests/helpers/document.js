class Element
{
    constructor()
    {
        this.style = { display: "", };
        this.innerHTML = "";
        this.id = "";
        this.children = [];
        this.onclick =
            () =>
            {
            };
        this.classList =
            {
                add:
                    () =>
                    {
                    },
            };
    }

    appendChild(aChild)
    {
        this.children.push(aChild);
    }
}

class Document
{
    constructor()
    {
        this.elements = {};
        this.elements["app-body"] = new Element();
    }

    getElementById(id)
    {
        let result = this.elements[id];
        if (result)
        {
            return result;
        }

        for (const element of Object.values(this.elements))
        {
            for (const child of element.children)
            {
                if (child.id == id)
                {
                    result = child;
                    break;
                }
            }
        }

        return result;
    }

    createElement(anElementName)
    {
        // this.elements[anElementName] = new Element();
        return new Element();
    }
}
module.exports = Document;
