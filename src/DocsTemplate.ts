export default class DocTemplate {
    private Doc: GoogleAppsScript.Document.Document;
    constructor(id: string) {
        this.Doc = DocumentApp.openById(id);
    }
    render(obj: Object) {
        let count = 0;
        while (this.Doc.getBody().findText("{{.+?}}")) {
            if (count > 100) { break; }
            const RangeElement = this.Doc.getBody().findText("{{.+?}}");
            if (!RangeElement) { break; };

            const Element = RangeElement.getElement();
            const Attributes = Element.getAttributes();
            const templateText = new TemplateText(Element.asText().getText());
            Element.asText().setText(templateText.render(obj)).setAttributes(Attributes);
            count++;
        }
        return this.Doc;
    }
}

class TemplateText {
    Text: string;
    constructor(text: string) {
        this.Text = text;
        return this;
    }
    render(obj: Object) {
        const Matched = this.Text.match(/{{.+?}}/g);
        if (Matched) {
            Matched.forEach(template => {
                const key = template.slice(2, -2);
                const value = (obj as any)[key];

                if (obj.hasOwnProperty(key)) {
                    switch (Object.prototype.toString.call(value).toLowerCase().slice(8, -1)) {
                        case 'string':
                            this.Text = this.Text.replace(template, value);
                            break;
                        case 'date':
                            this.Text = this.Text.replace(template, (value as Date).toISOString());
                            break;
                        case 'array':
                        case 'object':
                        default:
                            this.Text = this.Text.replace(template, JSON.stringify(value));
                            break;
                    }
                } else {
                    this.Text = this.Text.replace(template, "");
                }
            });
        }
        return this.Text;
    }
}