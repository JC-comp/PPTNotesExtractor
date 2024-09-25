// https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_notes_topic_ID0E3F3GB.html
class Styler {
    constructor(stylerXML) {
        this.stylerXML = stylerXML;
        this.isBullet = false;
        this.marL = 0;
        this.b = 0;
        this.i = 0;
        this.u = 0;

        if (stylerXML == null)
            return;
        
        if (stylerXML.getElementsByTagName("a:buChar").length > 0)
            this.isBullet = true;
        if (stylerXML.getElementsByTagName("a:buAutoNum").length > 0)
            this.isBullet = true;
        if (stylerXML.getAttribute("marL") != null)
            this.marL = parseInt(stylerXML.getAttribute("marL")) / 10000;
        if (stylerXML.getAttribute("indent") != null)
            this.marL += parseInt(stylerXML.getAttribute("indent")) / 10000;
        if (stylerXML.getAttribute("b") != null)
            this.b = parseInt(stylerXML.getAttribute("b"));
        if (stylerXML.getAttribute("i") != null)
            this.i = parseInt(stylerXML.getAttribute("i"));
        if (stylerXML.getAttribute("u") != null && stylerXML.getAttribute("u") != 'none') {
            this.u = stylerXML.getAttribute("u");
        }
    }

    getContent(content) {
        return this.getRawContent(content);
    }

    getRawContent(content) {
        return "<span " + this.getStyle() + ">" + content + "</span>"
    }

    hasWrap() {
        return this.isBullet;
    }

    getStyle() {
        var style = 'style="' + this.getRawStyle() + '"';
        return style
    }

    getRawStyle() {
        var style = '';
        if (this.marL != 0) {
            style += "margin-left: " + this.marL + "px;";
        }
        if (this.b == 1) {
            style += "font-weight: bold;";
        }
        if (this.i != 0) {
            style += "font-style: italic;";
        }
        if (this.u != 0) {
            style += "text-decoration: underline;";
        }
        return style;
    }

    wrap(content) {
        if (this.isBullet && content.length > 0)
            return "<li " + this.getStyle() + ">" + content + "</li>"
        else
            return "<p " + this.getStyle() + ">" + content + "</p>";
    }
}
class Run {
    constructor(runXML) {
        this.runXML = runXML;
        this.styler = new Styler();

        var rPr = this.runXML.getElementsByTagName("a:rPr");
        if (rPr.length > 0)
            this.styler = new Styler(rPr[0]);
    }

    getContent() {
        return this.styler.getContent(this.runXML.textContent);
    }

    getRawContent() {
        return this.styler.getRawContent(this.runXML.textContent);
    }
}
class Paragraph {
    constructor(paragraphXML) {
        this.paragraphXML = paragraphXML;
        this.runs = []
        this.styler = new Styler();
        
        var runs = this.paragraphXML.getElementsByTagName("a:r");
        for (let run of runs) {
            this.runs.push(new Run(run));
        }

        var pPr = this.paragraphXML.getElementsByTagName("a:pPr");
        if (pPr.length > 0)
            this.styler = new Styler(pPr[0]);
    }

    wrap(content) {
        return this.styler.wrap(content);
    }

    getContent() {
        var content = '';
        for (let run of this.runs) {
            content += run.getContent();
        }
        content = this.wrap(content);
        return content;
    }

    getRawContent() {
        var content = '';
        for (let run of this.runs) {
            content += run.getRawContent();
        }
        content = this.wrap(content);
        return content;
    }
}
class Page {
    constructor(pageXML) {
        this.pageXML = pageXML;
        this.sldNum = -1;
        this.paragraphs = [];
        
        var shapes = this.pageXML.getElementsByTagName("p:sp");
        for (let shape of shapes) {
            if (shape.querySelector('ph[type="body"]'))  {
                var body = shape.getElementsByTagName("p:txBody");
                if (body.length > 0) {
                    var ps = body[0].getElementsByTagName("a:p");
                    for (let p of ps) 
                        this.paragraphs.push(new Paragraph(p));
                }
            } else if (shape.querySelector('ph[type="sldNum"]')) {
                this.sldNum = parseInt(shape.textContent);
            }
        }
    }

    getContent() {
        var content = '';
        for (let paragraph of this.paragraphs) {
            content += paragraph.getContent();
        }
        return content;
    }

    getRawContent() {
        var content = '';
        for (let paragraph of this.paragraphs) {
            content += paragraph.getRawContent();
        }
        return content;
    }
}

