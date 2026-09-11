var docx = require("docx")
var xml = require("xml")
var htmlparser = require("htmlparser2")
var domutils = require("domutils")
const render = require('dom-serializer').default
const hljs = require('highlight.js');
 
function html2ooxml(html, style = '') {

 
    if (html === '')
        return html
 
    if (!html.match(/^<.+>/))
        html = `<p>${html}</p>`
 
    html = html.replace(/<comment.*?>/g, '').replace(/<\/comment>/g, '')
 
    var doc = new docx.Document({sections:[]});
 
    var paragraphs = []
    var cParagraph = null
 
    // Gestion tableaux
    var currentTable = null
    var currentRow = null
    var currentCell = null
    var inTable = false
 
    var cRunProperties = {}
    var cParagraphProperties = {}
    var list_state = []
    var inCodeBlock = false
 
 
    var parser = new htmlparser.Parser({
 
        onopentag(tag, attribs) {
 
 
            /*
             * TABLE SUPPORT
             */
 
            if (tag === "table") {
 
                inTable = true
                currentTable = []
 
            }
 
            else if (tag === "tr" && inTable) {
 
                currentRow = []
 
            }
 
            else if ((tag === "td" || tag === "th") && inTable) {
 
                currentCell = new docx.Paragraph({})
 
            }
 
 	   else if ((tag === "div" || tag === "p") && inTable) {
                cParagraph = currentCell;
	    }
            else if (tag === "h1") {
 
                cParagraph = new docx.Paragraph({
                    heading: 'Heading1'
                })
 
            }
 
            else if (tag === "h2") {
 
                cParagraph = new docx.Paragraph({
                    heading: 'Heading2'
                })
 
            }
 
            else if (tag === "h3") {
 
                cParagraph = new docx.Paragraph({
                    heading: 'Heading3'
                })
 
            }
 
            else if (tag === "h4") {
 
                cParagraph = new docx.Paragraph({
                    heading: 'Heading4'
                })
 
            }
 
            else if (tag === "h5") {
 
                cParagraph = new docx.Paragraph({
                    heading: 'Heading5'
                })
 
            }
 
            else if (tag === "h6") {
 
                cParagraph = new docx.Paragraph({
                    heading: 'Heading6'
                })
 
            }
 
 
            else if (tag === "div" || tag === "p") {
 
                if (style && typeof style === 'string')
                    cParagraphProperties.style = style
 
                cParagraph = new docx.Paragraph(cParagraphProperties)
 
            }
 
 
            else if (tag === "pre") {
 
                inCodeBlock = true
 
                cParagraph = new docx.Paragraph({
                    style: "Code"
                })
 
            }
 
 
            else if (tag === "b" || tag === "strong") {
 
                cRunProperties.bold = true
 
            }
 
 
            else if (tag === "i" || tag === "em") {
 
                cRunProperties.italics = true
 
            }
 
 
            else if (tag === "u") {
 
                cRunProperties.underline = {}
 
            }
 
 
            else if (tag === "strike" || tag === "s") {
 
                cRunProperties.strike = true
 
            }
 
 
            else if (tag === "code") {
 
                cRunProperties.style = "CodeChar"
 
            }
 
 
            else if (tag === "br") {
 
 
                if (inCodeBlock) {
 
                    paragraphs.push(cParagraph)
 
                    cParagraph = new docx.Paragraph({
                        style:"Code"
                    })
 
                }
 
                else if(cParagraph) {
 
                    cParagraph.addChildElement(
                        new docx.Run({
                            break:1
                        })
                    )
 
                }
 
            }
 
 
            else if (tag === "ul") {
 
                list_state.push('bullet')
 
            }
 
 
            else if (tag === "ol") {
 
                list_state.push('number')
 
            }
 
 
            else if (tag === "li") {
 
                var level = list_state.length - 1
 
                if(level >=0 && list_state[level] === 'bullet')
 
                    cParagraphProperties.bullet = {
                        level:level
                    }
 
                else if(level >=0 && list_state[level] === 'number')
 
                    cParagraphProperties.numbering = {
                        reference:2,
                        level:level
                    }
 
            }
 
        },
 
 
        ontext(text) {
 
 
            if(!text)
                return
 
 
            /*
             * Texte dans tableau
             */
 
            if(inTable && currentCell) {
 
                currentCell.addChildElement(
                    new docx.TextRun(text)
                )
 
            }
 
 
            else if(cParagraph) {
 
                cRunProperties.text = text
 
                cParagraph.addChildElement(
                    new docx.TextRun(cRunProperties)
                )
 
            }
 
        },
 
 
        onclosetag(tag) {
 
 
            /*
             * FIN TABLEAU
             */
 
 
            if(tag === "td" || tag === "th") {
 
 
                if(currentRow && currentCell) {
 
                    currentRow.push(
                        new docx.TableCell({
                            children:[
                                currentCell
                            ]
                        })
                    )
 
                }
 
 
                currentCell = null
 
                return
 
            }
 
 
 
            if(tag === "tr") {
 
 
                if(currentTable && currentRow) {
 
 
                    currentTable.push(
 
                        new docx.TableRow({
 
                            children:currentRow
 
                        })
 
                    )
 
                }
 
 
                currentRow=null
 
                return
 
            }
 
 
 
            if(tag === "table") {
 
 
                paragraphs.push(
 
                    new docx.Table({
 
                        rows:currentTable
 
                    })
 
                )
 
 
                currentTable=null
                inTable=false
 
                return
 
            }
 
 
 
            if([
                'h1',
                'h2',
                'h3',
                'h4',
                'h5',
                'h6',
                'div',
                'p',
                'pre'
            ].includes(tag)) {
 
 
                if(cParagraph && !inTable)
                    paragraphs.push(cParagraph)
 
 
                cParagraph=null
 
                cParagraphProperties={}
 
 
                if(tag==="pre")
                    inCodeBlock=false
 
            }
 
 
            else if(tag==="b" || tag==="strong") {
 
                delete cRunProperties.bold
 
            }
 
 
            else if(tag==="i" || tag==="em") {
 
                delete cRunProperties.italics
 
            }
 
 
            else if(tag==="u") {
 
                delete cRunProperties.underline
 
            }
 
 
            else if(tag==="strike" || tag==="s") {
 
                delete cRunProperties.strike
 
            }
 
 
            else if(tag==="code") {
 
                delete cRunProperties.style
 
            }
 
 
            else if(tag==="ul" || tag==="ol") {
 
                list_state.pop()
 
            }
 
        },
 
 
        onend() {
 
 
            doc.addSection({
 
                children:paragraphs
 
            })
 
 
        }
 
 
    },
    {
        decodeEntities:true
    })
 
 
 
    var htmlObject = htmlparser.parseDocument(html);
 
 
    var codeObjects = domutils.getElementsByTagName(
        'code',
        htmlObject
    );
 
 
    for(let codeIndex=0; codeIndex<codeObjects.length; codeIndex++) {
 
 
        let codeClass = codeObjects[codeIndex].attribs.class
 
        let codeLanguage =
            (codeClass && codeClass.split('-')[1])
            || 'plaintext'
 
 
        let codeText =
            codeObjects[codeIndex].children[0]
            ? codeObjects[codeIndex].children[0].data
            : ''
 
 
        codeObjects[codeIndex].children[0] =
            htmlparser.parseDocument(
                hljs.highlight(
                    codeText,
                    {
                        language:codeLanguage
                    }
                ).value
            )
 
    }
 
 
 
    html = render(htmlObject).replace(/\n/g,'<br>')
 
 
    parser.write(html)
 
    parser.end()
 
 
 
    var prepXml =
        doc.documentWrapper.document.body.prepForXml({
            stack:[]
        })
 
 
    /*
     * IMPORTANT :
     * conserver les paragraphes ET les tableaux
     */
 
    var filteredXml =
        prepXml["w:body"].filter(e => {
 
            return e &&
            (
                Object.keys(e)[0] === "w:p" ||
                Object.keys(e)[0] === "w:tbl"
            )
 
        })
 
 
    var dataXml = xml(filteredXml)
 
 
    dataXml =
        dataXml.replace(
            /w:numId w:val="{2-0}"/g,
            'w:numId w:val="2"'
        )
 
 
    return dataXml
 
}
 
 
module.exports = html2ooxml
