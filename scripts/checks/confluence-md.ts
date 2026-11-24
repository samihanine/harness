import { toMarkdown, toStorage } from "@server/confluence";
const html = `<p>A&nbsp;B &eacute;t&eacute; <ac:link><ri:page ri:content-title="Autre"/><ac:plain-text-link-body><![CDATA[voir]]></ac:plain-text-link-body></ac:link></p><ac:structured-macro ac:name="info"><ac:rich-text-body><p>Note importante</p></ac:rich-text-body></ac:structured-macro><h1>Guide</h1><p>Texte <strong>gras</strong> et <a href="https://x.y">lien</a>.</p><ul><li>un<ul><li>sous</li></ul></li><li>deux</li></ul><table><tbody><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></tbody></table><ac:structured-macro ac:name="code"><ac:plain-text-body><![CDATA[EVALUATE <x>]]></ac:plain-text-body></ac:structured-macro>`;
const md = toMarkdown(html);
console.log(md);
console.log("---");
console.log(toStorage(md));
