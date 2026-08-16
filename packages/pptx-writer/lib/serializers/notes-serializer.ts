import type { PptxTextBody } from '@hokkyss/pptx-core';
import { createXmlBuilder } from '../xml/xml-builder';
import { serializeTextBody } from './text-serializer';

/**
 * Serializes speaker notes (plain string or structured PptxTextBody) into an OpenXML `<p:notes>` document.
 */
export function serializeNotesSlide(notesInput: PptxTextBody | string): string {
  let txBodyNode: Record<string, unknown>;

  if (typeof notesInput === 'string') {
    const paragraphs = notesInput.split('\n').map((line) => ({
      properties: {},
      runs: [
        {
          properties: {},
          text: line,
        },
      ],
    }));
    txBodyNode = serializeTextBody({
      bodyProperties: {},
      paragraphs,
    });
  } else {
    txBodyNode = serializeTextBody(notesInput);
  }

  const builder = createXmlBuilder();
  const txBodyXml = builder.build({ 'p:txBody': txBodyNode });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Slide Image Placeholder 1"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="sldImg"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Notes Placeholder 2"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="body" idx="1"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        ${txBodyXml}
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:notes>`;
}

/**
 * Serializes the relationships for a notesSlide pointing back to its parent slide and notesMaster.
 */
export function serializeNotesSlideRels(slideNum: number): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${slideNum}.xml"/>
</Relationships>`;
}
