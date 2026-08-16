import { describe, it, expect } from 'vitest';
import { parseShapes } from '../../lib/parsers/shape-parser';

describe('Slide Parsing (Shapes & Text)', () => {
  it('should extract 2 shapes (title + subtitle) from slide 1 XML', () => {
    const xml = `
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Title 1"/>
                <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm><a:off x="100" y="200"/><a:ext cx="500" cy="300"/></a:xfrm>
              </p:spPr>
              <p:txBody>
                <a:bodyPr/>
                <a:p>
                  <a:r><a:t>SPA vs RSC</a:t></a:r>
                </a:p>
              </p:txBody>
            </p:sp>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="3" name="Subtitle 2"/>
                <p:nvPr><p:ph type="subTitle"/></p:nvPr>
              </p:nvSpPr>
              <p:txBody>
                <a:bodyPr/>
                <a:p>
                  <a:r><a:t>Architecture &amp; Tradeoffs</a:t></a:r>
                </a:p>
              </p:txBody>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sld>
    `;

    const shapes = parseShapes(xml);

    expect(shapes.length).toBe(2);
    expect(shapes[0].name).toBe('Title 1');
    expect(shapes[0].type).toBe('shape');
    expect(shapes[0].textBody?.paragraphs[0].runs[0].text).toBe('SPA vs RSC');
    expect(shapes[0].position.x).toBe(100);
    expect(shapes[0].position.y).toBe(200);

    expect(shapes[1].name).toBe('Subtitle 2');
    expect(shapes[1].textBody?.paragraphs[0].runs[0].text).toBe('Architecture & Tradeoffs');
  });

  it('should parse picture shapes blip embed ID', () => {
    const xml = `
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:cSld>
          <p:spTree>
            <p:pic>
              <p:nvPicPr>
                <p:cNvPr id="10" name="Picture 1"/>
              </p:nvPicPr>
              <p:blipFill>
                <a:blip r:embed="rId5"/>
              </p:blipFill>
            </p:pic>
          </p:spTree>
        </p:cSld>
      </p:sld>
    `;

    const shapes = parseShapes(xml);
    expect(shapes.length).toBe(1);
    expect(shapes[0].type).toBe('picture');
    expect(shapes[0].blipEmbedId).toBe('rId5');
  });
});
