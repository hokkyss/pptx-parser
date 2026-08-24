import { describe, expect, it } from 'vitest';
import { parseHyperlink } from '../../lib/parsers/hyperlink-parser';
import { parseShapes } from '../../lib/parsers/shape-parser';
import { parseRunProperties } from '../../lib/parsers/text-parser';
import { createRelationshipResolver } from '../../lib/resolvers/relationship-resolver';

describe('Hyperlink Parser (@hokkyss/pptx-reader)', () => {
  it('should parse external URL hyperlinks from relationship resolver', () => {
    const relsXml = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://hokkyss.dev" TargetMode="External"/>
      </Relationships>
    `;
    const resolver = createRelationshipResolver(relsXml, 'ppt/slides/slide1.xml');

    const hlinkNode = {
      '@_r:id': 'rId2',
      '@_tooltip': 'Visit Portfolio',
    };

    const parsed = parseHyperlink(hlinkNode, resolver);
    expect(parsed).toBeDefined();
    expect(parsed?.url).toBe('https://hokkyss.dev');
    expect(parsed?.tooltip).toBe('Visit Portfolio');
    expect(parsed?.rId).toBe('rId2');
  });

  it('should parse internal slide jumps from relationship resolver', () => {
    const relsXml = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide4.xml"/>
      </Relationships>
    `;
    const resolver = createRelationshipResolver(relsXml, 'ppt/slides/slide1.xml');

    const hlinkNode = {
      '@_action': 'ppaction://hlinksldjump',
      '@_r:id': 'rId3',
      '@_tooltip': 'Go to Slide 4',
    };

    const parsed = parseHyperlink(hlinkNode, resolver);
    expect(parsed).toBeDefined();
    expect(parsed?.slideIndex).toBe(4);
    expect(parsed?.tooltip).toBe('Go to Slide 4');
  });

  it('should parse slide show jump actions', () => {
    const nextParsed = parseHyperlink({ '@_action': 'ppaction://hlinkshowjump?jump=nextslide' });
    expect(nextParsed?.action).toBe('nextSlide');

    const prevParsed = parseHyperlink({ '@_action': 'ppaction://hlinkshowjump?jump=previousslide' });
    expect(prevParsed?.action).toBe('previousSlide');

    const firstParsed = parseHyperlink({ '@_action': 'ppaction://hlinkshowjump?jump=firstslide' });
    expect(firstParsed?.action).toBe('firstSlide');

    const lastParsed = parseHyperlink({ '@_action': 'ppaction://hlinkshowjump?jump=lastslide' });
    expect(lastParsed?.action).toBe('lastSlide');

    const endParsed = parseHyperlink({ '@_action': 'ppaction://hlinkshowjump?jump=endshow' });
    expect(endParsed?.action).toBe('endShow');
  });

  it('should extract hyperlink in parseRunProperties', () => {
    const relsXml = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:hokki@example.com" TargetMode="External"/>
      </Relationships>
    `;
    const resolver = createRelationshipResolver(relsXml, 'ppt/slides/slide1.xml');

    const rPr = {
      '@_b': '1',
      'a:hlinkClick': {
        '@_r:id': 'rId5',
        '@_tooltip': 'Send Email',
      },
    };

    const parsed = parseRunProperties(rPr, resolver);
    expect(parsed.bold).toBe(true);
    expect(parsed.hyperlink).toEqual({
      action: undefined,
      rId: 'rId5',
      slideIndex: undefined,
      tooltip: 'Send Email',
      url: 'mailto:hokki@example.com',
    });
  });

  it('should extract shape-level and picture-level hyperlinks in parseShapes', () => {
    const relsXml = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://google.com" TargetMode="External"/>
        <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://github.com" TargetMode="External"/>
      </Relationships>
    `;
    const resolver = createRelationshipResolver(relsXml, 'ppt/slides/slide1.xml');

    const slideXml = `
      <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:cSld>
          <p:spTree>
            <p:nvGrpSpPr>
              <p:cNvPr id="1" name=""/>
              <p:cNvGrpSpPr/>
              <p:nvPr/>
            </p:nvGrpSpPr>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="LinkShape">
                  <a:hlinkClick r:id="rId6" tooltip="Google"/>
                </p:cNvPr>
                <p:cNvSpPr/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="1000" cy="1000"/>
                </a:xfrm>
              </p:spPr>
            </p:sp>
            <p:pic>
              <p:nvPicPr>
                <p:cNvPr id="3" name="LinkPic">
                  <a:hlinkClick r:id="rId7" tooltip="GitHub"/>
                </p:cNvPr>
                <p:cNvPicPr/>
                <p:nvPr/>
              </p:nvPicPr>
              <p:blipFill>
                <a:blip r:embed="rId8"/>
              </p:blipFill>
              <p:spPr>
                <a:xfrm>
                  <a:off x="1000" y="1000"/>
                  <a:ext cx="2000" cy="2000"/>
                </a:xfrm>
              </p:spPr>
            </p:pic>
          </p:spTree>
        </p:cSld>
      </p:sld>
    `;

    const shapes = parseShapes(slideXml, resolver);
    expect(shapes).toHaveLength(2);
    expect(shapes[0].hyperlink).toEqual({
      action: undefined,
      rId: 'rId6',
      slideIndex: undefined,
      tooltip: 'Google',
      url: 'https://google.com',
    });
    expect(shapes[1].hyperlink).toEqual({
      action: undefined,
      rId: 'rId7',
      slideIndex: undefined,
      tooltip: 'GitHub',
      url: 'https://github.com',
    });
  });

  it('neutralizes malicious XSS and traversal targets during relationship parsing', () => {
    const relsXml = `
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdX1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>
        <Relationship Id="rIdX2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../../etc/passwd"/>
      </Relationships>
    `;
    const resolver = createRelationshipResolver(relsXml, 'ppt/slides/slide1.xml');

    const maliciousNode1 = {
      '@_r:id': 'rIdX1',
      '@_tooltip': 'Malicious\r\nScreenTip\0Attack',
    };
    const parsed1 = parseHyperlink(maliciousNode1, resolver);
    expect(parsed1?.url).toBeUndefined();
    expect(parsed1?.tooltip).toBe('MaliciousScreenTipAttack');

    const maliciousNode2 = {
      '@_r:id': 'rIdX2',
    };
    const parsed2 = parseHyperlink(maliciousNode2, resolver);
    expect(parsed2?.slideIndex).toBeUndefined();
  });
});

describe('parseHyperlink edge cases', () => {
  it('handles custom action and empty node', () => {
    expect(parseHyperlink({})).toBeUndefined();
    expect(parseHyperlink(undefined)).toBeUndefined();
    const custom = parseHyperlink({ '@_action': 'ppaction://customAction' });
    expect(custom?.action).toBe('ppaction://customAction');
    const invalidUrl = parseHyperlink({ '@_invalidUrl': 'https://example.com' });
    expect(invalidUrl?.url).toBe('https://example.com');
  });
});
