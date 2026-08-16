import { readFile, writeFile } from 'node:fs/promises';
import { relative, resolve } from 'node:path';
import { emuToInches, parsePptx, resolveSlideLayers } from '../dist/index.js';

/**
 * Inspection CLI tool for @hokkyss/pptx-reader.
 */
async function main() {
  const fileArg = process.argv[2];
  const samplePath = fileArg
    ? resolve(process.cwd(), fileArg)
    : resolve(process.cwd(), 'demo_generated.pptx');

  console.log('\n==================================================');
  console.log('      PPTX Parser Inspection CLI');
  console.log('==================================================\n');
  console.log(`📄 File: ${relative(process.cwd(), samplePath)}\n`);

  try {
    const buffer = await readFile(samplePath);
    const startTime = performance.now();
    const doc = await parsePptx(buffer);
    const endTime = performance.now();

    console.log(`⏱️  Parsed in ${(endTime - startTime).toFixed(2)} ms\n`);

    // 1. Metadata Summary
    console.log('📌 METADATA:');
    console.log(`  • Title:           ${doc.metadata.title || '(none)'}`);
    console.log(`  • Creator:         ${doc.metadata.creator || '(none)'}`);
    console.log(`  • Last Modified By:${doc.metadata.lastModifiedBy || '(none)'}`);
    console.log(`  • Revision:        ${doc.metadata.revision ?? '(none)'}`);
    console.log(`  • Slide Count:     ${doc.metadata.slideCount}`);
    console.log(`  • Slide Dimensions:${emuToInches(doc.metadata.slideWidth).toFixed(2)}" x ${emuToInches(doc.metadata.slideHeight).toFixed(2)}" (${doc.metadata.slideWidth} x ${doc.metadata.slideHeight} EMU)\n`);

    // 2. Themes & Color Scheme Summary
    console.log('🎨 THEMES:');
    doc.themes.forEach((t, i) => {
      console.log(`  Theme #${i + 1}: "${t.name}" (Major Font: ${t.fontScheme.majorFont}, Minor Font: ${t.fontScheme.minorFont})`);
      console.log(`    Accent1: #${t.colorScheme.accent1} | Accent2: #${t.colorScheme.accent2} | Accent3: #${t.colorScheme.accent3}`);
    });
    console.log('');

    // 3. Media Assets Summary
    console.log('🖼️  EMBEDDED MEDIA (ppt/media/*):');
    console.log(`  Total Media Files: ${doc.media.length}`);
    doc.media.slice(0, 5).forEach((m) => {
      const sizeKb = m.data ? (m.data.length / 1024).toFixed(1) : 'lazy';
      console.log(`  • ${m.filename} [${m.mimeType}] - ${sizeKb} KB`);
    });
    if (doc.media.length > 5) {
      console.log(`  ... and ${doc.media.length - 5} more media files`);
    }
    console.log('');

    // 4. Slide Masters & Layouts Summary
    console.log('📐 MASTERS & LAYOUTS:');
    console.log(`  • Slide Masters: ${doc.slideMasters.length} (${doc.slideMasters.map((m) => m.name || m.id).join(', ')})`);
    console.log(`  • Slide Layouts: ${doc.slideLayouts.length}`);
    console.log(`  • First 5 Layout Names: ${doc.slideLayouts.slice(0, 5).map((l) => `"${l.name}"`).join(', ')}\n`);

    // 5. Per-Slide Breakdown & Layer Inspection
    console.log('📑 SLIDE & LAYER BREAKDOWN:\n');

    doc.slides.slice(0, 3).forEach((slide) => {
      console.log(`--------------------------------------------------`);
      console.log(`Slide ${slide.slideNumber} (ID: ${slide.slideId}, Layout: ${slide.layoutId || 'default'})`);
      console.log(`--------------------------------------------------`);

      const layers = resolveSlideLayers(doc, slide.slideNumber);
      if (layers) {
        console.log(`  Layer Composition: Master (${layers.masterElements.length}) + Layout (${layers.layoutElements.length}) + Slide (${layers.slideElements.length}) = ${layers.allElementsInRenderOrder.length} total elements`);
      }

      console.log(`  Elements (${slide.elements.length}):`);
      slide.elements.forEach((el) => {
        const xIn = emuToInches(el.position.x).toFixed(2);
        const yIn = emuToInches(el.position.y).toFixed(2);
        const wIn = emuToInches(el.position.cx).toFixed(2);
        const hIn = emuToInches(el.position.cy).toFixed(2);

        const visState = el.isVisible ? '👁️ visible' : '🙈 hidden';
        const lockState = el.isLocked ? '🔒 locked' : '🔓 unlocked';

        let preview = '';
        if (el.textBody && el.textBody.paragraphs.length > 0) {
          const firstText = el.textBody.paragraphs[0].runs.map((r) => r.text).join('');
          preview = ` → "${firstText.slice(0, 45)}${firstText.length > 45 ? '...' : ''}"`;
        } else if (el.picture) {
          preview = ` → Picture (blip: ${el.blipEmbedId || el.picture.mediaId})`;
        } else if (el.chart) {
          preview = ` → Chart (${el.chart.chartType}, ${el.chart.series.length} series)`;
        } else if (el.table) {
          preview = ` → Table (${el.table.rows.length} rows x ${el.table.columnWidths.length} cols)`;
        }

        console.log(`    [zIndex:${el.zIndex}] ${el.name || el.type} (type: ${el.elementType || el.type}, pos: ${xIn}"x${yIn}", size: ${wIn}"x${hIn}", ${visState}, ${lockState})${preview}`);
      });
      if (slide.notes) {
        console.log(`  🎙️ Speaker Notes: "${slide.notes}"`);
      }
      console.log('');
    });

    if (doc.slides.length > 3) {
      console.log(`... and ${doc.slides.length - 3} more slides (total ${doc.slides.length} slides)\n`);
    }

    // 6. Write complete formatted JSON output to parsed_output.json
    const jsonPath = resolve(process.cwd(), 'parsed_output.json');
    // Prepare serializable copy omitting huge raw Uint8Array buffers for readable JSON
    const serializableDoc = {
      ...doc,
      media: doc.media.map((m) => ({
        byteLength: m.data ? m.data.length : 0,
        filename: m.filename,
        id: m.id,
        mimeType: m.mimeType,
        path: m.path,
      })),
    };

    await writeFile(jsonPath, JSON.stringify(serializableDoc, null, 2), 'utf8');
    console.log(`💾 Full parsed AST saved to: ${relative(process.cwd(), jsonPath)}\n`);
  } catch (err) {
    console.error('❌ Error parsing PPTX file:', err);
    process.exit(1);
  }
}

main();
