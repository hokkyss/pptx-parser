import { createServerFn } from '@tanstack/react-start';
import { renderServerComponent } from '@tanstack/react-start/rsc';
import MarkdownRenderer from '../../../components/markdown-renderer.component';
import { renderTeaserRequestDto } from '../dto/render-teaser.dto';

const CODE_TEASER_SNIPPET = `\`\`\`typescript
import { Presentation, inches, points } from '@hokkyss/pptx';

// 1. Initialize presentation
const pres = Presentation.create({ title: 'Cloud Architecture' });

// 2. Add slide with theme & multilevel bullets
const slide = pres.addSlide();
slide.setBackground('0F172A');

slide.addText('Distributed Edge Compute', {
  color: '38BDF8',
  fontSize: points(32),
  bold: true,
  x: inches(1),
  y: inches(1),
  w: inches(11),
  h: inches(0.8)
});

// 3. Attach connected cards
slide.addShape('roundRect', {
  id: 'gateway',
  text: 'Edge Gateway',
  x: inches(1),
  y: inches(2.5),
  w: inches(3),
  h: inches(1.5)
});

slide.addShape('roundRect', {
  id: 'auth',
  text: 'Auth Service',
  x: inches(6),
  y: inches(2.5),
  w: inches(3),
  h: inches(1.5)
});

slide.addConnector({
  from: {
    shapeId: 'gateway',
    position: 'right'
  },
  to: {
    shapeId: 'auth',
    position: 'left'
  },
  endArrow: 'triangle'
});

// 4. Save (or toArrayBuffer() in browser/workers)
await pres.save('deck.pptx');
\`\`\``;

const renderTeaserFunction = createServerFn({ method: 'GET' })
  .validator(renderTeaserRequestDto)
  .handler(async () => {
    return renderServerComponent(
      <MarkdownRenderer content={CODE_TEASER_SNIPPET} />,
    );
  });

export default renderTeaserFunction;
