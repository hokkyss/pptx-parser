import type { Plugin } from 'unified';
import type { Node, Parent } from 'unist';
import { visit } from 'unist-util-visit';

export interface TabItemMeta {
  label: string;
  value: string;
}

interface DirectiveNode extends Parent {
  attributes?: Record<string, string>;
  data?: {
    hName?: string;
    hProperties?: Record<string, unknown>;
  };
  name?: string;
  type: string;
}

/**
 * Slugifies a label into a valid HTML / React tab identifier.
 * @param str Label string
 * @returns Clean slug identifier
 */
function slugify(str: string): string {
  return str.toLowerCase().replace(/[^a-z0-9_-]+/g, '-').replace(/(^-|-$)/g, '');
}

/**
 * Unified remark plugin to transform container directives (:::tabs / :::tab) into structured HTML elements.
 * Generates <div className="markdown-tabs" data-tabs="..." data-sync-key="..." data-default-value="...">
 * containing <div className="markdown-tab-panel" data-tab-value="..." data-tab-label="...">.
 * @returns Unified transformer function
 */
export const remarkTabs: Plugin = () => {
  return (tree: Node) => {
    visit(tree, (node: Node) => {
      const directive = node as DirectiveNode;

      if (
        (directive.type === 'containerDirective' || directive.type === 'leafDirective')
        && directive.name === 'tabs'
      ) {
        const attributes = directive.attributes || {};
        const syncKey = attributes.sync || attributes['data-sync'] || attributes['sync-key'] || undefined;
        let defaultValue = attributes.default || attributes['default-value'] || undefined;

        const tabsMeta: TabItemMeta[] = [];

        if (Array.isArray(directive.children)) {
          directive.children.forEach((childNode: Node, index: number) => {
            const childDirective = childNode as DirectiveNode;

            if (
              (childDirective.type === 'containerDirective' || childDirective.type === 'leafDirective')
              && childDirective.name === 'tab'
            ) {
              const childAttrs = childDirective.attributes || {};
              const label = childAttrs.label || childAttrs.title || childAttrs.name || `Tab ${index + 1}`;
              const value = childAttrs.value || slugify(label) || `tab-${index}`;

              if (index === 0 && !defaultValue) {
                defaultValue = value;
              }

              tabsMeta.push({ label, value });

              childDirective.data = {
                hName: 'div',
                hProperties: {
                  className: 'markdown-tab-panel',
                  'data-tab-label': label,
                  'data-tab-value': value,
                },
              };
            }
          });
        }

        directive.data = {
          hName: 'div',
          hProperties: {
            className: 'markdown-tabs',
            'data-default-value': defaultValue || (tabsMeta[0]?.value ?? 'tab-0'),
            'data-sync-key': syncKey,
            'data-tabs': JSON.stringify(tabsMeta),
          },
        };
      }
    });
  };
};

export default remarkTabs;
