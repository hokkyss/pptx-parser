import type {
  ContainerDirective,
  LeafDirective,
  TextDirective,
} from 'mdast-util-directive';
import type { Plugin } from 'unified';
import type { Node } from 'unist';
import { visit } from 'unist-util-visit';

export interface TabItemMeta {
  label: string;
  value: string;
}

/**
 * Unified remark plugin to transform container directives (:::tabs / :::tab) into structured HTML elements.
 * Generates <div className="markdown-tabs" data-tabs="..." data-sync-key="..." data-default-value="...">
 * containing <div className="markdown-tab-panel" data-tab-value="..." data-tab-label="...">.
 * @returns Unified transformer function
 */
export const remarkTabs: Plugin = () => {
  return (tree, file) => {
    visit(tree, (node: Node) => {
      // We handle only container directives and text directives.
      // Content directive is below
      // ::: name [inline-content] {key=val}
      // <the content>
      // :::
      //
      // Text/inline directive is below
      // :name[content]{key=value}
      if (node.type !== 'containerDirective' && node.type !== 'textDirective' && node.type !== 'leafDirective') {
        return;
      }

      const directiveNode = node as ContainerDirective | LeafDirective | TextDirective;
      const data = directiveNode.data || (directiveNode.data = {});
      const attributes = directiveNode.attributes || {};

      /**
       * Reads up `:tab-item[label]{value="<stringvalue>"}`
       */
      if (directiveNode.type === 'textDirective') {
        if (directiveNode.name === 'tab-item') {
          if (!attributes.value) {
            file.fail(`tab-item must have value attribute`, node);
          }

          data.hName = 'div';
          data.hProperties = {
            className: ['markdown-tab-item'],
            'data-value': attributes.value,
          };
          return;
        }
        file.fail(`Invalid text/inline directive: "${directiveNode.name}"`, node);
      }

      if (directiveNode.type === 'containerDirective') {
        /**
         * Reads up `:::tabs{sync="name",defaultValue="<value>"}`
         *
         * If for some reason "value" is not a valid tab, that will be rendered as is without any validation.
         */
        if (directiveNode.name === 'tabs-root') {
          const syncKey = attributes.sync;
          const defaultValue = attributes.defaultValue;

          if (!defaultValue) {
            file.fail(`No defaultValue for tabs ${directiveNode.name}`, node);
          }

          data.hName = 'div';
          data.hProperties = {
            className: ['markdown-tabs'],
            'data-default-value': defaultValue,
            'data-sync-key': syncKey,
          };
          return;
        }

        if (directiveNode.name === 'tabs-list') {
          data.hName = 'div';
          data.hProperties = {
            className: ['markdown-tabs-list'],
          };
          return;
        }

        if (directiveNode.name === 'tab-content') {
          if (!attributes.value) {
            file.fail(`No value for ${directiveNode.name}`, node);
          }

          data.hName = 'div';
          data.hProperties = {
            className: ['markdown-tabs-content'],
            'data-value': attributes.value,
          };
          return;
        }

        file.fail(`Invalid container directive "${directiveNode.name}"`, node);
      }

      file.fail(`Invalid leaf directive: "${directiveNode.name}"`, node);
    });
  };
};

export default remarkTabs;
