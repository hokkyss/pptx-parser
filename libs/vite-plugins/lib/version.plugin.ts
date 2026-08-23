import type { Plugin } from 'vite';

/**
 *
 * @param version
 */
export default function viteVersionPlugin(version: string): Plugin {
  const ENTRY_NAME = 'virtual:vite-plugin-version';
  const RESOLVED_ENTRY_NAME = '\0' + ENTRY_NAME;

  return {
    apply: () => true,
    config() {
      return {
        build: {
          lib: {
            entry: {
              version: ENTRY_NAME,
            },
          },
        },
      };
    },
    enforce: 'pre',
    generateBundle() {
      this.emitFile({
        code: `
declare const version: string;
export default version;
`,
        fileName: 'version.d.ts',
        name: 'version.d.ts',
        type: 'prebuilt-chunk',
      });
    },
    // 2. Inject the dynamic module content at build/run time
    load(id) {
      if (id === RESOLVED_ENTRY_NAME) {
        return `
            export default ${JSON.stringify(version)};
          `;
      }
    },
    name: '@hokkyss/vite-plugin-version',
    resolveId(id) {
      if (id.endsWith(ENTRY_NAME)) {
        return RESOLVED_ENTRY_NAME;
      }
    },
  };
}
