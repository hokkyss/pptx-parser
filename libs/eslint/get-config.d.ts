import { Config } from "eslint/config";

export declare interface MonorepoEslintOptions {
    /**
     * Output directory for eslint to ignore
     */
    outDir: string;
    /**
     * Enable tanstack router eslint config.
     */
    tanstackRouter?: boolean;
    /**
     * Enable tanstack query eslint config.
     */
    tanstackQuery?: boolean;
    /**
     * Enable react eslint config.
     */
    react?: boolean;
    /**
     * The globals environment that would like to be used.
     * 
     * @default "isomorphic"
     */
    environment?: "browser" | "node" | "isomorphic";
    /**
     * Root directory of tsconfig.json file
     */
    tsconfigRootDir: string;
    /**
     * Custom tsconfig file path(s) relative to tsconfigRootDir.
     * E.g. "./tsconfig.test.json" or ["./tsconfig.json", "./tsconfig.test.json"]
     */
    tsconfigPath?: string | string[];
    /**
     * Custom project setting for typescript-eslint parserOptions.
     */
    project?: string | string[] | boolean;
    /**
     * Custom projectService setting for typescript-eslint parserOptions.
     */
    projectService?: boolean | Record<string, unknown>;
    /**
     * Additional ignore patterns
     */
    ignores?: string[];
}

export default function getConfig(opts: MonorepoEslintOptions): Config[];
