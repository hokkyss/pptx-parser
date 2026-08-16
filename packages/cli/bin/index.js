#!/usr/bin/env node
import { Command } from 'commander';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const program = new Command();

program
  .name('boil')
  .description('Scaffolding CLI for the monorepo')
  .version('1.0.0');

/**
 * 
 * @param {string} templateName 
 * @param {string} destName 
 * @param {'package'} type 
 */
function copyTemplate(templateName, destName, type) {
  const srcDir = path.join(__dirname, '..', 'templates', templateName);
  
  let rootDir = __dirname;
  while (!fs.existsSync(path.join(rootDir, 'pnpm-workspace.yaml'))) {
    const parent = path.dirname(rootDir);
    if (parent === rootDir) {
      console.error('Could not find pnpm workspace root (missing pnpm-workspace.yaml)');
      process.exit(1);
    }
    rootDir = parent;
  }
  const destDir = path.join(rootDir, 'packages', destName);

  if (!fs.existsSync(srcDir)) {
    console.error(`Template ${templateName} not found at ${srcDir}`);
    process.exit(1);
  }

  if (fs.existsSync(destDir)) {
    console.error(`Destination ${destDir} already exists.`);
    process.exit(1);
  }

  fs.mkdirSync(destDir, { recursive: true });

  /**
   * 
   * @param {string} currentSrc 
   * @param {string} currentDest 
   */
  function walkAndCopy(currentSrc, currentDest) {
    const items = fs.readdirSync(currentSrc);
    for (const item of items) {
      if (item === 'node_modules' || item === 'dist' || item === '.turbo') continue;
      const srcPath = path.join(currentSrc, item);
      const destPath = path.join(currentDest, item);

      const stat = fs.statSync(srcPath);
      if (stat.isDirectory()) {
        fs.mkdirSync(destPath);
        walkAndCopy(srcPath, destPath);
      } else {
        // Only read as utf8 if we are sure it's not a binary, but for simple template it's mostly text
        if (srcPath.endsWith('.png') || srcPath.endsWith('.ico') || srcPath.endsWith('.jpg')) {
          fs.copyFileSync(srcPath, destPath);
        } else {
          let content = fs.readFileSync(srcPath, 'utf8');
          // Replace template name references
          content = content.replace(/@monorepo\/template/g, `@monorepo/${destName}`);
          content = content.replace(/packages\/template/g, `packages/${destName}`);
          content = content.replace(/monorepo-template/g, destName);
          fs.writeFileSync(destPath, content);
        }
      }
    }
  }

  walkAndCopy(srcDir, destDir);
  console.log(`Successfully created ${type} ${destName} at ${destDir}`);
}

program
  .command('package')
  .description('Scaffold a new package')
  .argument('<name>', 'Name of the package')
  .action((name) => {
    copyTemplate('package', name, 'package');
  });

program.parse();
