import { execSync } from 'node:child_process';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';

const PKG_PATHS = [
  'packages/pptx-core',
  'packages/pptx-reader',
  'packages/pptx-writer',
  'packages/pptx',
];

function bumpSemver(version, type) {
  const parts = version.split('.').map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) {
    throw new Error(`Invalid semver string: ${version}`);
  }
  let [major, minor, patch] = parts;
  if (type === 'major') return `${major + 1}.0.0`;
  if (type === 'minor') return `${major}.${minor + 1}.0`;
  if (type === 'patch') return `${major}.${minor}.${patch + 1}`;
  return type; // explicit semver
}

function getLatestTag() {
  try {
    const tag = execSync('git describe --tags --abbrev=0', { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] }).trim();
    return tag;
  } catch {
    return null;
  }
}

function getCommitsSinceTag(tag) {
  const range = tag ? `${tag}..HEAD` : 'HEAD';
  try {
    const log = execSync(`git log ${range} --pretty=format:"%s|%h|%an"`, { encoding: 'utf8' }).trim();
    if (!log) return [];
    return log.split('\n').map((line) => {
      const [subject, hash, author] = line.split('|');
      return { subject, hash, author };
    });
  } catch {
    return [];
  }
}

function categorizeCommits(commits) {
  const groups = {
    Features: [],
    'Bug Fixes': [],
    Performance: [],
    Refactoring: [],
    Documentation: [],
    Maintenance: [],
  };

  for (const c of commits) {
    const sub = c.subject.trim();
    if (sub.startsWith('feat')) groups.Features.push(c);
    else if (sub.startsWith('fix')) groups['Bug Fixes'].push(c);
    else if (sub.startsWith('perf')) groups.Performance.push(c);
    else if (sub.startsWith('refactor')) groups.Refactoring.push(c);
    else if (sub.startsWith('docs')) groups.Documentation.push(c);
    else groups.Maintenance.push(c);
  }

  return groups;
}

function generateReleaseNotes(version, categorizedCommits) {
  const dateStr = new Date().toISOString().split('T')[0];
  let notes = `## v${version} (${dateStr})\n\n`;

  const emojis = {
    Features: '🚀 Features',
    'Bug Fixes': '🐛 Bug Fixes',
    Performance: '⚡ Performance & Optimization',
    Refactoring: '♻️ Refactoring',
    Documentation: '📝 Documentation',
    Maintenance: '🔧 Tooling & Chores',
  };

  let hasContent = false;
  for (const [groupName, commits] of Object.entries(categorizedCommits)) {
    if (commits.length > 0) {
      hasContent = true;
      notes += `### ${emojis[groupName] || groupName}\n\n`;
      for (const c of commits) {
        notes += `- ${c.subject} (\`${c.hash}\`)\n`;
      }
      notes += '\n';
    }
  }

  if (!hasContent) {
    notes += `- General maintenance and ecosystem stability updates.\n\n`;
  }

  return notes;
}

function updateChangelog(filePath, newNotes) {
  let currentContent = '';
  if (existsSync(filePath)) {
    currentContent = readFileSync(filePath, 'utf8');
  } else {
    currentContent = '# Changelog\n\nAll notable changes to this project will be documented in this file.\n\n';
  }

  // Prepend new release notes below header
  const headerMatch = currentContent.match(/^# [^\n]+\n\n(?:[^\n]+\n\n)?/);
  if (headerMatch) {
    const header = headerMatch[0];
    const rest = currentContent.slice(header.length);
    writeFileSync(filePath, `${header}${newNotes}${rest}`);
  } else {
    writeFileSync(filePath, `# Changelog\n\n${newNotes}${currentContent}`);
  }
}

async function runRelease() {
  const args = process.argv.slice(2);
  const isDryRun = args.includes('--dry-run');

  // Parse arguments
  let versionInput = 'patch';
  const typeIdx = args.indexOf('--type');
  if (typeIdx !== -1 && args[typeIdx + 1]) versionInput = args[typeIdx + 1];

  const verIdx = args.indexOf('--version');
  if (verIdx !== -1 && args[verIdx + 1]) versionInput = args[verIdx + 1];

  // Get current main package version
  const mainPkgJsonPath = resolve(process.cwd(), 'packages/pptx/package.json');
  const mainPkg = JSON.parse(readFileSync(mainPkgJsonPath, 'utf8'));
  const currentVersion = mainPkg.version || '0.1.0';

  const newVersion = ['major', 'minor', 'patch'].includes(versionInput)
    ? bumpSemver(currentVersion, versionInput)
    : versionInput;

  console.log(`\n========================================================================================`);
  console.log(` 🚀 RELEASING MONOREPO: v${currentVersion} ➔ v${newVersion} ${isDryRun ? '(DRY RUN)' : ''}`);
  console.log(`========================================================================================\n`);

  const latestTag = getLatestTag();
  console.log(`📌 Previous tag: ${latestTag || 'none (first release)'}`);
  const commits = getCommitsSinceTag(latestTag);
  console.log(`📝 Total commits since last tag: ${commits.length}`);

  const categorized = categorizeCommits(commits);
  const releaseNotes = generateReleaseNotes(newVersion, categorized);

  if (isDryRun) {
    console.log('\n--- Release Notes Preview ---\n');
    console.log(releaseNotes);
    console.log('--- End Preview ---\n');
    return;
  }

  // 1. Update all package.json files
  console.log(`📦 Bumping version to ${newVersion} across all packages...`);
  for (const pkgRelPath of PKG_PATHS) {
    const pkgPath = resolve(process.cwd(), pkgRelPath, 'package.json');
    if (existsSync(pkgPath)) {
      const content = JSON.parse(readFileSync(pkgPath, 'utf8'));
      content.version = newVersion;
      writeFileSync(pkgPath, JSON.stringify(content, null, 2) + '\n');
      console.log(`  ✓ Updated ${content.name}@${newVersion}`);
    }
  }

  // 2. Update Root CHANGELOG.md and package CHANGELOG.md
  console.log(`📑 Updating CHANGELOG.md...`);
  const rootChangelog = resolve(process.cwd(), 'CHANGELOG.md');
  updateChangelog(rootChangelog, releaseNotes);

  // Write release notes artifact for GitHub Release action
  const releaseNotesPath = resolve(process.cwd(), 'benchmarks/release-notes.md');
  writeFileSync(releaseNotesPath, releaseNotes);
  console.log(`💾 Release notes written to: ${releaseNotesPath}`);

  console.log(`\n✨ Version bump and changelog generation complete!`);
}

runRelease().catch(console.error);
