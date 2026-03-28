// scripts/build-git-deps.js
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const nodeModules = path.resolve(__dirname, '..', 'node_modules');

function run(cmd, cwd) {
  console.log(`\n> ${cmd} (cwd: ${cwd})`);
  execSync(cmd, { stdio: 'inherit', cwd });
}

function tryBuild(pkgPath) {
  const pkgJsonPath = path.join(pkgPath, 'package.json');
  if (!fs.existsSync(pkgJsonPath)) return false;
  try {
    const pkg = require(pkgJsonPath);
    // skip if already has a built entry file present
    const main = pkg.main || pkg.module || pkg.browser;
    if (main) {
      const mainPath = path.join(pkgPath, main);
      if (fs.existsSync(mainPath)) {
        return false; // already has built output
      }
    }
    // prefer package's build script
    const buildScript = pkg.scripts && (pkg.scripts.build || pkg.scripts.prepare || pkg.scripts.prepublishOnly);
    if (!buildScript) return false;

    // install deps and run build
    if (fs.existsSync(path.join(pkgPath, 'package.json'))) {
      // some repos include their own node_modules; but safe to run npm ci
      run('npm install --no-audit --no-fund --silent', pkgPath);
      run('npm run build', pkgPath);
      // run prepare or prepublishOnly if no build script produced dist
      if (!fs.existsSync(mainPath) && pkg.scripts.prepare) run('npm run prepare', pkgPath);
      if (!fs.existsSync(mainPath) && pkg.scripts.prepublishOnly) run('npm run prepublishOnly', pkgPath);
      return true;
    }
  } catch (e) {
    console.warn(`Skipping ${pkgPath}: ${e.message}`);
  }
  return false;
}

function walkAndBuild(base) {
  if (!fs.existsSync(base)) return;
  const entries = fs.readdirSync(base, { withFileTypes: true });
  for (const e of entries) {
    // handle scoped packages @scope/name
    if (e.isDirectory() && e.name.startsWith('@')) {
      walkAndBuild(path.join(base, e.name));
      continue;
    }
    if (!e.isDirectory()) continue;
    const pkgPath = path.join(base, e.name);
    // only attempt build for packages likely installed from GitHub:
    const gitmeta = path.join(pkgPath, '.git');
    const pkgjson = path.join(pkgPath, 'package.json');
    if (fs.existsSync(pkgjson)) {
      const built = tryBuild(pkgPath);
      if (!built) {
        // attempt to build nested workspace packages (monorepos)
        const packagesDir = ['packages', 'dist', 'projects'];
        for (const dir of packagesDir) {
          const candidate = path.join(pkgPath, dir);
          if (fs.existsSync(candidate)) {
            const subs = fs.readdirSync(candidate, { withFileTypes: true });
            for (const s of subs) {
              if (s.isDirectory()) tryBuild(path.join(candidate, s.name));
            }
          }
        }
      }
    }
  }
}

walkAndBuild(nodeModules);
console.log('\nDone building git-installed deps.');
