"""
Docx Master - Cleanup intermediate files after docx generation.
Usage: python cleanup.py <working_dir>

Cleans:
  1. Python generation scripts (gen_report.py, gen_*.py) — intermediate build scripts
  2. Temp JSON files (*_content.json, *_temp.json, template.json) — generation intermediates
  3. node_modules/ package.json package-lock.json — from npm install jszip bootstrap
  4. Word lock files (~$*.docx, ~$*.doc) — Word/OOXML temp lock artifacts
  5. Empty __pycache__ and .pyc files (optional, with --pycache)
"""
import os, sys, shutil, fnmatch, argparse

def is_intermediate_py(filepath):
    """Check if a .py file looks like a one-shot generation script."""
    basename = os.path.basename(filepath)
    patterns = ['gen_report.py', 'gen_*.py', 'generate_*.py', 'make_doc*.py']
    for p in patterns:
        if fnmatch.fnmatch(basename, p):
            try:
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    head = f.read(500)
                    if 'from docx import Document' in head or 'Document(' in head or 'JSZip' in head:
                        return True
            except:
                pass
    return False

def cleanup(workdir, dry_run=False, remove_pycache=False):
    removed = []
    skipped = []

    items = os.listdir(workdir)

    for item in items:
        full = os.path.join(workdir, item)

        # 1. Word lock files
        if item.startswith('~$') and (item.endswith('.docx') or item.endswith('.doc')):
            if not dry_run:
                try:
                    os.remove(full)
                    removed.append(full)
                except:
                    skipped.append(full)
            else:
                removed.append(full)
            continue

        # 2. node_modules from npm install
        if item == 'node_modules' and os.path.isdir(full):
            if not dry_run:
                try:
                    shutil.rmtree(full)
                    removed.append(full)
                except:
                    skipped.append(full)
            else:
                removed.append(full)
            continue

        # 3. package.json / package-lock.json from npm install
        if item in ('package.json', 'package-lock.json'):
            # Only delete if it looks like a npm bootstrap artifact (not a real project)
            try:
                import json
                with open(full, 'r') as f:
                    data = json.load(f)
                # If only has jszip or has no other real deps, it's a bootstrap artifact
                deps = data.get('dependencies', {})
                if item == 'package.json' and (not deps or deps == {'jszip': '^3.10.1'} or list(deps.keys()) == ['jszip']):
                    if not dry_run:
                        os.remove(full)
                        removed.append(full)
                    else:
                        removed.append(full)
                else:
                    skipped.append(full)
            except:
                skipped.append(full)
            continue

        # 4. Python generation scripts
        if item.endswith('.py') and os.path.isfile(full):
            if is_intermediate_py(full):
                if not dry_run:
                    os.remove(full)
                    removed.append(full)
                else:
                    removed.append(full)
            else:
                skipped.append(full)
            continue

        # 5. Temp JSON files (generation intermediates)
        if item.endswith('.json') and os.path.isfile(full):
            temp_patterns = ['*_content.json', '*_temp.json', 'template.json']
            for p in temp_patterns:
                if fnmatch.fnmatch(item, p):
                    if not dry_run:
                        try:
                            os.remove(full)
                            removed.append(full)
                        except:
                            skipped.append(full)
                    else:
                        removed.append(full)
                    break
            else:
                skipped.append(full)
            continue

    # 6. pycache cleanup (optional)
    if remove_pycache:
        for root, dirs, files in os.walk(workdir):
            for d in dirs:
                if d == '__pycache__':
                    full = os.path.join(root, d)
                    if not dry_run:
                        shutil.rmtree(full)
                        removed.append(full)
                    else:
                        removed.append(full)

    return removed, skipped

def main():
    parser = argparse.ArgumentParser(description='Cleanup docx intermediate files')
    parser.add_argument('workdir', help='Working directory to clean')
    parser.add_argument('--dry-run', action='store_true', help='Show what would be removed')
    parser.add_argument('--pycache', action='store_true', help='Also remove __pycache__ dirs')
    args = parser.parse_args()

    workdir = os.path.abspath(args.workdir)
    if not os.path.isdir(workdir):
        print(f'ERROR: {workdir} is not a directory')
        sys.exit(1)

    removed, skipped = cleanup(workdir, dry_run=args.dry_run, remove_pycache=args.pycache)

    if args.dry_run:
        print(f'[DRY RUN] Would remove {len(removed)} files:')
    else:
        print(f'Removed {len(removed)} files:')

    for r in removed:
        print(f'  ✓ {os.path.basename(r)}')
    for s in skipped:
        print(f'  - skipped: {os.path.basename(s)}')

    print(f'\nTotal: {len(removed)} removed, {len(skipped)} skipped')
    return 0

if __name__ == '__main__':
    sys.exit(main())
