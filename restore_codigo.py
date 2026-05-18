#!/usr/bin/env python3
"""
restore_codigo.py — Restaura src/Código.js desde commit 4dab482 con Hub Tasks agregado.

Uso (desde la raiz del repositorio):
    python3 restore_codigo.py
    git add "src/Código.js"
    git commit -m "fix: restaurar Código.js original + agregar Hub Tasks al final"
    git push
"""

import subprocess, sys, os

BASE_COMMIT = "4dab482397fcc4fef5a7ca720ab23a8fa4e78f70"
TARGET_FILE = "src/Código.js"

def run_git(*args):
    result = subprocess.run(["git"] + list(args), capture_output=True, text=True, encoding="utf-8")
    if result.returncode != 0:
        print(f"ERROR: git {' '.join(args)}")
        print(result.stderr)
        sys.exit(1)
    return result.stdout

def main():
    # 1. Get base content from commit 4dab482
    print(f"Getting {TARGET_FILE} from commit {BASE_COMMIT}...")
    content = run_git("show", f"{BASE_COMMIT}:{TARGET_FILE}")
    print(f"Base file: {len(content)} chars")

    # 2. Modify getCurrentUserInfo to include team field
    OLD = "return { email: email, isAdmin: _isAdmin_(email) };"
    NEW = "return { email: email, isAdmin: _isAdmin_(email), team: _getTeamForUser_(email) };"
    if OLD in content:
        content = content.replace(OLD, NEW)
        print("getCurrentUserInfo patched (added team field)")
    else:
        print("WARNING: getCurrentUserInfo pattern not found - manual fix required")

    # 3. Append Hub Tasks section
    script_dir = os.path.dirname(os.path.abspath(__file__))
    hub_tasks = open(os.path.join(script_dir, "hub_tasks_append.js"), "r", encoding="utf-8").read()
    content += hub_tasks
    print(f"Hub Tasks appended. Final file: {len(content)} chars")

    # 4. Write result
    with open(TARGET_FILE, "w", encoding="utf-8") as f:
        f.write(content)

    # 5. Verify
    assert "function doGet" in content, "FAIL: doGet missing!"
    assert "function getHubTasks" in content, "FAIL: getHubTasks missing!"
    assert "function saveHubTask" in content, "FAIL: saveHubTask missing!"
    assert "function deleteHubTask" in content, "FAIL: deleteHubTask missing!"
    assert "function _getTeamForUser_" in content, "FAIL: _getTeamForUser_ missing!"
    assert "team: _getTeamForUser_(email)" in content, "FAIL: team return missing!"
    print("All checks passed!")
    print()
    print("Next steps:")
    print(f'  git add "{TARGET_FILE}"')
    print('  git commit -m "fix: restaurar Código.js original + agregar Hub Tasks al final"')
    print("  git push")

if __name__ == "__main__":
    main()
