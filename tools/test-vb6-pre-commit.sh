#!/usr/bin/env bash
# Isolated Git/encoding regressions; no VB6 compiler or user checkout mutation.
set -euo pipefail
export LC_ALL=C
source_root=$(cd "$(dirname "$0")/.." && pwd)
scratch=$(mktemp -d)
trap 'rm -rf "$scratch"' EXIT
mkdir "$scratch/repo"
cd "$scratch/repo"
git init -q
git config user.name 'Hook Test'
git config user.email 'hook-test@example.invalid'
git config core.autocrlf true
git config core.hooksPath .githooks
mkdir .githooks
cp "$source_root/.gitattributes" .
cp "$source_root/.githooks/pre-commit" .githooks/
cp "$source_root/git_ignore_case.sh" .
chmod +x .githooks/pre-commit git_ignore_case.sh
printf 'Option Explicit\r\nPrivate CountValue As Long\r\nCountValue = 1\r\n\047 Caf\351\r\n' > 'Spaced Module.bas'
cp 'Spaced Module.bas' Deleted.bas
printf 'Type=Exe\r\nReference=*\\G{123}#1.0#0#old/path#Library\r\nName="Fixture"\r\n' > Project.vbp
git add .
git commit -qm 'Fixture baseline'
git checkout-index --force --all
passed=0
pass() { passed=$((passed + 1)); printf 'PASS %s\n' "$1"; }
hook() { bash .githooks/pre-commit; }

if grep -q $'\r' .githooks/pre-commit git_ignore_case.sh; then exit 1; fi
pass 'LF hook checkout with core.autocrlf=true'
hook
git diff --exit-code
pass 'clean worktree'

cp 'Spaced Module.bas' "$scratch/expected"
sed --binary 's/Option Explicit/option explicit/; s/CountValue/countvalue/g' "$scratch/expected" > 'Spaced Module.bas'
hook
cmp "$scratch/expected" 'Spaced Module.bas'
pass 'case-only changes and paths containing spaces'

sed --binary 's/= 1/= 2/' "$scratch/expected" > "$scratch/expected2"
sed --binary 's/Option Explicit/option explicit/' "$scratch/expected2" > 'Spaced Module.bas'
hook
cmp "$scratch/expected2" 'Spaced Module.bas'
pass 'meaningful edits, Windows-1252 accents and CRLF preserved'

git add 'Spaced Module.bas'
staged=$(git write-tree)
sed --binary 's/= 2/= 3/' "$scratch/expected2" > 'Spaced Module.bas'
cp 'Spaced Module.bas' "$scratch/expected3"
hook
test "$(git write-tree)" = "$staged"
cmp "$scratch/expected3" 'Spaced Module.bas'
pass 'partial staging is unchanged'

printf 'Option Explicit\r\n' > Added.bas
git add Added.bas
rm Deleted.bas
staged=$(git write-tree)
hook
test "$(git write-tree)" = "$staged"
test ! -e Deleted.bas
test -f Added.bas
pass 'added and deleted files untouched'

cp Project.vbp "$scratch/project"
sed --binary 's@old/path@new/path@' "$scratch/project" > Project.vbp
hook
cmp "$scratch/project" Project.vbp
pass 'project reference path noise'
sed --binary 's/Fixture/Changed/' "$scratch/project" > Project.vbp
cp Project.vbp "$scratch/project-changed"
hook
cmp "$scratch/project-changed" Project.vbp
pass 'meaningful project edits'

mkdir "$scratch/bin"
printf '#!/usr/bin/env bash\nexit 2\n' > "$scratch/bin/diff"
chmod +x "$scratch/bin/diff"
if PATH="$scratch/bin:$PATH" bash .githooks/pre-commit; then
    echo 'Hook swallowed a diff failure' >&2
    exit 1
fi
cmp "$scratch/expected3" 'Spaced Module.bas'
pass 'tool failures abort without overwriting source'

git add 'Spaced Module.bas' Project.vbp Deleted.bas
git commit -qm 'Commit through real hook'
git diff --exit-code
git diff --cached --exit-code
pass 'real commit executes the hook'
printf '%s/%s hook tests passed\n' "$passed" "$passed"
