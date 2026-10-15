#!/usr/bin/env bash
# Normalize unstaged VB6 IDE casing/line-ending noise against the index.
# Staged content is never changed, including partially staged files.
set -euo pipefail
export LC_ALL=C
export LANG=C

cd "$(git rev-parse --show-toplevel)"
scratch=$(mktemp -d)
trap 'rm -rf "$scratch"' EXIT

make_patch() {
    local status=0
    diff -u -i --ignore-blank-lines "$1" "$2" > "$scratch/change.patch" || status=$?
    # diff returns 1 for differences; an actual error must abort the hook.
    if (( status > 1 )); then
        return "$status"
    fi
}

# NUL-delimited paths preserve spaces and non-ASCII filenames. Only tracked,
# unstaged modifications are eligible; added/deleted/renamed files stay intact.
git diff --name-only -z --diff-filter=M > "$scratch/files"
while IFS= read -r -d '' file; do
    case "$file" in
        *.[bB][aA][sS]|*.[cC][lL][sS]|*.[fF][rR][mM]|*.[dD][sS][rR]|*.[vV][bB][pP]) ;;
        *) continue ;;
    esac
    if [[ -L "$file" ]]; then
        echo "Refusing to normalize symlink: $file" >&2
        exit 1
    fi

    # Git stores UTF-8 but VB6 uses Windows-1252. Apply checkout filters before
    # comparing bytes; copying the raw index would corrupt accented source text.
    git cat-file --filters ":$file" > "$scratch/index"
    sed --binary 's/\r$//' "$scratch/index" > "$scratch/original"
    sed --binary 's/\r$//' "$file" > "$scratch/working"

    case "$file" in
        *.[vV][bB][pP])
            # Preserve the existing rule for machine-specific COM reference tails.
            awk -F'#' '{print $1"#"$2"#"$3}' "$scratch/original" > "$scratch/left"
            awk -F'#' '{print $1"#"$2"#"$3}' "$scratch/working" > "$scratch/right"
            status=0
            diff -q -i --ignore-blank-lines "$scratch/left" "$scratch/right" >/dev/null || status=$?
            if (( status > 1 )); then exit "$status"; fi
            if (( status == 0 )); then
                : > "$scratch/change.patch"
            else
                make_patch "$scratch/original" "$scratch/working"
            fi
            ;;
        *) make_patch "$scratch/original" "$scratch/working" ;;
    esac

    if [[ -s "$scratch/change.patch" ]]; then
        patch --batch --silent "$scratch/original" < "$scratch/change.patch"
    fi
    sed --binary 's/$/\r/' "$scratch/original" > "$scratch/result"
    if ! cmp -s "$scratch/result" "$file"; then
        cp "$scratch/result" "$file"
    fi
done < "$scratch/files"
