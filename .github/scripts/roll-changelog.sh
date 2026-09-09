#!/usr/bin/env bash
#
# Rolls the "Unreleased" section of CHANGELOG.md into a dated version heading and
# leaves a fresh, empty "Unreleased" section at the top.
#
#   ## Unreleased          ->   ## Unreleased
#
#   ### Fixed                    ## v0.106.0 - 2026-07-25
#   - ...
#                                ### Fixed
#                                - ...
#
# Usage: roll-changelog.sh <version> [date] [changelog-path]
#
#   version         Version being released, with or without a leading "v".
#   date            Release date as YYYY-MM-DD. Defaults to today (UTC).
#   changelog-path  Defaults to CHANGELOG.md in the repository root.
#
# Exits non-zero if the changelog has no "Unreleased" heading, or if that section
# is empty — releasing with nothing written down is almost always a mistake.

set -euo pipefail

version=${1:?usage: roll-changelog.sh <version> [date] [changelog-path]}
date=${2:-$(date -u +%Y-%m-%d)}
changelog=${3:-"$(dirname "$0")/../../CHANGELOG.md"}

# Normalise to a leading "v" so headings read "## v0.106.0".
version=${version#v}
heading="## v${version} - ${date}"

if [[ ! -f $changelog ]]; then
  echo "::error::Changelog not found: $changelog" >&2
  exit 2
fi

if ! grep -qE '^## Unreleased[[:space:]]*$' "$changelog"; then
  echo "::error::No '## Unreleased' heading found in $changelog" >&2
  exit 3
fi

if grep -qE "^## v${version//./\\.} " "$changelog"; then
  echo "::error::$changelog already contains a heading for v${version}" >&2
  exit 4
fi

# Everything between "## Unreleased" and the next "## " heading (or EOF).
unreleased_body=$(awk '
  /^## Unreleased[[:space:]]*$/ { inside = 1; next }
  inside && /^## / { exit }
  inside { print }
' "$changelog")

if [[ -z ${unreleased_body//[[:space:]]/} ]]; then
  echo "::error::The Unreleased section of $changelog is empty — nothing to release" >&2
  exit 5
fi

tmp=$(mktemp)
trap 'rm -f "$tmp"' EXIT

awk -v heading="$heading" '
  { print }
  !done && /^## Unreleased[[:space:]]*$/ {
    print ""
    print heading
    done = 1
  }
' "$changelog" >"$tmp"

mv "$tmp" "$changelog"
trap - EXIT

# Rebuild the "Contents" list from the "## " headings, so it cannot drift out of
# date as releases are rolled. Anchors follow GitHub's rule: lower-case, drop every
# character that is not a letter, digit, space or hyphen, then spaces to hyphens.
contents=$(
  grep -E '^## ' "$changelog" | while IFS= read -r line; do
    text=${line#'## '}
    text=${text%$'\r'}
    [[ $text == Contents ]] && continue
    anchor=$(printf '%s' "$text" | tr '[:upper:]' '[:lower:]' | tr -cd 'a-z0-9 -' | tr ' ' '-')
    # Link text is the heading without its " - YYYY-MM-DD" suffix.
    label=${text% - [0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]}
    printf -- '- [%s](#%s)\n' "$label" "$anchor"
  done
)

if [[ -n $contents ]]; then
  tmp=$(mktemp)
  trap 'rm -f "$tmp"' EXIT

  awk -v contents="$contents" '
    !seen && /^## Contents[[:space:]]*$/ {
      print; print ""; print contents; print ""
      seen = 1; skip = 1
      next
    }
    skip && /^## / { skip = 0 }
    skip { next }
    { print }
  ' "$changelog" >"$tmp"

  mv "$tmp" "$changelog"
  trap - EXIT
fi

echo "Rolled Unreleased into '${heading}' in ${changelog}"
