#!/usr/bin/env bash
set -euo pipefail

usage() {
  cat <<'EOF'
Usage:
  scripts/release_quick.sh <version> [options]

Example:
  scripts/release_quick.sh 0.1.2

Options:
  --branch <name>              Branch to release from (default: main)
  --remote <name>              Git remote to push to (default: origin)
  --owner-repo <owner/repo>    GitHub owner/repo for release assets (default: N10ELabs/docwarp)
  --tap-repo-url <url>         Homebrew tap git URL (default: https://github.com/N10ELabs/homebrew-tap)
  --skip-tests                 Skip cargo test + version preflight checks
  --skip-rc                    Skip prerelease tag push (vX.Y.Z-rc.1)
  --skip-tap                   Skip Homebrew tap update
  --max-wait-attempts <n>      Release asset polling attempts (default: 30)
  --wait-seconds <n>           Sleep between polling attempts (default: 10)
  --help                       Show this help text

Environment:
  GITHUB_TOKEN or GH_TOKEN     Optional token for private release asset fetches
EOF
}

log() {
  printf '[release-quick] %s\n' "$*"
}

fail() {
  printf '[release-quick] error: %s\n' "$*" >&2
  exit 1
}

require_cmd() {
  local cmd="$1"
  command -v "${cmd}" >/dev/null 2>&1 || fail "required command not found: ${cmd}"
}

workspace_version_from_cargo_toml() {
  awk '
    /^\[workspace\.package\]/ { in_pkg = 1; next }
    /^\[/ { in_pkg = 0 }
    in_pkg && /^[[:space:]]*version[[:space:]]*=/ {
      if (match($0, /"[^"]+"/)) {
        value = substr($0, RSTART + 1, RLENGTH - 2)
        print value
        exit
      }
    }
  ' Cargo.toml
}

tag_exists_local() {
  local tag="$1"
  git rev-parse -q --verify "refs/tags/${tag}" >/dev/null 2>&1
}

tag_exists_remote() {
  local remote="$1"
  local tag="$2"
  git ls-remote --exit-code --tags "${remote}" "refs/tags/${tag}" >/dev/null 2>&1
}

create_and_push_tag() {
  local remote="$1"
  local tag="$2"

  if tag_exists_local "${tag}"; then
    fail "tag already exists locally: ${tag}"
  fi
  if tag_exists_remote "${remote}" "${tag}"; then
    fail "tag already exists on ${remote}: ${tag}"
  fi

  git tag "${tag}"
  git push "${remote}" "${tag}"
  log "pushed tag ${tag}"
}

wait_for_release_asset() {
  local url="$1"
  local attempts="$2"
  local wait_seconds="$3"
  local token="${GITHUB_TOKEN:-${GH_TOKEN:-}}"

  for ((attempt = 1; attempt <= attempts; attempt += 1)); do
    if [[ -n "${token}" ]]; then
      if curl -fsSIL -H "Authorization: Bearer ${token}" "${url}" >/dev/null; then
        return 0
      fi
    elif curl -fsSIL "${url}" >/dev/null; then
      return 0
    fi
    log "waiting for release asset (${attempt}/${attempts}): ${url}"
    sleep "${wait_seconds}"
  done

  return 1
}

download_release_asset() {
  local url="$1"
  local output="$2"
  local token="${GITHUB_TOKEN:-${GH_TOKEN:-}}"

  if [[ -n "${token}" ]]; then
    curl -fsSL -H "Authorization: Bearer ${token}" "${url}" -o "${output}"
  else
    curl -fsSL "${url}" -o "${output}"
  fi
}

main() {
  local branch="main"
  local remote="origin"
  local owner_repo="N10ELabs/docwarp"
  local tap_repo_url="https://github.com/N10ELabs/homebrew-tap"
  local run_tests=1
  local push_rc=1
  local update_tap=1
  local max_wait_attempts=30
  local wait_seconds=10
  local version=""

  while [[ $# -gt 0 ]]; do
    case "$1" in
      --help|-h)
        usage
        exit 0
        ;;
      --branch)
        [[ $# -ge 2 ]] || fail "missing value for --branch"
        branch="$2"
        shift 2
        ;;
      --remote)
        [[ $# -ge 2 ]] || fail "missing value for --remote"
        remote="$2"
        shift 2
        ;;
      --owner-repo)
        [[ $# -ge 2 ]] || fail "missing value for --owner-repo"
        owner_repo="$2"
        shift 2
        ;;
      --tap-repo-url)
        [[ $# -ge 2 ]] || fail "missing value for --tap-repo-url"
        tap_repo_url="$2"
        shift 2
        ;;
      --skip-tests)
        run_tests=0
        shift
        ;;
      --skip-rc)
        push_rc=0
        shift
        ;;
      --skip-tap)
        update_tap=0
        shift
        ;;
      --max-wait-attempts)
        [[ $# -ge 2 ]] || fail "missing value for --max-wait-attempts"
        max_wait_attempts="$2"
        shift 2
        ;;
      --wait-seconds)
        [[ $# -ge 2 ]] || fail "missing value for --wait-seconds"
        wait_seconds="$2"
        shift 2
        ;;
      -*)
        fail "unknown option: $1"
        ;;
      *)
        if [[ -n "${version}" ]]; then
          fail "multiple version arguments provided: '${version}' and '$1'"
        fi
        version="$1"
        shift
        ;;
    esac
  done

  [[ -n "${version}" ]] || fail "missing <version> argument (example: 0.1.2)"
  [[ "${version}" =~ ^[0-9]+\.[0-9]+\.[0-9]+([-.][0-9A-Za-z.]+)?$ ]] || fail "invalid version '${version}'"
  [[ "${max_wait_attempts}" =~ ^[0-9]+$ ]] || fail "--max-wait-attempts must be an integer"
  [[ "${wait_seconds}" =~ ^[0-9]+$ ]] || fail "--wait-seconds must be an integer"

  require_cmd git
  require_cmd cargo
  require_cmd curl
  require_cmd mktemp
  require_cmd awk

  local script_dir repo_root
  script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
  repo_root="$(cd "${script_dir}/.." && pwd)"
  cd "${repo_root}"

  git rev-parse --show-toplevel >/dev/null 2>&1 || fail "not inside a git repository"

  local current_branch
  current_branch="$(git rev-parse --abbrev-ref HEAD)"
  [[ "${current_branch}" == "${branch}" ]] || fail "current branch is '${current_branch}', expected '${branch}'"

  local workspace_version
  workspace_version="$(workspace_version_from_cargo_toml)"
  [[ -n "${workspace_version}" ]] || fail "failed to read [workspace.package].version from Cargo.toml"
  [[ "${workspace_version}" == "${version}" ]] || fail "Cargo.toml version is '${workspace_version}', expected '${version}'"

  local tag rc_tag
  tag="v${version}"
  rc_tag="${tag}-rc.1"

  log "fetching latest refs and tags from ${remote}"
  git fetch "${remote}" --tags

  log "pulling ${remote}/${branch} with --ff-only"
  git pull --ff-only "${remote}" "${branch}"

  if (( run_tests == 1 )); then
    log "running cargo test --workspace --all-targets"
    cargo test --workspace --all-targets

    log "verifying CLI version output"
    local cli_version_output expected_version_line
    cli_version_output="$(cargo run -q -p docwarp-cli -- --version)"
    expected_version_line="docwarp ${version}"
    [[ "${cli_version_output}" == *"${expected_version_line}"* ]] || {
      fail "unexpected --version output: '${cli_version_output}' (expected to include '${expected_version_line}')"
    }
  fi

  log "staging and committing release changes (if any)"
  git add -A
  if git diff --cached --quiet; then
    log "no staged changes to commit"
  else
    git commit -m "release: ${tag}"
  fi

  log "pushing ${branch} to ${remote}"
  git push "${remote}" "${branch}"

  if (( push_rc == 1 )); then
    create_and_push_tag "${remote}" "${rc_tag}"
  else
    log "skipping prerelease tag push (--skip-rc)"
  fi

  create_and_push_tag "${remote}" "${tag}"

  if (( update_tap == 1 )); then
    local formula_asset_url
    formula_asset_url="https://github.com/${owner_repo}/releases/download/${tag}/docwarp.rb"
    log "waiting for release formula asset: ${formula_asset_url}"
    wait_for_release_asset "${formula_asset_url}" "${max_wait_attempts}" "${wait_seconds}" \
      || fail "release asset not available after waiting: ${formula_asset_url}"

    local temp_dir tap_dir formula_path
    temp_dir="$(mktemp -d)"
    trap 'rm -rf "${temp_dir}"' EXIT
    tap_dir="${temp_dir}/homebrew-tap"
    formula_path="${tap_dir}/Formula/docwarp.rb"

    log "cloning tap repo: ${tap_repo_url}"
    git clone "${tap_repo_url}" "${tap_dir}"

    log "downloading release formula into tap repo"
    download_release_asset "${formula_asset_url}" "${formula_path}"

    if git -C "${tap_dir}" diff --quiet -- "Formula/docwarp.rb"; then
      log "tap formula unchanged; skipping tap commit"
    else
      git -C "${tap_dir}" add "Formula/docwarp.rb"
      git -C "${tap_dir}" commit -m "docwarp ${version}"
      git -C "${tap_dir}" push
      log "pushed tap update for docwarp ${version}"
    fi
  else
    log "skipping tap update (--skip-tap)"
  fi

  log "release quick action complete"
  log "next manual step (optional): open/update Homebrew/core PR for brew install docwarp"
}

main "$@"
