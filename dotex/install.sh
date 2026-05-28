#!/usr/bin/env bash
set -euo pipefail

APP_NAME="dotex"
REPO_SLUG_DEFAULT="MittyLeeisOK/dotex"

ACTION=install
INSTALL_SCOPE=user
YES=false
INSTALL_DEPS=true
PYTHON_BIN="${PYTHON_BIN:-python3}"
REPO_TARBALL_URL="${REPO_TARBALL_URL:-https://codeload.github.com/${REPO_SLUG_DEFAULT}/tar.gz/refs/heads/main}"
STEP_COUNTER=0
TOTAL_STEPS=6

INSTALL_DIR=""
BIN_DIR=""
SRC_ROOT=""
VENV_DIR=""
LAUNCHER=""
LEGACY_LAUNCHER_TOOLKIT=""
LEGACY_LAUNCHER_TOOL=""
STAGE_DIR=""

if [ -t 1 ] && [ -z "${NO_COLOR:-}" ] && [ "${TERM:-}" != "dumb" ]; then
  GREEN='\033[32m'
  YELLOW='\033[33m'
  RED='\033[31m'
  BLUE='\033[34m'
  RESET='\033[0m'
else
  GREEN=''
  YELLOW=''
  RED=''
  BLUE=''
  RESET=''
fi

ok() { echo -e "${GREEN}[OK] $*${RESET}"; }
warn() { echo -e "${YELLOW}[WARN] $*${RESET}"; }
err() { echo -e "${RED}[ERROR] $*${RESET}"; }
info() { echo -e "${BLUE}[INFO] $*${RESET}"; }
progress() {
  STEP_COUNTER=$((STEP_COUNTER + 1))
  info "[${STEP_COUNTER}/${TOTAL_STEPS}] $*"
}

cleanup_stage() {
  if [ -n "${STAGE_DIR}" ] && [ -d "${STAGE_DIR}" ]; then
    rm -rf "${STAGE_DIR}"
  fi
}

trap cleanup_stage EXIT

usage() {
  cat <<EOF
Usage: ./install.sh [options]

Actions:
  --install                Install the toolkit (default)
  --upgrade                Reinstall from the latest source into the same target
  --uninstall              Remove the installed toolkit and launchers
  --purge                  Same as uninstall for this project

Install scope:
  --user                   Install into ~/.local/share and ~/.local/bin (default)
  --system                 Install into /opt and /usr/local/bin (root required)

Other options:
  --yes                    Skip confirmation prompts
  --install-deps           Kept for compatibility; dependency auto-install is already on by default
  --skip-deps              Skip automatic dependency installation
  --python PATH            Use a specific Python interpreter
  --repo-tarball-url URL   Download source from a specific GitHub tarball URL
  --help                   Show this help message

Examples:
  ./install.sh --install
  ./install.sh --install --skip-deps
  ./install.sh --upgrade
  ./install.sh --system --install
  ./install.sh --uninstall
EOF
}

while [ $# -gt 0 ]; do
  case "$1" in
    --install)
      ACTION=install
      ;;
    --upgrade)
      ACTION=upgrade
      ;;
    --uninstall)
      ACTION=uninstall
      ;;
    --purge)
      ACTION=purge
      ;;
    --user)
      INSTALL_SCOPE=user
      ;;
    --system)
      INSTALL_SCOPE=system
      ;;
    --yes)
      YES=true
      ;;
    --install-deps)
      INSTALL_DEPS=true
      ;;
    --skip-deps)
      INSTALL_DEPS=false
      ;;
    --python)
      shift
      [ $# -gt 0 ] || { err "missing value for --python"; exit 1; }
      PYTHON_BIN="$1"
      ;;
    --repo-tarball-url)
      shift
      [ $# -gt 0 ] || { err "missing value for --repo-tarball-url"; exit 1; }
      REPO_TARBALL_URL="$1"
      ;;
    --help)
      usage
      exit 0
      ;;
    *)
      err "unknown option: $1"
      usage
      exit 1
      ;;
  esac
  shift
done

resolve_layout() {
  if [ "${INSTALL_SCOPE}" = system ]; then
    INSTALL_DIR="/opt/${APP_NAME}"
    BIN_DIR="/usr/local/bin"
  else
    INSTALL_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/${APP_NAME}"
    BIN_DIR="${HOME}/.local/bin"
  fi
  SRC_ROOT="${INSTALL_DIR}/src"
  VENV_DIR="${INSTALL_DIR}/venv"
  LAUNCHER="${BIN_DIR}/dotex"
  LEGACY_LAUNCHER_TOOLKIT="${BIN_DIR}/tex-docx-toolkit"
  LEGACY_LAUNCHER_TOOL="${BIN_DIR}/tex-docx-tool"
}

require_root() {
  if [ "$(id -u)" -ne 0 ]; then
    err "this operation requires root privileges"
    exit 1
  fi
}

run_with_elevation() {
  if [ "$(id -u)" -eq 0 ]; then
    "$@"
    return 0
  fi
  if command -v sudo >/dev/null 2>&1; then
    sudo "$@"
    return 0
  fi
  return 1
}

confirm() {
  local prompt="$1"
  if [ "${YES}" = true ]; then
    return 0
  fi
  printf "%s [y/N] " "${prompt}"
  read -r reply
  case "${reply}" in
    y|Y|yes|YES)
      return 0
      ;;
    *)
      return 1
      ;;
  esac
}

install_deps_if_requested() {
  if [ "${INSTALL_DEPS}" != true ]; then
    info "automatic dependency installation skipped"
    return 0
  fi

  if command -v pandoc >/dev/null 2>&1; then
    ok "pandoc already available"
    return 0
  fi

  info "pandoc not found; attempting automatic installation"

  if [ "$(uname -s)" = Darwin ]; then
    command -v brew >/dev/null 2>&1 || {
      warn "Homebrew was not found; install pandoc manually or rerun after installing brew"
      return 0
    }
    brew install pandoc
    return 0
  fi

  if command -v apt-get >/dev/null 2>&1; then
    run_with_elevation apt-get update || {
      warn "failed to elevate for apt-get update; install pandoc manually"
      return 0
    }
    run_with_elevation apt-get install -y python3 python3-venv python3-pip curl pandoc || {
      warn "apt-get install failed; install pandoc manually"
      return 0
    }
    return 0
  fi

  if command -v dnf >/dev/null 2>&1; then
    run_with_elevation dnf install -y python3 python3-pip python3-virtualenv curl pandoc || {
      warn "dnf install failed; install pandoc manually"
      return 0
    }
    return 0
  fi

  if command -v yum >/dev/null 2>&1; then
    run_with_elevation yum install -y python3 python3-pip curl pandoc || {
      warn "yum install failed; install pandoc manually"
      return 0
    }
    return 0
  fi

  if command -v pacman >/dev/null 2>&1; then
    run_with_elevation pacman -Sy --noconfirm python python-pip curl pandoc || {
      warn "pacman install failed; install pandoc manually"
      return 0
    }
    return 0
  fi

  warn "automatic dependency installation is not supported on this platform; install pandoc manually"
}

check_python() {
  "${PYTHON_BIN}" - <<'PY' >/dev/null
import sys
import venv
assert sys.version_info >= (3, 9)
PY
}

ensure_prereqs() {
  command -v curl >/dev/null 2>&1 || { err "curl is required"; exit 1; }
  if ! check_python 2>/dev/null; then
    err "Python 3.9+ with venv support is required: ${PYTHON_BIN}"
    exit 1
  fi
}

stage_source() {
  STAGE_DIR="$(mktemp -d)"

  if [ -f "./pyproject.toml" ] && [ -d "./src/dotex" ]; then
    cp -a . "${STAGE_DIR}/source"
    echo "${STAGE_DIR}/source"
    return 0
  fi

  local tarball="${STAGE_DIR}/repo.tar.gz"
  curl -fsSL "${REPO_TARBALL_URL}" -o "${tarball}" || {
    err "failed to download source tarball: ${REPO_TARBALL_URL}"
    exit 1
  }

  tar -xzf "${tarball}" -C "${STAGE_DIR}"
  local extracted
  extracted="$(find "${STAGE_DIR}" -maxdepth 1 -mindepth 1 -type d | head -n 1)"
  if [ -z "${extracted}" ] || [ ! -f "${extracted}/pyproject.toml" ] || [ ! -d "${extracted}/src/dotex" ]; then
    err "downloaded source tree is missing pyproject.toml or src/dotex"
    exit 1
  fi

  echo "${extracted}"
}

write_launcher() {
  local launcher_path="$1"
  local entrypoint="$2"

  cat > "${launcher_path}" <<EOF
#!/usr/bin/env bash
set -euo pipefail
exec "${VENV_DIR}/bin/${entrypoint}" "\$@"
EOF
  chmod 755 "${launcher_path}"
}

perform_install() {
  local source_dir="$1"
  local first_install=false

  if [ ! -x "${LAUNCHER}" ] || [ ! -d "${VENV_DIR}" ]; then
    first_install=true
  fi

  mkdir -p "${INSTALL_DIR}" "${BIN_DIR}"
  rm -rf "${SRC_ROOT}" "${VENV_DIR}"
  mkdir -p "${SRC_ROOT}"
  cp -a "${source_dir}"/. "${SRC_ROOT}/"

  progress "creating virtual environment"
  "${PYTHON_BIN}" -m venv "${VENV_DIR}"

  progress "installing Python package and runtime dependencies"
  "${VENV_DIR}/bin/python" -m pip install --upgrade pip setuptools wheel
  "${VENV_DIR}/bin/python" -m pip install "${SRC_ROOT}"

  progress "writing command launchers"
  write_launcher "${LAUNCHER}" "dotex"
  rm -f "${LEGACY_LAUNCHER_TOOLKIT}" "${LEGACY_LAUNCHER_TOOL}"

  ok "${APP_NAME} installed"
  ok "launcher: ${LAUNCHER}"

  if [ "${first_install}" = true ]; then
    print_first_install_notes
  else
    print_upgrade_notes
  fi
}

remove_installation() {
  rm -f "${LAUNCHER}" "${LEGACY_LAUNCHER_TOOLKIT}" "${LEGACY_LAUNCHER_TOOL}"
  rm -rf "${INSTALL_DIR}"
  ok "${APP_NAME} removed"
}

print_first_install_notes() {
  cat <<EOF

首次安装说明

常用命令：
  dotex --help
  dotex inspect-template /path/to/reference-template.docx --output artifacts/template.json
  dotex convert-docx /path/to/manuscript.tex -t /path/to/reference-template.docx -o /path/to/manuscript.docx
  dotex convert-tex /path/to/manuscript.docx --output /path/to/manuscript.tex
  dotex compare-roundtrip /path/to/original.docx /path/to/source.tex /path/to/generated.docx

注意事项：
  1. 运行时需要 Python 3.9+。
  2. 运行 convert-docx 和 convert-tex 需要系统中可用的 pandoc；安装脚本会在首次安装时自动尝试安装它，失败时再手动补装。
  3. Zotero 模式默认读取 ~/Zotero/zotero.sqlite；如果本地库不完整，工具仍会输出 Zotero field，并另外写出 xlsx 检查表供补齐条目。
  4. 工具会在 artifacts/ 和生成的 DOCX 附近写出中间产物与检查文件，里面可能包含你的稿件内容，不要把运行结果直接提交到公开仓库。
EOF

  if [ "${INSTALL_SCOPE}" = user ] && [[ ":${PATH}:" != *":${BIN_DIR}:"* ]]; then
    warn "${BIN_DIR} is not in PATH; add it to your shell profile before using dotex directly"
  fi

  if ! command -v pandoc >/dev/null 2>&1; then
    warn "pandoc was not found in PATH; install it before running document conversions"
  fi
}

print_upgrade_notes() {
  info "upgrade completed"
  if ! command -v pandoc >/dev/null 2>&1; then
    warn "pandoc was not found in PATH; install it before running document conversions"
  fi
}

resolve_layout

if [ "${INSTALL_SCOPE}" = system ]; then
  require_root
fi

case "${ACTION}" in
  uninstall|purge)
    if [ -d "${INSTALL_DIR}" ] || [ -e "${LAUNCHER}" ] || [ -e "${LEGACY_LAUNCHER_TOOLKIT}" ] || [ -e "${LEGACY_LAUNCHER_TOOL}" ]; then
      if confirm "Remove ${APP_NAME} from ${INSTALL_DIR}?"; then
        remove_installation
      else
        warn "aborted"
        exit 1
      fi
    else
      warn "nothing to remove"
    fi
    exit 0
    ;;
esac

progress "checking or installing system dependencies"
install_deps_if_requested

progress "checking Python prerequisites"
ensure_prereqs

progress "staging source code"
SOURCE_DIR="$(stage_source)"
perform_install "${SOURCE_DIR}"