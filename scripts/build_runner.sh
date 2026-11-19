#!/usr/bin/env bash
# Build the encrypted (compiled) runner binary with Nuitka and publish it to
# the server so installed runners self-update.
#
# Usage:
#   scripts/build_runner.sh <version> [platform]
#
#   <version>  semver baked into the binary, e.g. 1.0.2 (also names the file)
#   [platform] defaults to "macos-arm64"; the binary is only usable on the
#              platform it was built on, so rebuild per OS.
#
# Requirements: a venv with requirements.txt + nuitka installed. The script
# uses $BUILD_VENV if set, else falls back to a conventional path.

set -euo pipefail

VERSION="${1:?usage: build_runner.sh <version> [platform]}"
PLATFORM="${2:-macos-arm64}"
HERE="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "$HERE/.." && pwd)"
WORK="$ROOT/.runner-build"
PUBLISH="$ROOT/web/static/runner"
NAME="customsappify-runner-$VERSION-$PLATFORM"
CLIENT="$ROOT/web/runner/client.py"

BUILD_VENV="${BUILD_VENV:-$WORK/venv}"

echo "==> staging source at $WORK"
rm -rf "$WORK/src"
mkdir -p "$WORK/src/web"
cp -R "$ROOT/web/runner" "$ROOT/web/modules" "$WORK/src/web/"
cp "$ROOT/web/__init__.py" "$WORK/src/web/"
cp "$ROOT/requirements.txt" "$WORK/src/"

echo "==> stamping RUNNER_VERSION=$VERSION"
python3 - "$CLIENT" "$WORK/src/web/runner/client.py" "$VERSION" <<'PY'
import re, sys
src, dst, version = sys.argv[1], sys.argv[2], sys.argv[3]
with open(src) as fh:
    text = fh.read()
text = re.sub(r'^RUNNER_VERSION = ".*?"', f'RUNNER_VERSION = "{version}"', text, count=1, flags=re.M)
with open(dst, "w") as fh:
    fh.write(text)
PY

if [ ! -x "$BUILD_VENV/bin/python" ]; then
  echo "==> creating build venv at $BUILD_VENV"
  python3 -m venv "$BUILD_VENV"
  "$BUILD_VENV/bin/pip" install -q -U pip
  "$BUILD_VENV/bin/pip" install -q -r "$WORK/src/requirements.txt"
fi

echo "==> compiling with Nuitka (this takes several minutes)"
cd "$WORK/src"
export CCACHE_DIR="${CCACHE_DIR:-$WORK/ccache}"
mkdir -p "$CCACHE_DIR"
export PATH="/opt/homebrew/bin:$PATH"
"$BUILD_VENV/bin/python" -m nuitka \
  --onefile --standalone --assume-yes-for-downloads --jobs=4 \
  --include-package=websockets --include-package=websockets.asyncio \
  --include-package=pandas --include-package=dateutil \
  --include-package=selenium --include-package=webdriver_manager \
  --include-package=PIL --include-package=pikepdf \
  --include-package=undetected_chromedriver --include-package=numpy \
  --nofollow-import-to=pymupdf --nofollow-import-to=pymupdf.mupdf \
  --nofollow-import-to=web.modules --nofollow-import-to=web.modules.* \
  --no-deployment-flag=excluded-module-usage \
  --output-dir="$WORK/dist" --output-filename="customsappify-runner" \
  web/runner/client.py

mkdir -p "$PUBLISH"
mv "$WORK/dist/customsappify-runner" "$PUBLISH/$NAME"
chmod +x "$PUBLISH/$NAME"
echo "==> published $PUBLISH/$NAME"
ls -lh "$PUBLISH/$NAME"
