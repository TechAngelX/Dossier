#!/bin/bash

# ============================================================
# Playwrighter - Build to Desktop Script
# ============================================================
# Builds a self-contained, single-file executable and copies
# it to the user's Desktop folder (as .app bundle on macOS).
# ============================================================

set -e  # Exit on error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

echo -e "${CYAN}"
echo "╔══════════════════════════════════════════════════════════════╗"
echo "║                                                              ║"
echo "║     ██████╗ ██╗      █████╗ ██╗   ██╗██╗    ██╗██████╗      ║"
echo "║     ██╔══██╗██║     ██╔══██╗╚██╗ ██╔╝██║    ██║██╔══██╗     ║"
echo "║     ██████╔╝██║     ███████║ ╚████╔╝ ██║ █╗ ██║██████╔╝     ║"
echo "║     ██╔═══╝ ██║     ██╔══██║  ╚██╔╝  ██║███╗██║██╔══██╗     ║"
echo "║     ██║     ███████╗██║  ██║   ██║   ╚███╔███╔╝██║  ██║     ║"
echo "║     ╚═╝     ╚══════╝╚═╝  ╚═╝   ╚═╝    ╚══╝╚══╝ ╚═╝  ╚═╝     ║"
echo "║                                                              ║"
echo "║                 Build to Desktop Script                      ║"
echo "╚══════════════════════════════════════════════════════════════╝"
echo -e "${NC}"

# Detect the script's directory (project root)
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Detect OS and architecture
OS="$(uname -s)"
ARCH="$(uname -m)"

case "$OS" in
    Darwin)
        if [ "$ARCH" = "arm64" ]; then
            RUNTIME="osx-arm64"
            echo -e "  ${GREEN}●${NC} Platform: ${CYAN}macOS Apple Silicon (M1/M2/M3/M4)${NC}"
        else
            RUNTIME="osx-x64"
            echo -e "  ${GREEN}●${NC} Platform: ${CYAN}macOS Intel${NC}"
        fi
        DESKTOP="$HOME/Desktop"
        OUTPUT_NAME="Playwrighter"
        CREATE_APP_BUNDLE=true
        ;;
    Linux)
        RUNTIME="linux-x64"
        echo -e "  ${GREEN}●${NC} Platform: ${CYAN}Linux x64${NC}"
        DESKTOP="$HOME/Desktop"
        OUTPUT_NAME="Playwrighter"
        CREATE_APP_BUNDLE=false
        ;;
    MINGW*|MSYS*|CYGWIN*)
        RUNTIME="win-x64"
        echo -e "  ${GREEN}●${NC} Platform: ${CYAN}Windows x64${NC}"
        DESKTOP="$USERPROFILE/Desktop"
        OUTPUT_NAME="Playwrighter.exe"
        CREATE_APP_BUNDLE=false
        ;;
    *)
        echo -e "${RED}✗ Error: Unsupported operating system: $OS${NC}"
        exit 1
        ;;
esac

echo -e "  ${GREEN}●${NC} Runtime:  ${CYAN}${RUNTIME}${NC}"
echo ""

# Clean previous builds
echo -e "${BLUE}[1/4]${NC} Cleaning previous builds..."
dotnet clean -c Release --nologo -v q 2>/dev/null || true

# Restore dependencies
echo -e "${BLUE}[2/4]${NC} Restoring dependencies..."
dotnet restore --nologo -v q

# Build and publish
echo -e "${BLUE}[3/4]${NC} Building self-contained executable..."
dotnet publish -c Release -r "$RUNTIME" \
    --self-contained true \
    -p:PublishSingleFile=true \
    -p:IncludeNativeLibrariesForSelfExtract=true \
    -p:EnableCompressionInSingleFile=true \
    --nologo \
    -v q

# Determine publish output path
PUBLISH_DIR="$SCRIPT_DIR/bin/Release/net10.0/$RUNTIME/publish"

if [ ! -f "$PUBLISH_DIR/$OUTPUT_NAME" ]; then
    PUBLISH_DIR="$SCRIPT_DIR/bin/Release/$RUNTIME/publish"
fi

if [ ! -f "$PUBLISH_DIR/$OUTPUT_NAME" ]; then
    echo -e "${RED}✗ Error: Could not find built executable${NC}"
    echo "  Expected at: $PUBLISH_DIR/$OUTPUT_NAME"
    exit 1
fi

# Copy to Desktop
echo -e "${BLUE}[4/4]${NC} Packaging for Desktop..."

if [ "$CREATE_APP_BUNDLE" = true ]; then
    # Create macOS .app bundle
    APP_NAME="Playwrighter.app"
    APP_PATH="$DESKTOP/$APP_NAME"

    # Remove existing app bundle
    rm -rf "$APP_PATH"

    # Create app bundle structure
    mkdir -p "$APP_PATH/Contents/MacOS"
    mkdir -p "$APP_PATH/Contents/Resources"

    # Copy executable
    cp "$PUBLISH_DIR/$OUTPUT_NAME" "$APP_PATH/Contents/MacOS/Playwrighter"
    chmod +x "$APP_PATH/Contents/MacOS/Playwrighter"

    # Copy icon if exists
    if [ -f "$SCRIPT_DIR/Assets/Playwrighter.icns" ]; then
        cp "$SCRIPT_DIR/Assets/Playwrighter.icns" "$APP_PATH/Contents/Resources/AppIcon.icns"
    fi

    # Create Info.plist
    cat > "$APP_PATH/Contents/Info.plist" << 'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>Playwrighter</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundleIdentifier</key>
    <string>com.techangelx.playwrighter</string>
    <key>CFBundleName</key>
    <string>Playwrighter</string>
    <key>CFBundleDisplayName</key>
    <string>Playwrighter</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0.0</string>
    <key>CFBundleVersion</key>
    <string>1</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.15</string>
    <key>NSHighResolutionCapable</key>
    <true/>
    <key>NSHumanReadableCopyright</key>
    <string>Copyright © 2024 Tech Angel X. All rights reserved.</string>
</dict>
</plist>
PLIST

    # Get file size
    FILE_SIZE=$(du -sh "$APP_PATH" | cut -f1)
    FINAL_PATH="$APP_PATH"
else
    # Copy as regular executable
    cp "$PUBLISH_DIR/$OUTPUT_NAME" "$DESKTOP/$OUTPUT_NAME"

    if [ "$OS" != "MINGW"* ] && [ "$OS" != "MSYS"* ] && [ "$OS" != "CYGWIN"* ]; then
        chmod +x "$DESKTOP/$OUTPUT_NAME"
    fi

    FILE_SIZE=$(du -h "$DESKTOP/$OUTPUT_NAME" | cut -f1)
    FINAL_PATH="$DESKTOP/$OUTPUT_NAME"
fi

echo ""
echo -e "${GREEN}╔══════════════════════════════════════════════════════════════╗${NC}"
echo -e "${GREEN}║                                                              ║${NC}"
echo -e "${GREEN}║                    ✓ Build Complete!                         ║${NC}"
echo -e "${GREEN}║                                                              ║${NC}"
echo -e "${GREEN}╚══════════════════════════════════════════════════════════════╝${NC}"
echo ""
echo -e "  ${BLUE}Location:${NC} $FINAL_PATH"
echo -e "  ${BLUE}Size:${NC}     $FILE_SIZE"
echo ""
echo -e "  ${YELLOW}This is a self-contained app that includes the .NET runtime.${NC}"
echo -e "  ${YELLOW}No additional installation required.${NC}"
echo ""

# macOS-specific: remind about Gatekeeper
if [ "$OS" = "Darwin" ]; then
    echo -e "  ${CYAN}Tip: If macOS shows 'cannot be opened' error:${NC}"
    echo -e "       Right-click → Open, or run:"
    echo -e "       ${BLUE}xattr -cr ~/Desktop/Playwrighter.app${NC}"
    echo ""
fi

echo -e "${GREEN}Done!${NC}"
