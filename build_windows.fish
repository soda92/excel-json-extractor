#!/usr/bin/env fish

echo "Building for Windows (amd64)..."

# in fish, we use 'env' to set variables for a single command execution
# or 'set -x' to export variables for the current shell scope.
# Using 'env' here for a cleaner one-off build.

env CGO_ENABLED=1 \
    CC=x86_64-w64-mingw32-gcc \
    GOOS=windows \
    GOARCH=amd64 \
    go build -o extractor.exe

if test $status -eq 0
    echo "Successfully built: extractor.exe"
else
    echo "Build failed."
    exit 1
end
