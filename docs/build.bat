@echo off
echo 1. Download docfx from https://github.com/dotnet/docfx/releases
echo 2. Add path to the folder where dotfx is unzipped

@docfx metadata docfx.json
@docfx build docfx.json
