#!/bin/bash
# Render build script for .NET Mail Merge SaaS

echo "Starting .NET build process..."

# Install .NET 8.0 SDK if not present
if ! command -v dotnet &> /dev/null; then
    echo "Installing .NET 8.0 SDK..."
    wget https://packages.microsoft.com/config/ubuntu/20.04/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
    sudo dpkg -i packages-microsoft-prod.deb
    sudo apt-get update
    sudo apt-get install -y apt-transport-https && sudo apt-get update && sudo apt-get install -y dotnet-sdk-8.0
fi

# Restore packages and build
echo "Restoring NuGet packages..."
dotnet restore

echo "Building application..."
dotnet publish -c Release -o ./publish --self-contained false

echo "Build completed successfully!"
echo "Application will be served from ./publish/MailMergeSaaS.dll"