if ($IsMacOS) {
    $brewCheck = Get-Command brew -ErrorAction SilentlyContinue
    if (-not $brewCheck) {
        Write-Error "Homebrew is not installed. Please install Homebrew and try again."
    } else {
        # Install mas (Mac App Store Command Line Interface)
        brew install mas

        # Check if Xcode is installed
        $xcodeCheck = mas list | Select-String "497799835"
        if (-not $xcodeCheck) {
            # Install Xcode
            mas install 497799835
        } else {
            Write-Host "Xcode is already installed."
        }

        # Now that Xcode is ensured to be installed, install mono-libgdiplus
        brew install mono-libgdiplus
    }
} else {
    Write-Error "This script is intended to be run on a macOS system."
}