# VS Code Extensions Installer
# Installs all extensions discussed in your setup guide.
# Safe to re-run; VS Code will skip already-installed extensions.

$extensions = @(
    # Vim Layer
    "vscodevim.vim",
    "VSpaceCode.whichkey"   # optional

    # Core Extensions
    "EditorConfig.EditorConfig",
    "usernamehw.errorlens",
    "eamodio.gitlens",
    "GitHub.copilot",
    "GitHub.copilot-chat",
    "GitHub.vscode-pull-request-github",

    # C++
    "ms-vscode.cpptools",
    "xaver.clang-format",

    # Python
    "ms-python.python",
    "ms-python.vscode-pylance",
    "charliermarsh.ruff",

    # Perl
    "richterger.perl",

    # Bash
    "mads-hartmann.bash-ide-vscode",
    "timonwong.shellcheck",

    # PowerShell
    "ms-vscode.PowerShell",

    # Quality of Life
    "christian-kohler.path-intellisense",
    "alefragnani.Bookmarks",
    "Gruntfuggly.todo-tree",
    "aaron-bond.better-comments",
    "streetsidesoftware.code-spell-checker"
)

Write-Host "Installing VS Code extensions..." -ForegroundColor Cyan

foreach ($ext in $extensions) {
    Write-Host "→ Installing $ext" -ForegroundColor Yellow
    code --install-extension $ext --force
}

Write-Host "`nAll extensions processed." -ForegroundColor Green
