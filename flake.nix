{
  description = "Teams CLI - Playwright automation for Microsoft Teams";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-26.05";
    # Pinned to the nixpkgs revision (from nixos-unstable) whose playwright-driver
    # is 1.64.1, so the bundled browsers match the npm @playwright/test@1.64.1 in
    # package.json. Keep in sync with package.json.
    nixpkgs-playwright.url = "github:NixOS/nixpkgs/2786a2a72088ecf32aea5b3ec52099e6a7cff78f";
  };

  outputs = { self, nixpkgs, nixpkgs-playwright }:
    let
      system = "x86_64-linux";
      pkgs = import nixpkgs { inherit system; };
      pkgsPlaywright = import nixpkgs-playwright { inherit system; };
      # Only Chromium is used. WebKit is left out because it isn't in the binary
      # cache, so including it means a large local build.
      playwrightBrowsers = pkgsPlaywright.playwright-driver.selectBrowsers {
        withWebkit = false;
      };
    in {

      devShells.${system} = {
        default = pkgs.mkShell {
          buildInputs = [
            pkgs.nodejs
            pkgs.pnpm
          ];

          shellHook = ''
            echo "Node.js version: $(node -v)"
            echo "pnpm version: $(pnpm -v)"
          '';
        };

        playwright = pkgs.mkShell {
          buildInputs = [
            pkgs.nodejs
            pkgs.pnpm
          ];
          nativeBuildInputs = [ playwrightBrowsers ];
          shellHook = ''
            export PLAYWRIGHT_BROWSERS_PATH=${playwrightBrowsers}
            export PLAYWRIGHT_SKIP_VALIDATE_HOST_REQUIREMENTS=true
          '';
        };
      };

    };
}
