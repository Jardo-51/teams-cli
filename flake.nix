{
  description = "Teams CLI - Playwright automation for Microsoft Teams";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-26.05";
    # Pinned to the nixpkgs revision whose playwright-driver is 1.56.1, so the
    # bundled browsers match the npm @playwright/test@1.56.1 in package.json.
    nixpkgs-playwright.url = "github:NixOS/nixpkgs/95506c9ea100652793b5c2a893d7d9ef182db731";
  };

  outputs = { self, nixpkgs, nixpkgs-playwright }:
    let
      system = "x86_64-linux";
      pkgs = import nixpkgs { inherit system; };
      pkgsPlaywright = import nixpkgs-playwright { inherit system; };
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
          nativeBuildInputs = [ pkgsPlaywright.playwright-driver.browsers ];
          shellHook = ''
            export PLAYWRIGHT_BROWSERS_PATH=${pkgsPlaywright.playwright-driver.browsers}
            export PLAYWRIGHT_SKIP_VALIDATE_HOST_REQUIREMENTS=true
          '';
        };
      };

    };
}
