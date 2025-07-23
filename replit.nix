{ pkgs }: {
  deps = [
    pkgs.ghostscript # Noch für PyPDF2 nötig, falls nicht entfernt
    pkgs.python3Packages.tk
    pkgs.jdk # Für Java
  ];
}