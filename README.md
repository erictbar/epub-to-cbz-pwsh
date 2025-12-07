# epub-to-cbz-pwsh

Extracts images from a comic in epub format and turns it into a cbz.

## Tested Sources

This tool has been tested with EPUB files from the following publishers and platforms:

- **Image Comics** (Humble Bundle)
- **Kodansha** (via Humble Bundle & Kobo)
- **Yen Press** (via Kobo)
- **Seven Seas Entertainment** (via Kobo)

## How It Works

The script uses hash-based detection to determine if the cover image needs to be renamed. It calculates a hash of each extracted image and compares it against known patterns to identify cover pages, ensuring proper naming and ordering in the final CBZ file.

## Important Notes

⚠️ **EPUB files are complex.** Each publisher packages their EPUB files differently with varying internal structures, image compression methods, and metadata formats. Due to this variability:

- **Always test with a small sample** of EPUB files from each publisher before processing your entire library
- **Backup your original EPUB files** before running the conversion - do not delete your EPUBs immediately after conversion
- **Verify the output CBZ** in a comic reader to ensure all pages are in correct order and quality is acceptable
- Some EPUBs may not convert perfectly depending on their internal structure

## Usage

```powershell
.\ComicEpubToCBZ.ps1 -Path "path\to\comic.epub"
```

Alternatively, add the batch wrapper to your PATH and edit the path to your PowerShell command:

```cmd
ComicEpubToCBZ.bat "path\to\comic.epub"
```

## Output

The CBZ file will be created in the same directory as the source EPUB with the same base filename.
