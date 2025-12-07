# epub-to-cbz-pwsh

Extracts images from EPUB comic files and reorganizes them into properly-ordered CBZ archives. The script parses EPUB metadata to determine correct page sequence and eliminates duplicate cover pages that may appear in the content spine.

## Key Features

- **Smart Image Extraction**: Parses XHTML files within the EPUB and extracts image references from both standard `<img>` tags and SVG `<image>` elements
- **Spine-Based Ordering**: Uses the EPUB manifest and spine to determine correct page order, respecting the publisher's intended reading sequence
- **Duplicate Detection**: Uses MD5 file hashing to detect and skip duplicate first pages that match the cover image
- **Metadata Preservation**: Automatically generates ComicInfo.xml with extracted title, publisher, creator, publication date, and other metadata from the EPUB

## Tested Sources

This tool has been tested with EPUB files from the following publishers and platforms:

- **Image Comics** (Humble Bundle)
- **Kodansha** (via Humble Bundle & Kobo)
- **Yen Press** (via Kobo)
- **Seven Seas Entertainment** (via Kobo)

## How It Works

The script:
1. **Extracts the EPUB**: Unzips the EPUB file to access its internal structure
2. **Parses Metadata**: Reads the OPF (Open Packaging Format) file to get the spine order and manifest information
3. **Extracts Images**: Finds all XHTML content files and uses regex to extract image references from both standard HTML `<img>` tags and SVG `<image>` elements
4. **Detects Duplicates**: Compares the MD5 hash of the first content page against the extracted cover image to skip duplicate covers (common when publishers include cover in both cover and content sections)
5. **Orders Pages**: Renames extracted images with sequential numbering based on spine order to ensure correct reading sequence
6. **Generates Metadata**: Creates ComicInfo.xml with EPUB metadata for better comic reader integration
7. **Creates CBZ**: Packages all ordered images into a standard CBZ format

## Important Notes

⚠️ **EPUB files are complex.** Each publisher packages their EPUB files differently with varying internal structures, image compression methods, and metadata formats. Due to this variability:

- **Always test with a small sample** of EPUB files from each publisher before processing your entire library
- **Backup your original EPUB files** before running the conversion - do not delete your EPUBs immediately after conversion
- **Verify the output CBZ** in a comic reader to ensure all pages are in correct order and quality is acceptable
- Some EPUBs may not convert perfectly depending on their internal structure
- The script includes a fallback image discovery method if spine-based extraction fails

## Usage

```powershell
.\ComicEpubToCBZ.ps1 -Path "path\to\comic.epub"
```

Alternatively, add the batch wrapper to your PATH and edit the path to your PowerShell command:

```cmd
ComicEpubToCBZ.bat "path\to\comic.epub"
```

Or drag and drop EPUB files directly onto the script.

## Output

The CBZ file will be created in the same directory as the source EPUB with the same base filename. A ComicInfo.xml file is embedded with metadata to enhance compatibility with comic readers like Komga.
