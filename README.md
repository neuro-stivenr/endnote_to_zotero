# endnote_to_zotero

Simple commandline script to facilitate the conversion of Word documents with EndNote citations into LibreOffice documents with Zotero citations.

Currently only works on Mac, but it should be fairly straightforward to make it work on other platforms.

## System Dependencies

1. EndNote + Microsoft Word Cite-While-You-Write Plugin
2. Zotero + LibreOffice Plugin
3. [LibreOffice](https://www.libreoffice.org/download/download-libreoffice)
4. [Zotero ODF Scan Plugin](https://zotero-odf-scan.github.io/zotero-odf-scan)
5. You might need to install [Java](https://www.oracle.com/java/technologies/downloads) for the LibreOffice Zotero plugin.
6. Python 3 w/ pip

## Installation

Install the Python package:
```bash
pip install git+https://github.com/neuro-stivenr/endnote_to_zotero
```

Register an online Zotero account, and create an application API key.

Add the following lines to your .bashrc or .zshrc (depending on the shell you use):
```bash
export ZOTERO_UID=<ZOTERO USER API KEY>
export ZOTERO_API_KEY=<ZOTERO APPLICATION API KEY>
```

Make sure that in LibreOffice -> Preferences -> Advanced
Java Runtime Environment is installed.

Double click on endnote_to_zotero/endnote/RTF.ens and then "Save As..." the style into your EndNote.

Make sure that in your Zotero -> Preferences -> Export -> Item Format
the "Scannable Cite" option is selected. If it is unavailable, you must download "Scannable Cite.js" file from
the [Zotero ODF Scan Plugin](https://zotero-odf-scan.github.io/zotero-odf-scan) page and save it to your $HOME/Zotero/translators directory.
Make sure the file is saved with a .js and not a .txt extension.

## Usage

Our starting point is a Word document with EndNote citations.

1. Export travelling bibliography from the document using BibTex format.
2. Import it into Zotero as a new library and sync your local Zotero libraries with the cloud.
    * Thus far, the program was only tested with the default record number citekey export option.
3. Change the citation style to the "RTF" style that you previously installed into EndNote.
4. Inside Microsoft Word click into "Tools" -> "Macro" -> "Macros..." and run the ENWRemoveFieldCodes macro.
5. This will open a new document where all the field codes were converted to plain text. Save it as a .docx file.
6. On the commandline, run endnote_to_zotero command and follow the prompt.
7. Run the ODF scan Zotero plugin on the .odt output file generated by endnote_to_zotero
8. Open the output generated by the ODF scan plugin in LibreOffice
9. Click on "Document Preferences" or "Refresh Citations" to update the Zotero citations.
10. The resulting .odt output file can either be opened in Microsoft Word or uploaded to Google Docs directly as needed.

## Under the Hood

[This forum post](https://forums.zotero.org/discussion/34233/almost-convert-citations-from-endnote-to-zotero-in-word-docx) was my starting point for getting this to work. The script applies a similar set of principles as described there.

1. A custom EndNote style is used such that in-text citations are formatted to be compatible with the Zotero ODF scan plugin.
2. After the Word document is stripped from field codes, it is saved, and converted to .odt and then to .html via LibreOffice commands. Not sure why I have to convert to the intermediate .odt format to get this working, might have something to do with how .docx and .odt are converted to .html by LibreOffice.
3. The endnote_to_zotero.py script uses pyzotero along with user credentials to query the item IDs pertaining to the document library initially imported to Zotero.
4. Item IDs are linked to in-text citations via the citekey, and the relevant in-text citations are reformatted to conform to ODF scan Zotero plugin.
5. Resulting .html file is converted back to .odt, and ODF scan Zotero plugin is applied to the document.
6. After opening the output of that in LibreOffice, the recognized citations are converted to user-defined Zotero style and unrecognized citations are brought to the user's attention.


