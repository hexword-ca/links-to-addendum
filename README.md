# Links-to-Addendum

This is a python script to convert all hyperlinks in a *.docx file into a list of clickable hyperlinks addendum section. This script leaves the original file unchanged, and creates a copy of it with an additional section added at the end. The added section at the end
of the new file conatins a Indexed and clickable URL Text links. In addition if source document used Cross-reference Links in the style [1], [1, 2], or (1, 2, 3) - then these numeric refences are updated in the copy to reflect the Index number of the link in the Addendum section.


## Description

Hyperlink to Addendum Converter for DOCX files (Refined Version).
- Extracts all external hyperlinks.
- If an inline hyperlink's anchor text is a full URL (starts with http://, https://, or www.), OR if it conatins anchor text
 other then "()[]:.,0123456789 " it keeps the URL text intact in the body.
- If an inline hyperlink's anchor text is custom text/phrases, it replaces it with the citation index [idx].
- Appends an Addendum section with clickable URLs corresponding to the citation index.


## Installation of Requirements On Debian/Ubuntu

```
sudo apt install python3 python3-pip python3-docx -y
```

## Usage

```
python3 links_to_addendum.py <FileName.docx>
```

## License

MIT.
