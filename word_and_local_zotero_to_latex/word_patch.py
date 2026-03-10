from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from lxml import etree

import win32com.client as win32

from .models import CitationField
from .placeholders import make_citation_placeholder
from .zotero_field import parse_zotero_code
from .mendeley_field import parse_mendeley_code

class WordCitationPatcher:
    """
    Replace Zotero citation fields with stable placeholders inside a Word document.

    This enables pandoc conversion without losing citation linkage.
    """

    def __init__(self, *, word_path: Path, output_docx_path: Path) -> None:
        assert word_path.exists()
        self.word_path = word_path
        self.output_docx_path = output_docx_path

    def patch_to_placeholders(self) -> dict[str, list[CitationField]]:
        """
        Patch the Word document and return mapping: placeholder -> CitationField list.
        """
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = None

        fields: dict[str, list[CitationField]] = defaultdict(list)
        title_list: list[str] = []

        try:
            doc = word.Documents.Open(str(self.word_path.absolute()))
            # Reverse iteration is important when modifying fields in-place.
            
            #Loop through Zotero fields
            for i in range(doc.Fields.Count, 0, -1):
                f = doc.Fields.Item(i)
                raw_code = f.Code.Text
                parsed = parse_zotero_code(raw_code)
                if not parsed:
                    continue
                
                print("Found Zotero reference") 
                
                citation_id = parsed.get("citationID")
                if not isinstance(citation_id, str) or not citation_id:
                    continue

                placeholder = make_citation_placeholder(citation_id)

                for citation_item in parsed.get("citationItems", []) or []:
                    if not isinstance(citation_item, dict):
                        continue
                    item_data = citation_item.get("itemData")
                    if not isinstance(item_data, dict):
                        continue
                    title = item_data.get("title")
                    if isinstance(title, str) and title:
                        title_list.append(title)

                    fields[placeholder].append(
                        CitationField(
                            payload_raw=raw_code,
                            payload=item_data,  # type: ignore[arg-type]
                        )
                    )

                # Replace visible content with placeholder.
                f.Result.Text = placeholder

                # Normalize formatting to avoid pandoc quirks.
                f.Result.Font.Superscript = False
                f.Result.Font.Subscript = False
                f.Result.Font.Bold = False
                f.Result.Font.Italic = False
                f.Result.Font.Underline = False
                f.Result.Font.StrikeThrough = False
                
            doc.SaveAs(str(self.output_docx_path.absolute()))
        
        finally:
            if doc is not None:
                # Do not prompt for saving changes again.
                doc.Close(False)
            # Ensure Word process is always cleaned up.
            word.Quit()
            
        #Loop through Mendeley fields
        try:
            docXML = Document(str(self.output_docx_path.absolute()))
            
            for sdt in docXML.element.xpath('//w:sdt'):
                tag = sdt.find('.//w:tag', docXML.element.nsmap)
                raw_code = tag.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val','')
                parsed = parse_mendeley_code(raw_code)
                if not parsed:
                    continue
                
                print("Found Mendeley reference") 
                
                # Find the index of the sdt inside its parent to know where to enter the text
                parent = sdt.getparent()
                while parent.tag != '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
                    parent = parent.getparent()
                citation_id = parsed.get("citationID")
                
                # Define placeholder
                placeholder = make_citation_placeholder(citation_id)
                new_run = OxmlElement('w:r')
                new_text = OxmlElement('w:t')
                new_text.text = placeholder
                new_run.append(new_text)
                
                # Replace visible content with placeholder.
                parent.remove(sdt)
                parent.append(new_run)

                docXML.save(str(self.output_docx_path.absolute()))
                
                #Save into dictionary of fields
                for citation_item in parsed.get("citationItems", []) or []:
                    if not isinstance(citation_item, dict):
                        continue
                    item_data = citation_item.get("itemData")
                    if not isinstance(item_data, dict):
                        continue
                    title = item_data.get("title")
                    if isinstance(title, str) and title:
                        title_list.append(title)

                    fields[placeholder].append(
                        CitationField(
                            payload_raw=raw_code,
                            payload=item_data,  # type: ignore[arg-type]
                        )
                    )
                    
        except Exception as e:
            print("No word file:", e)

        print(f"Found {len(title_list)} refs")
        print(f"Found {len(set(title_list))} unique titles")
        return fields

