from docling.document_converter import DocumentConverter
converter = DocumentConverter()
md = converter.convert('Aktif 31.03.2025.pdf').document.export_to_markdown()
print(md[:8000])

