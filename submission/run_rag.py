from pathlib import Path
from mcp_server_3 import build_rag_tool

result = build_rag_tool(
    root_dir=r"C:\Users\Олегсандр\Desktop\rag_test",
    max_files=200,
    collection_name="word_standards",
    include_exts=[".docx"],
    clear_collection=True
)

print(result)