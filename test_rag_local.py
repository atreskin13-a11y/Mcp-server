from pathlib import Path
from mcp_server_3 import build_rag_tool, query_rag_tool, update_rag_settings_tool, read_docx

doc_path = Path(r"C:\Users\Олегсандр\Desktop\rag_test\test.docx")
print("=== READ DOCX ===")
print(read_docx(doc_path))

print("=== UPDATE SETTINGS ===")
print(update_rag_settings_tool(
    chunk_size=800,
    chunk_overlap=100,
    top_k=3,
    include_exts={".py", ".md", ".txt", ".docx"},
    exclude_dirs={"__pycache__", ".git", "venv", "env", ".idea", ".mypy_cache", ".pytest_cache", "chroma_storage"},
))

print("\n=== BUILD RAG ===")
print(build_rag_tool(
    root_dir=r"C:\Users\Олегсандр\Desktop\rag_test",
    clear_collection=True,
    max_files=10,
    max_file_chars=50000,
    batch_size=20,
))

print("\n=== QUERY RAG ===")
print(query_rag_tool(
    question="Кубик мышка стол блютуз смайлик",
    top_k=1,
    include_sources=True,
))