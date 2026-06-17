import sys
import json
from pathlib import Path

sys.path.append(str(Path(__file__).parent))
from mcp_server_3 import load_rag_settings, update_rag_settings_tool

print("\n" + "="*80)
print("ТЕКУЩИЕ НАСТРОЙКИ RAG")
print("="*80)
current_settings = load_rag_settings()
print(json.dumps(current_settings, ensure_ascii=False, indent=2))
print("\n" + "="*80)
print("ИЗМЕНЯЕМ НАСТРОЙКИ")
print("="*80)

# Меняем chunk_size, chunk_overlap и top_k
result = update_rag_settings_tool(
    chunk_size=1500,
    chunk_overlap=200,
    top_k=7
)

print(result)

print("\n" + "="*80)
print("НОВЫЕ НАСТРОЙКИ ПОСЛЕ ИЗМЕНЕНИЯ")
print("="*80)
new_settings = load_rag_settings()
print(json.dumps(new_settings, ensure_ascii=False, indent=2))