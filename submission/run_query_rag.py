import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent))
from mcp_server_3 import query_rag_tool

result = query_rag_tool(
    question="Галилео",
    collection_name="word_standards",
    top_k=3,
    include_sources=True
)

print("\n" + "="*80)
print("РЕЗУЛЬТАТ ЗАПРОСА:")
print("="*80)
print(result)