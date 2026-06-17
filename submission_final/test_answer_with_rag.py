import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent))
from mcp_server_3 import answer_with_rag_tool

# Тест: отвечает ли RAG на вопрос
result = answer_with_rag_tool(
    question="Что такое тестовый документ?",
    collection_name="word_standards",
    top_k=3
)

print("\n" + "="*80)
print("РЕЗУЛЬТАТ: Ответ на вопрос по собранной информации")
print("="*80)
print(result)