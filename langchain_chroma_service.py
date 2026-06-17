from view.utils import *
from view.utils_assistant  import list_assistants, list_name_assistants
import requests
from langchain_community.document_loaders import (
    PyPDFLoader,
    Docx2txtLoader,
    UnstructuredExcelLoader,
    UnstructuredFileLoader,  # для .doc (устаревшие)
)

from langchain_chroma import Chroma
from langchain_text_splitters import RecursiveCharacterTextSplitter
from sentence_transformers import SentenceTransformer
from langchain_core.embeddings import Embeddings

MODEL_NAME = r"F:\IdeaProjects20241106\2\llm_models\LaBSE"

class CustomSentenceTransformerEmbeddings(Embeddings):
    def __init__(self, model_name: str):
        self.name = model_name
        self.model = SentenceTransformer(model_name)

    def embed_documents(self, texts: list[str]) -> list[list[float]]:
        return self.model.encode(texts, convert_to_tensor=False).tolist()

    def embed_query(self, text: str) -> list[float]:
        return self.model.encode([text], convert_to_tensor=False)[0].tolist()

class ChromaService:
    @staticmethod
    def read_doc(save_path):
        # Загружаем документ
        ext = Path(save_path).suffix.lower()
        if ext == ".pdf":
            loader = PyPDFLoader(str(save_path))
        elif ext == ".docx":
            loader = Docx2txtLoader(str(save_path))
        elif ext in [".xls", ".xlsx"]:
            loader = UnstructuredExcelLoader(str(save_path))
        elif ext == ".doc":  # старый формат
            loader = UnstructuredFileLoader(str(save_path), mode="elements")
        else:
            raise ValueError("Файл не поддерживается. Используйте .pdf, .docx, .xlsx")

        doc = loader.load()
        return doc

    @staticmethod
    def get_assistant_id(assistant_name):
        if not isinstance(assistant_name, str):  # for gradio
            print(list_assistants())
            d = {name:id for id, name in list_assistants()}
            assistant_id = d[assistant_name[0]['value']]

        else: # for streamlit
            print(list_assistants())
            d = {name:id for id, name in list_assistants()}

            assistant_id = d.get(assistant_name, None)
        return assistant_id

    @staticmethod
    def make_directory_for_assistant(name):
        config = load_config()
        # Генерируем уникальный ID
        new_id = f"assistant_{len(config['assistants']) + 1}"
        config["assistants"].append({"id": new_id, "name": name})
        save_config(config)
        persist_directory = os.path.join(ASSISTANTS_DIR, new_id, "vector_store")
        # Создаём папку
        os.makedirs(persist_directory, exist_ok=True)
        return persist_directory

    @staticmethod
    def generate_answer(prompt):
        try:
            response = requests.post(
                "http://localhost:1234/v1/chat/completions",
                json={
                    "model": "local-model",
                    "messages": [
                        {"role": "system", "content": "Ты — полезный и точный ассистент. Отвечай на русском/английском языке, как в вопросе."},
                        {"role": "user", "content": prompt}
                    ],
                    "temperature": 0.3,
                },
                timeout=600
            )
            response.raise_for_status()
            answer = response.json()["choices"][0]["message"]["content"]
            return answer

        except requests.exceptions.ConnectionError:
            return "⚠️ LM Studio не запущен или недоступен по адресу http://localhost:1234. Запустите LM Studio → вкладка *Local Server* → Start Server."
        except Exception as e:
            return  f"❌ Ошибка при запросе к LM Studio: {e}"

    def __init__(self, assistant_name, model_name=None):
        self.assistant_name = assistant_name

        if not assistant_name in list_name_assistants():
            self.persist_directory = self.make_directory_for_assistant(assistant_name)
            self.assistant_id = self.get_assistant_id(assistant_name)
        else:
            self.assistant_id = self.get_assistant_id(assistant_name)
            self.persist_directory =  os.path.join(ASSISTANTS_DIR, self.assistant_id, "vector_store")
        print(f"{self.assistant_name=}, {self.assistant_id=}, {self.persist_directory=}")
        # Создаём embeder
        model_name = model_name or MODEL_NAME
        self.model = CustomSentenceTransformerEmbeddings(model_name)
        self.db = Chroma(persist_directory=self.persist_directory, embedding_function=self.model)

    def make_content(self, question):
        docs =self.db.search(question, 'similarity', k=5)
        context = "\n\n".join([d.page_content for d in docs])
        return context

    def make_answer_llm(self, question):
        context = self.make_content(question)
        prompt = f"""Ответь на вопрос, используя предоставленную информацию.

        Контекст:
        {context}
        
        Вопрос: {question}"""

        return self.generate_answer(prompt)


    def add_docs(self, doc):
        splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=100)
        chunks = splitter.split_documents(doc)
        for chunk in chunks:
            print(chunk)
            print("*******************************************")
        # Сохраняем в ChromaDB текущего ассистента
        ids = [f"{i}" for i in range(len(chunks))]
        metadatas = [d.metadata for d in chunks]

        texts = [d.page_content for d in chunks]
        self.db.add_texts(texts=texts, metadatas=metadatas, ids=ids)


    # --- Загрузка документов в БД ---
    def process_and_store_file(self, file_obj):
        try:
            if isinstance(file_obj, str) and os.path.exists(file_obj):
                doc = self.read_doc(file_obj)
                self.add_docs(doc)
                return f"✅ Файл «{file_obj.name}» успешно обработан и добавлен в базу ассистента {self.assistant_name=}."
            else:
                tempfile = save_temp_file(file_obj)
                doc = self.read_doc(tempfile)
                self.add_docs(doc)
                return f"✅ Файл «{file_obj.name}» успешно обработан и добавлен в базу ассистента {self.assistant_name=}."
        except Exception as e:
            return f"❌ {file_obj.name=}, {self.assistant_name=}, {self.assistant_id=}. Ошибка: {e}"