import chromadb
from chromadb.utils import embedding_functions as EF


class VectorDBWorker:
    client = None

    @staticmethod
    def set_client(path):
        VectorDBWorker.client = chromadb.PersistentClient(path)

    @staticmethod
    def reset():
        VectorDBWorker.client.reset()

    def __init__(self, name_collection:str, embedding_function=None):
        self.name_collection = name_collection
        self.emb_function = embedding_function if not embedding_function is None else EF.DefaultEmbeddingFunction()

        self.collection = self.client.get_or_create_collection(name=name_collection)
        self.collection._embedding_function = self.emb_function


    def add_texts(self, texts:list, ids:list, metadatas:list):
        if len(texts) != len(ids) or len(texts) != len(metadatas):
            raise ValueError("texts, ids and metadatas must have the same length")
        try:
            self.collection.add(documents=texts, ids=ids, metadatas=metadatas)
            print(f"--- done metadatas: {metadatas[0]}, OK")
        except Exception as e:
            print(f"Error add_texts: {e}, texts={texts}, ids={ids}, metadatas={metadatas}")
            raise MemoryError(f"Error add_texts: {e}")

    async def add_embeddings(self, embeddings:list, ids:list, metadatas:list, texts:list = list()):
        if len(embeddings) != len(ids) or len(embeddings) != len(metadatas):
            raise ValueError("texts, ids and metadatas must have the same length", len(embeddings), len(ids), len(metadatas))
        self.collection.add(embeddings=embeddings, ids=ids, metadatas=metadatas, documents=texts)

    def get_all_texts(self):
        # Извлечение всех документов из коллекции
        all_documents = self.collection.get()  # Получаем все данные из коллекции
        # Список всех текстов документов
        texts = all_documents["documents"]
        return [text.lower() for text in texts]


    def get_documents(self, query:str, n_results:int):
        query_embedding = self.emb_function(query)  # List[float]

        results = self.collection.query(query_embeddings=[query_embedding], n_results=n_results)
        return results['documents'][0]

    def get_ids(self, query:str, n_results:int):
        results = self.collection.query(query=query, n_results=n_results)
        return results.ids


    def get_answer(self, query:str, n_results:int=1, where_document=None):
        if where_document is None or len(where_document) == 0:
            results = self.collection.query(query_texts=[query], n_results=n_results)
        else:
            results = self.collection.query(query_texts=[query], n_results=n_results, where_document=where_document)
        return results