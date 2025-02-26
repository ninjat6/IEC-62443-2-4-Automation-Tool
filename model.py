from sentence_transformers import SentenceTransformer

# 下載模型並存到指定資料夾，這裡假設你要儲存到 models/all-MiniLM-L12-v2
model = SentenceTransformer('sentence-transformers/all-MiniLM-L12-v2')
model.save('models/all-MiniLM-L12-v2')
