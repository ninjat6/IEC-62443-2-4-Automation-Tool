# core/analyzer.py
import os
import json
from sentence_transformers import SentenceTransformer, util
from conformity_analysis_module.core.file_processor import FileProcessor
from conformity_analysis_module.utils.logger import logger
from conformity_analysis_module.config import Config

class Analyzer:
    def __init__(self, model_name='models/all-MiniLM-L12-v2'):
        self.model = SentenceTransformer(model_name)

    def analyze(self, folder_path, requirements, threshold=0.5):
        results = []
        requirement_embeddings = {}
        for req_key, req_text in requirements.items():
            try:
                requirement_embeddings[req_key] = self.model.encode(req_text, convert_to_tensor=True)
            except Exception as e:
                logger.error(f"計算條款 {req_key} 向量時發生錯誤: {e}")

        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                ext = file.lower().split('.')[-1]
                if ext not in ['docx', 'xlsx', 'pdf']:
                    continue
                logger.info(f"處理檔案: {file_path}")
                snippets = FileProcessor.extract_text_snippets(file_path)
                if not snippets:
                    continue

                try:
                    snippet_embeddings = self.model.encode(snippets, convert_to_tensor=True)
                except Exception as e:
                    logger.error(f"計算檔案 {file_path} 中片段向量時發生錯誤: {e}")
                    continue

                for req_key, req_embedding in requirement_embeddings.items():
                    try:
                        cosine_scores = util.cos_sim(req_embedding, snippet_embeddings)[0].cpu().numpy()
                    except Exception as e:
                        logger.error(f"計算相似度時發生錯誤: {e}")
                        continue

                    for idx, score in enumerate(cosine_scores):
                        if score >= threshold:
                            results.append({
                                "requirement": req_key,
                                "requirement_text": requirements[req_key],
                                "snippet": snippets[idx],
                                "similarity": float(score),
                                "source_file": file_path
                            })

        with open(Config.ANALYSIS_OUTPUT, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=4)
        logger.info(f"分析完成，共找到 {len(results)} 筆符合結果")
        return results