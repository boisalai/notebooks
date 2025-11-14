# The essence of RAG lies in Retrieval, not Generation

Most teams building RAG systems focus on the Generation layer, when 90% of the value and challenges are in Retrieval.

The author, with 4 years of experience in document retrieval systems, identifies the key challenges:

**Extraction**:
Documents vary enormously (scanned PDFs, complex layouts, tables, handwritten notes). Three effective strategies:
- Layout detection + OCR to preserve structure
- Vision-language OCR models that treat the document as an image
- Multi-vector image embeddings to retrieve visually similar content

**Centralization**:
Group all pipelines in a mono-repo with versioning, testing, and CI/CD deployment. Use clean architecture and orchestrate with tools like Airflow or Prefect.

**Define obsolescence**:
Work with stakeholders to determine when a document becomes outdated and how often to refresh the data.

**Evaluation**:
Measure performance with synthetic queries, Recall@k, Mean Reciprocal Rank, and integrate user feedback.

**Conclusion**: Production RAG systems often fail because teams polish the Generation while Retrieval fails silently.