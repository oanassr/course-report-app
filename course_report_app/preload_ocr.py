"""Download and validate OCR models during deployment, not during a user request."""

from rapidocr import EngineType, LangDet, LangRec, ModelType, OCRVersion, RapidOCR


RapidOCR(
    params={
        "Det.engine_type": EngineType.ONNXRUNTIME,
        "Det.lang_type": LangDet.CH,
        "Det.model_type": ModelType.MOBILE,
        "Det.ocr_version": OCRVersion.PPOCRV5,
        "Rec.engine_type": EngineType.ONNXRUNTIME,
        "Rec.lang_type": LangRec.ARABIC,
        "Rec.model_type": ModelType.MOBILE,
        "Rec.ocr_version": OCRVersion.PPOCRV5,
    }
)
print("OCR models ready")
