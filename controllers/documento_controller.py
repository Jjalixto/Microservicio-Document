from fastapi import APIRouter, HTTPException
from models.document_request import DocumentRequest
from services.documento_service import DocumentoService

router = APIRouter()

@router.get("/saludo", status_code=200)
async def saludo():
    return {"message": "Hola mundo!"}

@router.post("/generate-document", status_code=201)
async def generate_document(request: DocumentRequest):
    try:
        documento_service = DocumentoService()
        result = documento_service.servicioCentral(request)
        return result
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        # Log para ayudar a la depuración
        print(f"Error inesperado: {str(e)}")
        raise HTTPException(status_code=500, detail="Error interno del servidor.")
