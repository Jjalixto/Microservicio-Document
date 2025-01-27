from .contract_strategy import ContractStrategy
from models.document_request import DocumentRequest

class OneOwner(ContractStrategy):
    """
    Estrategia para generar contratos para un único propietario.
    """
    def process_request(self, request:DocumentRequest):
        return {"message": "Contrato generado para un único propietario."}