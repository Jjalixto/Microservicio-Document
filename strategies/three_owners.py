from .contract_strategy import ContractStrategy
from models.document_request import DocumentRequest

class ThreeOwners(ContractStrategy):
    """
    Estrategia para generar contratos para tres propietarios.
    """
    def process_request(self, request: DocumentRequest):
        return {"message": "Contrato generado para un 3 propietarios."}