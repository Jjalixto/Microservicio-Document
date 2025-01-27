from abc import ABC, abstractmethod
from models.document_request import DocumentRequest

class ContractStrategy(ABC):
    """
    Interfaz para las estrategias de generación de contratos.
    Cada estrategia debe implementar el método generate_contract.
    """
    
    @abstractmethod
    def process_request(self,request: DocumentRequest):
        pass