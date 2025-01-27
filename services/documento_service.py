from models.document_request import DocumentRequest
from strategies.one_owner import OneOwner
from strategies.two_owners import TwoOwners
from strategies.three_owners import ThreeOwners

class DocumentoService:
    
    """
    Servicio central para manejar la generación de documentos.
    """
    def servicioCentral(self, request: DocumentRequest):
        
        # Validación del número de propietarios
        if request.propietario not in [1, 2, 3]:
            raise ValueError("El número de propietarios debe ser 1, 2 o 3.")
        # Selección de la estrategia adecuada según el número de propietarios
        
        if request.propietario == 1:
            strategy = OneOwner()
        elif request.propietario == 2:
            strategy = TwoOwners()  
        elif request.propietario == 3:
            strategy = ThreeOwners()
        
        # Procesar la solicitud con la estrategia seleccionada
        return strategy.process_request(request)

