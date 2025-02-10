from models.document_request import DocumentRequest
from strategies.married.two_owners import TwoOwners as MarriedTwoOwners
from strategies.single.one_owner import OneOwner as SingleOneOwner
from strategies.single.two_owners import TwoOwners as SingleTwoOwners

class DocumentoService:
    
    """
    Servicio central para manejar la generación de documentos.
    """
    def servicioCentral(self, request: DocumentRequest):
        
        if request.marital_status not in ["casado", "soltero"]:
            raise ValueError("El estado civil debe ser 'casado' o 'soltero'.")
        
        # Conversión y validación del número de propietarios
        try:
            request.propietario = int(request.propietario)
        except ValueError:
            raise ValueError("El número de propietarios debe ser un número entero.")

        if request.propietario not in [1, 2, 3]:
            raise ValueError("El número de propietarios debe ser 1, 2 o 3.")
        
        # Selección de la estrategia adecuada según el estado civil y el número de propietarios
        if request.marital_status == "casado" or int(request.propietario) == 2:
            strategy = MarriedTwoOwners()
        elif request.marital_status == "soltero" or int(request.propietario) == 1:
            strategy = SingleOneOwner()
        elif request.marital_status == "soltero" or int(request.propietario) == 2:
            strategy = SingleTwoOwners()
        
        # Procesar la solicitud con la estrategia seleccionada
        return strategy.process_request(request)