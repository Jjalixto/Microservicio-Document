from pydantic import BaseModel
from typing import Optional

class DocumentRequest(BaseModel):

    #primero evaluar si es casado viudo o divorciado
    marital_status: str
    
    #Definimos cuantos propietarios tendra el contrato
    propietario:str
    
    #Definimos el tipo de contrato
    condicion: str
    
    #Datos para el contrato para el word
    name_1: Optional[str] = ''
    dni_1: Optional[str] = ''
    ocupation_1: Optional[str] = ''
    marital_status_1: Optional[str] = ''
    address_1: Optional[str] = ''
    mail_1: Optional[str] = ''
    phone_1: Optional[str] = ''
    
    name_2: Optional[str] = ''
    dni_2: Optional[str] = ''
    ocupation_2: Optional[str] = ''
    marital_status_2: Optional[str] = ''
    address_2: Optional[str] = ''
    mail_2: Optional[str] = ''
    phone_2: Optional[str] = ''
    
    number_batch: Optional[str] = ''
    approximate_area: Optional[str] = ''
    
    #Fecha de la firma del contrato
    day: Optional[str] = ''
    month: Optional[str] = ''
    year: Optional[str] = ''
    
    #Datos para el excel
    monto_venta: Optional[str] = ''
    monto_letras: Optional[str] = ''
    monto_reserva: Optional[str] = ''
    reserva_letras: Optional[str] = ''
    cuota_inicial: Optional[str] = ''
    cuo_init_letras: Optional[str] = ''
    cantidad_anios: Optional[str] = ''
    fecha_primera_cuota: Optional[str] = ''
    saldo_restante: Optional[str] = ''
    saldo_restante_letras: Optional[str] = ''
    day_c: Optional[str] = ''
    month_c: Optional[str] = ''
    year_c: Optional[str] = ''
    
    #Datos para cuadros word
    precio_letras: Optional[str] = ''
    cuota_inicial_letras: Optional[str] = ''