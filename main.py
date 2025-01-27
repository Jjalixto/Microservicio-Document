from fastapi import FastAPI
from controllers.documento_controller import router

app = FastAPI(
    title="Document Generation API",
    description="API para generar contratos en Word y PDF según propietarios y condiciones",
    version="1.0.0"
)

app.include_router(router)