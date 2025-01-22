from fastapi import FastAPI
from controllers.documento_controller import router

app = FastAPI()

app.include_router(router)